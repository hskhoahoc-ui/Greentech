const express = require('express');
const ExcelJS = require('exceljs');
const supabase = require('./supabaseClient');
require('dotenv').config();

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));
app.set('view engine', 'ejs');

let isAdminLoggedIn = false;
let currentUsername = null;

// Trang chủ & Dữ liệu các bảng
app.get('/', async (req, res) => {
    try {
        const { data: categories } = await supabase.from('categories').select('*');
        const { data: services } = await supabase.from('services').select('*');
        const { data: projects } = await supabase.from('project_customers').select('*');
        const { data: quotes } = await supabase.from('quotes').select('*');
        const { data: expenses } = await supabase.from('expenses').select('*');
        const { data: products } = await supabase.from('products').select('*');
        const { data: exportTickets } = await supabase.from('export_tickets').select('*');

        const now = new Date();
        const updatedProjects = (projects || []).map(p => {
            const compDate = new Date(p.completion_date);
            const expiryDate = new Date(compDate.setMonth(compDate.getMonth() + p.warranty_months));
            const isExpired = now > expiryDate;
            return { ...p, isExpired, expiryDate: expiryDate.toISOString().split('T')[0] };
        });

        res.render('index', { 
            categories: categories || [], 
            services: services || [], 
            projects: updatedProjects,
            quotes: quotes || [],
            expenses: expenses || [],
            products: products || [],
            exportTickets: exportTickets || [],
            isAdmin: isAdminLoggedIn 
        });
    } catch (err) {
        console.error('Lỗi tải trang chủ:', err);
        res.status(500).send('Lỗi kết nối cơ sở dữ liệu!');
    }
});

// Xem sản phẩm theo danh mục
app.get('/category/:id', async (req, res) => {
    try {
        const categoryId = req.params.id;
        let { data: category } = await supabase.from('categories').select('*').eq('id', categoryId).maybeSingle();
        if (!category) {
            const { data: allCats } = await supabase.from('categories').select('*');
            category = (allCats || []).find(c => String(c.id) === String(categoryId)) || { id: categoryId, name: 'Danh mục sản phẩm' };
        }
        const { data: products } = await supabase.from('products').select('*').eq('category_id', categoryId);

        res.render('category', { category: category || {}, products: products || [], isAdmin: isAdminLoggedIn });
    } catch (err) {
        res.status(500).send('Lỗi tải danh mục!');
    }
});

// --- API ĐĂNG NHẬP / ĐĂNG XUẤT ---
app.post('/api/login', async (req, res) => {
    try {
        const { username, password } = req.body;
        const { data: users, error } = await supabase.from('app_users').select('*').eq('username', username).eq('password', password);
        if (error || !users || users.length === 0) {
            return res.json({ success: false, message: 'Sai tên đăng nhập hoặc mật khẩu!' });
        }
        isAdminLoggedIn = true;
        currentUsername = users[0].username;
        return res.json({ success: true });
    } catch (err) {
        return res.json({ success: false, message: 'Lỗi hệ thống!' });
    }
});

app.post('/api/logout', (req, res) => {
    isAdminLoggedIn = false;
    currentUsername = null;
    res.json({ success: true });
});

// --- API DANH MỤC SẢN PHẨM ---
app.post('/api/categories/save', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập quyền Admin!' });
    const { id, name, avatar_url } = req.body;
    let result = id 
        ? await supabase.from('categories').update({ name, avatar_url }).eq('id', id)
        : await supabase.from('categories').insert([{ name, avatar_url }]);
    if (result.error) return res.json({ success: false, message: result.error.message });
    res.json({ success: true });
});

app.post('/api/categories/delete', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập quyền Admin!' });
    const { id } = req.body;
    const { error } = await supabase.from('categories').delete().eq('id', id);
    if (error) return res.json({ success: false, message: error.message });
    res.json({ success: true });
});

// --- API SẢN PHẨM (THÊM / SỬA / XÓA) ---
app.post('/api/products/save', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập Admin!' });
    const { id, category_id, name, description, quantity, cost_price, selling_price, image_url } = req.body;
    let result = id
        ? await supabase.from('products').update({ name, description, quantity, cost_price, selling_price, image_url }).eq('id', id)
        : await supabase.from('products').insert([{ category_id, name, description, quantity, cost_price, selling_price, image_url }]);
    if (result.error) return res.json({ success: false, message: result.error.message });
    res.json({ success: true });
});

app.post('/api/products/delete', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập Admin!' });
    const { id } = req.body;
    const { error } = await supabase.from('products').delete().eq('id', id);
    if (error) return res.json({ success: false, message: error.message });
    res.json({ success: true });
});

// --- API DỰ ÁN KHÁCH HÀNG (THÊM / SỬA / XÓA) ---
app.post('/api/projects/save', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập!' });
    const { id, customer_name, project_name, completion_date, warranty_months } = req.body;
    let result = id
        ? await supabase.from('project_customers').update({ customer_name, project_name, completion_date, warranty_months }).eq('id', id)
        : await supabase.from('project_customers').insert([{ customer_name, project_name, completion_date, warranty_months }]);
    if (result.error) return res.json({ success: false, message: result.error.message });
    res.json({ success: true });
});

app.post('/api/projects/delete', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập!' });
    const { id } = req.body;
    const { error } = await supabase.from('project_customers').delete().eq('id', id);
    if (error) return res.json({ success: false, message: error.message });
    res.json({ success: true });
});

// --- API KHÁC (DỊCH VỤ, BÁO GIÁ, CHI PHÍ, XUẤT KHO) ---
app.post('/api/services/save', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập!' });
    const { service_type, content, warranty_maintenance, price } = req.body;
    const { error } = await supabase.from('services').insert([{ service_type, content, warranty_maintenance, price }]);
    if (error) return res.json({ success: false, message: error.message });
    res.json({ success: true });
});

app.post('/api/quotes/save', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập!' });
    const { quote_code, customer_name, project_name, address, total_amount } = req.body;
    const { error } = await supabase.from('quotes').insert([{ quote_code, customer_name, project_name, address, total_amount }]);
    if (error) return res.json({ success: false, message: error.message });
    res.json({ success: true });
});

app.post('/api/expenses/save', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập!' });
    const { expense_type, description, amount, expense_date } = req.body;
    const { error } = await supabase.from('expenses').insert([{ expense_type, description, amount, expense_date }]);
    if (error) return res.json({ success: false, message: error.message });
    res.json({ success: true });
});

app.post('/api/export-warehouse', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập!' });
    const { customer_name, address, phone, project_name } = req.body;
    const ticket_code = 'PXK-' + Date.now();
    const { error } = await supabase.from('export_tickets').insert([{ ticket_code, customer_name, address, phone, project_name }]);
    if (error) return res.json({ success: false, message: error.message });
    res.json({ success: true, ticket_code });
});

app.get('/api/export-revenue-excel', async (req, res) => {
    const workbook = new ExcelJS.Workbook();
    const sheetRevenue = workbook.addWorksheet('Doanh Thu');
    sheetRevenue.columns = [
        { header: 'Mã Báo Giá', key: 'code', width: 15 },
        { header: 'Khách Hàng', key: 'customer', width: 25 },
        { header: 'Dự Án', key: 'project', width: 30 },
        { header: 'Số Tiền', key: 'amount', width: 15 }
    ];
    const { data: quotes } = await supabase.from('quotes').select('*');
    (quotes || []).forEach(q => sheetRevenue.addRow({ code: q.quote_code, customer: q.customer_name, project: q.project_name, amount: q.total_amount }));

    const sheetExpense = workbook.addWorksheet('Chi Phí');
    sheetExpense.columns = [
        { header: 'Loại Chi Phí', key: 'type', width: 25 },
        { header: 'Mô Tả', key: 'desc', width: 35 },
        { header: 'Số Tiền', key: 'amount', width: 15 },
        { header: 'Ngày', key: 'date', width: 15 }
    ];
    const { data: expenses } = await supabase.from('expenses').select('*');
    (expenses || []).forEach(e => sheetExpense.addRow({ type: e.expense_type, desc: e.description, amount: e.amount, date: e.expense_date }));

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Bao_Cao_Doanh_Thu_Chi_Phi.xlsx');
    await workbook.xlsx.write(res);
    res.end();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

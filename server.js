const express = require('express');
const ExcelJS = require('exceljs');
const supabase = require('./supabaseClient');
require('dotenv').config();

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));
app.set('view engine', 'ejs');

// --- KHAI BÁO BIẾN TOÀN CỤC PHIÊN DĂNG NHẬP ---
let isAdminLoggedIn = false;
let currentUsername = null;

// --- TRANG CHỦ & DANH MỤC SẢN PHẨM ---
app.get('/', async (req, res) => {
    try {
        const { data: categories } = await supabase.from('categories').select('*');
        const { data: services } = await supabase.from('services').select('*');
        const { data: projects } = await supabase.from('project_customers').select('*');
        const { data: quotes } = await supabase.from('quotes').select('*');
        const { data: expenses } = await supabase.from('expenses').select('*');

        // Tính trạng thái bảo hành
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
            isAdmin: isAdminLoggedIn 
        });
    } catch (err) {
        console.error('Lỗi trang chủ:', err);
        res.status(500).send('Lỗi kết nối cơ sở dữ liệu!');
    }
});

// Trang web con: Hiển thị sản phẩm theo danh mục
app.get('/category/:id', async (req, res) => {
    try {
        const categoryId = req.params.id;
        
        // Bắt lỗi ép kiểu hoặc truy vấn linh hoạt
        let { data: category } = await supabase.from('categories').select('*').eq('id', categoryId).maybeSingle();
        if (!category) {
            const { data: allCats } = await supabase.from('categories').select('*');
            category = (allCats || []).find(c => String(c.id) === String(categoryId)) || { id: categoryId, name: 'Danh mục sản phẩm' };
        }

        const { data: products } = await supabase.from('products').select('*').eq('category_id', categoryId);

        res.render('category', { 
            category: category || {}, 
            products: products || [],
            isAdmin: isAdminLoggedIn 
        });
    } catch (err) {
        console.error('Lỗi tải danh mục:', err);
        res.status(500).send('Không thể tải sản phẩm của danh mục này!');
    }
});

// --- API XỬ LÝ ĐĂNG NHẬP / ĐĂNG XUẤT ---
app.post('/api/login', async (req, res) => {
    try {
        const { username, password } = req.body;
        const { data: users, error } = await supabase
            .from('app_users')
            .select('*')
            .eq('username', username)
            .eq('password', password);
        
        if (error || !users || users.length === 0) {
            return res.json({ success: false, message: 'Sai tên đăng nhập hoặc mật khẩu!' });
        }
        
        isAdminLoggedIn = true;
        currentUsername = users[0].username;
        return res.json({ success: true, message: 'Đăng nhập thành công!' });
    } catch (err) {
        console.error('Lỗi đăng nhập:', err);
        return res.json({ success: false, message: 'Lỗi hệ thống khi đăng nhập!' });
    }
});

app.post('/api/logout', (req, res) => {
    isAdminLoggedIn = false;
    currentUsername = null;
    res.json({ success: true });
});

// --- API QUẢN LÝ SẢN PHẨM & KHO ---
app.post('/api/products/save', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Chưa đăng nhập quyền Admin!' });
    const { id, name, category_id, description, quantity, cost_price, selling_price, image_url } = req.body;
    
    let result;
    if (id) {
        result = await supabase.from('products').update({ name, category_id, description, quantity, cost_price, selling_price, image_url }).eq('id', id);
    } else {
        result = await supabase.from('products').insert([{ name, category_id, description, quantity, cost_price, selling_price, image_url }]);
    }

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

// --- API QUẢN LÝ TÀI KHOẢN ADMIN ---
app.post('/api/change-password', async (req, res) => {
    if (!isAdminLoggedIn || !currentUsername) return res.json({ success: false, message: 'Chưa đăng nhập!' });
    const { oldPassword, newPassword } = req.body;
    const { data: user } = await supabase.from('app_users').select('*').eq('username', currentUsername).eq('password', oldPassword).maybeSingle();
    
    if (!user) return res.json({ success: false, message: 'Mật khẩu cũ không đúng!' });
    await supabase.from('app_users').update({ password: newPassword }).eq('username', currentUsername);
    res.json({ success: true, message: 'Đổi mật khẩu thành công!' });
});

app.post('/api/users/save', async (req, res) => {
    if (!isAdminLoggedIn) return res.json({ success: false, message: 'Không có quyền!' });
    const { username, password } = req.body;
    const { error } = await supabase.from('app_users').insert([{ username, password, role: 'admin' }]);
    if (error) return res.json({ success: false, message: error.message });
    res.json({ success: true });
});

// --- API XUẤT KHO ---
app.post('/api/export-warehouse', async (req, res) => {
    const { quote_id, customer_name, address, phone, project_name, items } = req.body;
    const ticket_code = 'PXK-' + Date.now();

    const { data: ticket, error } = await supabase.from('export_tickets').insert([{
        ticket_code, quote_id: quote_id || null, customer_name, address, phone, project_name
    }]).select().single();

    if (error) return res.status(500).json({ error: error.message });

    for (const item of (items || [])) {
        await supabase.from('export_items').insert([{
            export_id: ticket.id,
            product_id: item.product_id,
            quantity: item.quantity
        }]);

        const { data: prod } = await supabase.from('products').select('quantity').eq('id', item.product_id).maybeSingle();
        if (prod) {
            const newQty = Math.max(0, prod.quantity - item.quantity);
            await supabase.from('products').update({ quantity: newQty }).eq('id', item.product_id);
        }
    }

    res.json({ success: true, ticket_code });
});

// --- XUẤT BÁO CÁO EXCEL DOANH THU & CHI PHÍ ---
app.get('/api/export-revenue-excel', async (req, res) => {
    const { startDate, endDate } = req.query;
    const workbook = new ExcelJS.Workbook();

    const sheetRevenue = workbook.addWorksheet('Doanh Thu Dự Án');
    sheetRevenue.columns = [
        { header: 'Mã Báo Giá/Dự Án', key: 'code', width: 20 },
        { header: 'Khách Hàng', key: 'customer', width: 25 },
        { header: 'Tên Dự Án', key: 'project', width: 30 },
        { header: 'Địa Chỉ', key: 'address', width: 30 },
        { header: 'Số Tiền', key: 'amount', width: 15 }
    ];

    let queryQuotes = supabase.from('quotes').select('*');
    if (startDate) queryQuotes = queryQuotes.gte('created_at', startDate);
    if (endDate) queryQuotes = queryQuotes.lte('created_at', endDate);
    const { data: quotes } = await queryQuotes;

    (quotes || []).forEach(q => {
        sheetRevenue.addRow({
            code: q.quote_code,
            customer: q.customer_name,
            project: q.project_name,
            address: q.address,
            amount: q.total_amount
        });
    });

    const sheetExpense = workbook.addWorksheet('Chi Phí');
    sheetExpense.columns = [
        { header: 'Loại Chi Phí', key: 'type', width: 25 },
        { header: 'Mô Tả Chi Tiết', key: 'desc', width: 35 },
        { header: 'Số Tiền', key: 'amount', width: 15 },
        { header: 'Ngày Nhập', key: 'date', width: 20 }
    ];

    let queryExp = supabase.from('expenses').select('*');
    if (startDate) queryExp = queryExp.gte('expense_date', startDate);
    if (endDate) queryExp = queryExp.lte('expense_date', endDate);
    const { data: expenses } = await queryExp;

    (expenses || []).forEach(e => {
        sheetExpense.addRow({
            type: e.expense_type,
            desc: e.description,
            amount: e.amount,
            date: e.expense_date
        });
    });

    const sheetProfit = workbook.addWorksheet('Tổng Hợp Lợi Nhuận');
    const totalRev = (quotes || []).reduce((sum, item) => sum + Number(item.total_amount || 0), 0);
    const totalExp = (expenses || []).reduce((sum, item) => sum + Number(item.amount || 0), 0);
    
    sheetProfit.addRow(['TỔNG DOANH THU', totalRev]);
    sheetProfit.addRow(['TỔNG CHI PHÍ', totalExp]);
    sheetProfit.addRow(['TỔNG LỢI NHUẬN', totalRev - totalExp]);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=' + 'Bao_Cao_Doanh_Thu.xlsx');

    await workbook.xlsx.write(res);
    res.end();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));

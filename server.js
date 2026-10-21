const express = require('express');
const ExcelJS = require('exceljs');
const supabase = require('./supabaseClient');
require('dotenv').config();

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));
app.set('view engine', 'ejs');

// --- KHAI BÁO BIẾN TOÀN CỤC QUẢN LÝ PHIÊN ĐĂNG NHẬP ---
let isAdminLoggedIn = false;
let currentUsername = null;

// --- TRANG CHỦ & DANH MỤC KHÁCH XEM ---
app.get('/', async (req, res) => {
    const { data: categories } = await supabase.from('categories').select('*');
    const { data: services } = await supabase.from('services').select('*');
    const { data: projects } = await supabase.from('project_customers').select('*');

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
        isAdmin: isAdminLoggedIn 
    });
});

// Web con: Hiển thị sản phẩm theo danh mục (Lưới 3 ô / dòng, Không hiển thị giá nhập)
app.get('/category/:id', async (req, res) => {
    const categoryId = req.params.id;
    const { data: category } = await supabase.from('categories').select('*').eq('id', categoryId).single();
    const { data: products } = await supabase.from('products').select('id, name, image_url, description, quantity, selling_price').eq('category_id', categoryId);

    res.render('category', { category, products: products || [] });
});

// --- XỬ LÝ ĐĂNG NHẬP & QUẢN LÝ TÀI KHOẢN ADMIN ---
app.post('/api/login', async (req, res) => {
    try {
        const { username, password } = req.body;
        
        // Truy vấn kiểm tra thông tin trong bảng app_users của Supabase
        const { data: users, error } = await supabase
            .from('app_users')
            .select('*')
            .eq('username', username)
            .eq('password', password);
        
        if (error || !users || users.length === 0) {
            return res.json({ success: false, message: 'Sai tên đăng nhập hoặc mật khẩu!' });
        }
        
        // Đăng nhập thành công, lưu trạng thái phiên
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

// Đổi mật khẩu
app.post('/api/change-password', async (req, res) => {
    if (!isAdminLoggedIn || !currentUsername) {
        return res.json({ success: false, message: 'Bạn chưa đăng nhập quyền admin!' });
    }

    const { oldPassword, newPassword } = req.body;
    const { data: user } = await supabase.from('app_users').select('*').eq('username', currentUsername).eq('password', oldPassword).single();
    
    if (!user) {
        return res.json({ success: false, message: 'Mật khẩu cũ không đúng!' });
    }

    await supabase.from('app_users').update({ password: newPassword }).eq('username', currentUsername);
    return res.json({ success: true, message: 'Đổi mật khẩu thành công!' });
});

// Thêm tài khoản admin mới
app.post('/api/users/save', async (req, res) => {
    if (!isAdminLoggedIn) {
        return res.json({ success: false, message: 'Không có quyền thực hiện!' });
    }

    const { username, password } = req.body;
    const { error } = await supabase.from('app_users').insert([{ username, password, role: 'admin' }]);
    if (error) return res.json({ success: false, message: error.message });
    res.json({ success: true });
});

app.post('/api/users/delete', async (req, res) => {
    if (!isAdminLoggedIn) {
        return res.json({ success: false, message: 'Không có quyền thực hiện!' });
    }

    const { id } = req.body;
    await supabase.from('app_users').delete().eq('id', id);
    res.json({ success: true });
});

// --- TAB CHI PHÍ & DOANH THU EXCEL EXPORT ---
app.get('/api/export-revenue-excel', async (req, res) => {
    const { startDate, endDate } = req.query;

    const workbook = new ExcelJS.Workbook();

    // Sheet 1: Dự án / Doanh thu
    const sheetRevenue = workbook.addWorksheet('Doanh Thu Dự Án');
    sheetRevenue.columns = [
        { header: 'Mã Dự Án/Báo Giá', key: 'code', width: 20 },
        { header: 'Khách Hàng', key: 'customer', width: 25 },
        { header: 'Nội Dung/Tên Dự Án', key: 'project', width: 30 },
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

    // Sheet 2: Chi phí
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

    // Sheet 3: Tổng Hợp Lợi Nhuận
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

// --- API XUẤT KHO VÀ GIẢM TRỪ TỒN KHO ---
app.post('/api/export-warehouse', async (req, res) => {
    const { quote_id, customer_name, address, phone, project_name, items } = req.body;
    const ticket_code = 'PXK-' + Date.now();

    // 1. Tạo phiếu xuất
    const { data: ticket, error } = await supabase.from('export_tickets').insert([{
        ticket_code, quote_id: quote_id || null, customer_name, address, phone, project_name
    }]).select().single();

    if (error) return res.status(500).json({ error: error.message });

    // 2. Thêm items & trừ tồn kho
    for (const item of (items || [])) {
        await supabase.from('export_items').insert([{
            export_id: ticket.id,
            product_id: item.product_id,
            quantity: item.quantity
        }]);

        // Giảm tồn kho sản phẩm trong DB
        const { data: prod } = await supabase.from('products').select('quantity').eq('id', item.product_id).single();
        if (prod) {
            const newQty = Math.max(0, prod.quantity - item.quantity);
            await supabase.from('products').update({ quantity: newQty }).eq('id', item.product_id);
        }
    }

    res.json({ success: true, ticket_code });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));

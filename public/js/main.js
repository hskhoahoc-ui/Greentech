document.addEventListener("DOMContentLoaded", function () {
    console.log("Hệ thống Web Quản Lý Thiết Bị Điện đã sẵn sàng!");

    // Hiệu ứng cuộn mượt cho menu
    document.querySelectorAll('nav.menu a').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            const targetId = this.getAttribute('href');
            if (targetId.startsWith('#')) {
                e.preventDefault();
                const targetElement = document.querySelector(targetId);
                if (targetElement) {
                    targetElement.scrollIntoView({ behavior: 'smooth' });
                }
            }
        });
    });
});

// Hàm hỗ trợ gọi API đăng nhập Admin khi cần thiết
async function adminLogin(username, password) {
    try {
        const response = await fetch('/api/login-admin', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ username, password })
        });
        const result = await response.json();
        if (result.success) {
            alert('Đăng nhập quản trị thành công!');
            window.location.reload();
        } else {
            alert('Lỗi: ' + result.message);
        }
    } catch (error) {
        console.error('Lỗi kết nối:', error);
    }
}
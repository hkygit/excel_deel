// 全局交互逻辑
document.addEventListener('DOMContentLoaded', () => {
    console.log('✨ Excel 智能工具集已加载');

    // 为所有提交按钮添加加载状态（可选）
    document.querySelectorAll('form').forEach(form => {
        form.addEventListener('submit', () => {
            const btn = form.querySelector('button[type="submit"]');
            if (btn) {
                btn.disabled = true;
                btn.innerHTML = '<span class="spinner"></span> 处理中...';
            }
        });
    });
});

document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('transposerForm');
    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        const statusArea = document.getElementById('statusArea');
        const statusMsg = document.getElementById('statusMessage');
        const submitBtn = document.getElementById('submitBtn');

        statusArea.style.display = 'block';
        statusMsg.textContent = '处理中，请稍候...';
        submitBtn.disabled = true;

        const formData = new FormData(form);
        try {
            const response = await fetch('/transposer/convert', {
                method: 'POST',
                body: formData
            });
            if (!response.ok) {
                const errData = await response.json();
                statusMsg.textContent = '错误：' + (errData.error || '未知错误');
                submitBtn.disabled = false;
                return;
            }
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const filename = getFilenameFromContentDisposition(response.headers.get('Content-Disposition')) || '转置结果.xlsx';
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            statusMsg.textContent = '转换成功，文件已下载。';
        } catch (err) {
            statusMsg.textContent = '网络错误：' + err.message;
        } finally {
            submitBtn.disabled = false;
        }
    });
});

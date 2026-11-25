// static/js/concat.js
(function () {
    const form = document.getElementById('concatForm');
    if (!form) return;

    form.addEventListener('submit', async function (e) {
        e.preventDefault();
        const statusArea = document.getElementById('statusArea');
        const statusMsg = document.getElementById('statusMessage');
        const submitBtn = document.getElementById('submitBtn');

        statusArea.style.display = 'block';
        statusMsg.textContent = '处理中，请稍候...';
        submitBtn.disabled = true;

        const formData = new FormData(form);
        try {
            const resp = await fetch('/concat/process', {
                method: 'POST',
                body: formData,
            });
            if (!resp.ok) {
                let msg = `请求失败 (${resp.status})`;
                try {
                    const err = await resp.json();
                    msg = err.error || err.message || msg;
                } catch (_) {}
                statusMsg.textContent = '错误：' + msg;
                submitBtn.disabled = false;
                return;
            }
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const filename = getFilenameFromContentDisposition(
                resp.headers.get('Content-Disposition')
            ) || '分组拼接结果.xlsx';
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            statusMsg.textContent = '处理完成，文件已下载。';
        } catch (err) {
            statusMsg.textContent = '网络错误：' + err.message;
        } finally {
            submitBtn.disabled = false;
        }
    });
})();

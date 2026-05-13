// convert.js
(function() {
    document.getElementById('convertForm').addEventListener('submit', async function(e) {
        e.preventDefault();
        const fileInput = document.getElementById('fileInput');
        const targetFormat = document.getElementById('targetFormat').value;
        const encoding = document.getElementById('encoding').value;
        const separator = document.getElementById('separator').value;
        if (!fileInput.files.length) { alert('请选择文件'); return; }

        const btn = document.getElementById('btnSubmit');
        btn.disabled = true;
        btn.textContent = '提交中...';
        document.getElementById('statusArea').style.display = 'block';
        document.getElementById('taskStatus').textContent = '提交中...';
        document.getElementById('progressBar').style.width = '0%';

        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        formData.append('target_format', targetFormat);
        formData.append('encoding', encoding);
        formData.append('separator', separator);

        try {
            const resp = await fetch('/convert/process', { method: 'POST', body: formData });
            const result = await resp.json();
            if (result.error) {
                document.getElementById('taskStatus').textContent = '错误: ' + result.error;
                btn.disabled = false;
                return;
            }
            const taskId = result.task_id;
            document.getElementById('taskStatus').textContent = '转换中...';

            const poll = setInterval(async () => {
                const statusResp = await fetch('/convert/status/' + taskId);
                const status = await statusResp.json();
                if (status.error) {
                    clearInterval(poll);
                    document.getElementById('taskStatus').textContent = '错误: ' + status.error;
                    btn.disabled = false;
                    return;
                }
                if (status.progress_percent !== undefined) {
                    document.getElementById('progressBar').style.width = status.progress_percent + '%';
                }
                if (status.progress_message) {
                    document.getElementById('progressMessage').textContent = status.progress_message;
                }
                if (status.status === 'completed') {
                    clearInterval(poll);
                    document.getElementById('taskStatus').textContent = '完成！';
                    document.getElementById('downloadLink').innerHTML = `<a href="/convert/download/${taskId}" class="btn btn-success">下载结果</a>`;
                    btn.disabled = false;
                    btn.textContent = '开始转换';
                } else if (status.status === 'failed') {
                    clearInterval(poll);
                    document.getElementById('taskStatus').textContent = '失败: ' + (status.error || '');
                    btn.disabled = false;
                    btn.textContent = '开始转换';
                }
            }, 1000);
        } catch (err) {
            document.getElementById('taskStatus').textContent = '网络错误: ' + err.message;
            btn.disabled = false;
        }
    });
})();

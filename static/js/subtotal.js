// static/js/subtotal.js
(function() {
    const form = document.getElementById('subtotalForm');
    const statusArea = document.getElementById('statusArea');
    const addOverrideBtn = document.getElementById('addOverrideBtn');
    const overrideContainer = document.getElementById('overrideContainer');
    const sheetsModeRadios = document.getElementsByName('sheets_mode');
    const sheetsCustom = document.getElementById('sheetsCustom');

    addOverrideBtn.addEventListener('click', function() {
        const template = document.getElementById('overrideRowTemplate');
        const clone = template.content.cloneNode(true);
        overrideContainer.appendChild(clone);
    });

    sheetsModeRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            sheetsCustom.style.display = this.value === 'custom' ? 'inline-block' : 'none';
        });
    });

    form.addEventListener('submit', async function(e) {
        e.preventDefault();
        // 收集特殊列函数
        const overrides = [];
        document.querySelectorAll('.override-row').forEach(row => {
            const col = row.querySelector('.override-col').value.trim().toUpperCase();
            const func = row.querySelector('.override-func').value;
            if (col) {
                overrides.push(col + ':' + func);
            }
        });

        const formData = new FormData(form);
        if (overrides.length > 0) {
            formData.set('agg_overrides', overrides.join(','));
        }
        const mode = formData.get('sheets_mode');
        if (mode === 'custom') {
            const customVal = sheetsCustom.value.trim();
            formData.set('sheets_mode', customVal || '1');
        }

        const submitBtn = form.querySelector('button[type="submit"]');
        submitBtn.disabled = true;
        statusArea.style.display = 'block';
        document.getElementById('taskStatus').textContent = '提交中...';
        document.getElementById('progressBar').style.width = '0%';
        document.getElementById('downloadLink').innerHTML = '';

        try {
            const resp = await fetch('/subtotal/process', { method: 'POST', body: formData });
            if (!resp.ok) {
                const err = await resp.json();
                document.getElementById('taskStatus').textContent = '错误: ' + (err.error || err.message);
                submitBtn.disabled = false;
                return;
            }
            const contentType = resp.headers.get('Content-Type');
            if (contentType && contentType.includes('application/json')) {
                // 异步任务
                const result = await resp.json();
                const taskId = result.task_id;
                pollTask(taskId, submitBtn);
            } else {
                // 同步下载
                    const blob = await resp.blob();
                const url = URL.createObjectURL(blob);
                const filename = getFilenameFromContentDisposition(resp.headers.get('Content-Disposition')) || 'subtotal_result.xlsx';
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                a.remove();
                URL.revokeObjectURL(url);
                document.getElementById('taskStatus').textContent = '处理完成，已下载。';
                submitBtn.disabled = false;
            }
        } catch (err) {
            document.getElementById('taskStatus').textContent = '网络错误: ' + err.message;
            submitBtn.disabled = false;
        }
    });

    function pollTask(taskId, submitBtn) {
        const interval = setInterval(async () => {
            try {
                const resp = await fetch('/subtotal/status/' + taskId);
                const status = await resp.json();
                if (status.error) {
                    clearInterval(interval);
                    document.getElementById('taskStatus').textContent = '错误: ' + status.error;
                    submitBtn.disabled = false;
                    return;
                }
                if (status.progress_percent !== undefined) {
                    document.getElementById('progressBar').style.width = status.progress_percent + '%';
                }
                document.getElementById('taskStatus').textContent =
                    status.status === 'completed' ? '完成' : (status.status || '处理中...');
                if (status.status === 'completed') {
                    clearInterval(interval);
                    document.getElementById('downloadLink').innerHTML =
                        `<a href="/subtotal/download/${taskId}" class="btn btn-success">下载结果</a>`;
                    submitBtn.disabled = false;
                } else if (status.status === 'failed') {
                    clearInterval(interval);
                    document.getElementById('taskStatus').textContent = '失败';
                    submitBtn.disabled = false;
                }
            } catch (err) {
                console.error(err);
            }
        }, 2000);
    }
})();

// team_report.js (updated to include monthly files)
(function() {
    const form = document.getElementById('reportForm');
    if (!form) return;

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const submitBtn = document.getElementById('submitBtn');
        const statusArea = document.getElementById('statusArea');
        const taskStatus = document.getElementById('taskStatus');
        const progressBar = document.getElementById('progressBar');
        const progressMessage = document.getElementById('progressMessage');

        submitBtn.disabled = true;
        statusArea.style.display = 'block';
        taskStatus.textContent = '提交中...';
        progressBar.style.width = '0%';
        progressMessage.textContent = '';

        const formData = new FormData(form);

        try {
            const resp = await fetch('/team-report/generate', {
                method: 'POST',
                body: formData
            });
            if (!resp.ok) {
                taskStatus.textContent = '提交失败 (HTTP ' + resp.status + ')';
                submitBtn.disabled = false;
                return;
            }
            const result = await resp.json();
            if (result.error) {
                taskStatus.textContent = '错误: ' + result.error;
                submitBtn.disabled = false;
                return;
            }
            const taskId = result.task_id;
            taskStatus.textContent = '任务已提交，正在处理...';
            progressBar.style.width = '10%';

            const poll = setInterval(async () => {
                try {
                    const statusResp = await fetch('/team-report/status/' + taskId);
                    if (!statusResp.ok) {
                        if (statusResp.status === 524 || statusResp.status === 502) {
                            console.warn('收到 ' + statusResp.status + '，继续等待...');
                            return;
                        }
                        clearInterval(poll);
                        taskStatus.textContent = '查询状态失败 (HTTP ' + statusResp.status + ')';
                        submitBtn.disabled = false;
                        return;
                    }
                    const status = await statusResp.json();
                    if (status.error) {
                        clearInterval(poll);
                        taskStatus.textContent = '错误: ' + status.error;
                        submitBtn.disabled = false;
                        return;
                    }
                    if (status.progress_percent !== undefined) {
                        progressBar.style.width = status.progress_percent + '%';
                    }
                    if (status.progress_message) {
                        progressMessage.textContent = status.progress_message;
                    }
                    if (status.status === 'completed') {
                        clearInterval(poll);
                        taskStatus.textContent = '处理完成，正在下载...';
                        progressBar.style.width = '100%';
                        window.location.href = '/team-report/download/' + taskId;
                        setTimeout(() => { submitBtn.disabled = false; }, 2000);
                    } else if (status.status === 'failed') {
                        clearInterval(poll);
                        taskStatus.textContent = '处理失败: ' + (status.error || '未知错误');
                        submitBtn.disabled = false;
                    }
                } catch (err) {
                    console.error('轮询错误:', err);
                }
            }, 4000);
        } catch (err) {
            taskStatus.textContent = '网络错误: ' + err.message;
            submitBtn.disabled = false;
        }
    });
})();

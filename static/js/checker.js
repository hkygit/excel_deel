// checker.js
(function() {
    // 确保 DOM 加载完成后再绑定事件
    document.addEventListener('DOMContentLoaded', function() {
        // 定义全局 addFilter 函数
        window.addFilter = function() {
            const container = document.getElementById('filtersContainer');
            const template = document.getElementById('filterRowTemplate');
            if (template) container.appendChild(template.content.cloneNode(true));
        };

        // 解析列名
        function parseColumns(text) {
            if (!text) return [];
            const lines = text.split(/[\n\r]+/);
            const cols = [];
            lines.forEach(line => {
                const parts = line.split(/[,，\s]+/);
                parts.forEach(p => {
                    const trimmed = p.trim();
                    if (trimmed) cols.push(trimmed);
                });
            });
            return cols;
        }

        // 收集对比列映射
        function collectCompareMap() {
            const stdText = document.getElementById('compareStandardText').value;
            const cmpText = document.getElementById('compareCompareText').value;
            const stdCols = parseColumns(stdText);
            let cmpCols = parseColumns(cmpText);
            if (cmpCols.length === 0) cmpCols = [...stdCols];
            if (stdCols.length === 0) {
                alert('请填写标准对比列');
                return null;
            }
            if (stdCols.length !== cmpCols.length) {
                alert(`标准对比列数量 (${stdCols.length}) 与比较对比列数量 (${cmpCols.length}) 不一致，请检查输入。`);
                return null;
            }
            const map = [];
            for (let i = 0; i < stdCols.length; i++) {
                map.push({ standard: stdCols[i], compare: cmpCols[i] });
            }
            return map;
        }

        // 获取 DOM 元素
        const btnPrecheck = document.getElementById('btnPrecheck');
        const btnStart = document.getElementById('btnStart');
        const statusArea = document.getElementById('statusArea');
        const taskStatus = document.getElementById('taskStatus');
        const progressBar = document.getElementById('progressBar');
        const progressMessage = document.getElementById('progressMessage');
        const precheckResult = document.getElementById('precheckResult');
        const downloadLink = document.getElementById('downloadLink');

        // 预检按钮点击事件
        btnPrecheck.addEventListener('click', async function() {
            const filesA = document.getElementById('filesA').files;
            const filesB = document.getElementById('filesB').files;
            if (filesA.length === 0 || filesB.length === 0) {
                alert('请上传文件');
                return;
            }

            btnPrecheck.disabled = true;
            btnPrecheck.textContent = '⏳ 预检中...';
            statusArea.style.display = 'block';
            taskStatus.textContent = '正在预检，请稍后...';
            progressBar.style.width = '10%';
            progressMessage.textContent = '正在上传文件并扫描...';
            precheckResult.innerHTML = '';

            const formData = new FormData();
            for (let f of filesA) formData.append('files_a', f);
            for (let f of filesB) formData.append('files_b', f);

            try {
                const resp = await fetch('/checker/precheck', { method: 'POST', body: formData });
                const result = await resp.json();
                if (resp.ok) {
                    precheckResult.innerHTML = `
                        <p>标准文件: ${result.total_files_a} 个, 约 ${result.total_rows_a.toLocaleString()} 行</p>
                        <p>比较文件: ${result.total_files_b} 个, 约 ${result.total_rows_b.toLocaleString()} 行</p>
                        <p>预估内存: ${result.estimated_memory_mb} MB | 预估时间: ${result.estimated_time_seconds} 秒</p>
                        <p style="color:${result.advice.includes('安全') ? 'green' : 'orange'}">${result.advice}</p>
                    `;
                    taskStatus.textContent = '预检完成，可以开始校验';
                    progressBar.style.width = '100%';
                    btnStart.disabled = false;
                } else {
                    taskStatus.textContent = '预检失败';
                    precheckResult.innerHTML = `<p style="color:red;">错误: ${result.message || result.error}</p>`;
                    btnStart.disabled = true;
                }
            } catch (err) {
                taskStatus.textContent = '网络错误';
                precheckResult.innerHTML = `<p style="color:red;">网络请求失败: ${err.message}</p>`;
                btnStart.disabled = true;
            } finally {
                btnPrecheck.disabled = false;
                btnPrecheck.textContent = '🔎 预检';
            }
        });

        // 开始校验按钮点击事件
        btnStart.addEventListener('click', async function() {
            const filesA = document.getElementById('filesA').files;
            const filesB = document.getElementById('filesB').files;
            const keyA = document.getElementById('keyMapA').value.trim();
            const keyB = document.getElementById('keyMapB').value.trim();
            const duplicateAction = document.getElementById('duplicateAction').value;
            const compareMap = collectCompareMap();
            if (!compareMap) return;

            const filters = [];
            document.querySelectorAll('#filtersContainer .filter-row').forEach(row => {
                const field = row.querySelector('input[name="filter_field"]')?.value.trim();
                const op = row.querySelector('select[name="filter_op"]')?.value;
                const value = row.querySelector('input[name="filter_value"]')?.value.trim();
                if (field && value) filters.push({field, op, value});
            });

            if (!keyA) { alert('请填写标准群主键'); return; }

            const formData = new FormData();
            for (let f of filesA) formData.append('files_a', f);
            for (let f of filesB) formData.append('files_b', f);
            formData.append('key_map_a', keyA);
            formData.append('key_map_b', keyB || keyA);
            formData.append('compare_map', JSON.stringify(compareMap));
            formData.append('filters_a', JSON.stringify(filters));
            formData.append('duplicate_action', duplicateAction);

            btnStart.disabled = true;
            btnStart.textContent = '提交中...';
            statusArea.style.display = 'block';
            taskStatus.textContent = '提交中...';
            progressBar.style.width = '0%';
            progressMessage.textContent = '';

            const resp = await fetch('/checker/process', { method: 'POST', body: formData });
            const result = await resp.json();
            if (result.error) {
                taskStatus.textContent = '错误: ' + result.error;
                btnStart.disabled = false;
                btnStart.textContent = '▶ 开始校验';
                return;
            }
            const taskId = result.task_id;
            taskStatus.textContent = '任务已提交，处理中...';

            const poll = setInterval(async () => {
                try {
                    const statusResp = await fetch('/checker/status/' + taskId);
                    const status = await statusResp.json();
                    if (status.error) {
                        clearInterval(poll);
                        taskStatus.textContent = '错误: ' + status.error;
                        btnStart.disabled = false;
                        btnStart.textContent = '▶ 开始校验';
                        return;
                    }
                    if (status.progress_percent !== undefined && status.progress_percent !== null) {
                        progressBar.style.width = status.progress_percent + '%';
                    }
                    if (status.progress_message) {
                        progressMessage.textContent = status.progress_message;
                    }
                    if (status.processed_rows !== undefined && status.total_rows) {
                        taskStatus.textContent = `扫描中 (${status.processed_rows.toLocaleString()} / ${status.total_rows.toLocaleString()} 行)`;
                    }
                    if (status.status === 'completed') {
                        clearInterval(poll);
                        taskStatus.textContent = '完成！';
                        progressBar.style.width = '100%';
                        downloadLink.innerHTML = `<a href="/checker/download/${taskId}" class="btn btn-success">下载结果</a>`;
                        btnStart.disabled = true;
                    } else if (status.status === 'failed') {
                        clearInterval(poll);
                        taskStatus.textContent = '失败: ' + (status.error || '');
                        btnStart.disabled = false;
                        btnStart.textContent = '▶ 开始校验';
                    }
                } catch (err) {
                    console.error('轮询失败', err);
                }
            }, 1000);
        });
    });
})();

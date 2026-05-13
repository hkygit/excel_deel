(function() {
    document.addEventListener('DOMContentLoaded', () => {
        const form = document.getElementById('simplifyForm');
        if (!form) return;

        window.addFilter = function() {
            const container = document.getElementById('filtersContainer');
            const template = document.getElementById('filterRowTemplate');
            if (template) {
                const clone = template.content.cloneNode(true);
                container.appendChild(clone);
            } else {
                const div = document.createElement('div');
                div.className = 'filter-row';
                div.innerHTML = `
                    <input type="text" name="filter_field" placeholder="字段名" class="filter-field" />
                    <select name="filter_op" class="filter-op">
                        <option value="=">=</option>
                        <option value="!=">!=</option>
                        <option value=">">></option>
                        <option value="<"><</option>
                        <option value=">=">>=</option>
                        <option value="<="><=</option>
                        <option value="contains">包含</option>
                        <option value="not_contains">不包含</option>
                    </select>
                    <input type="text" name="filter_value" placeholder="值（多值用英文逗号）" class="filter-value" />
                    <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">删除</button>
                `;
                container.appendChild(div);
            }
        };

        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            const statusArea = document.getElementById('statusArea');
            const taskStatus = document.getElementById('taskStatus');
            const progressBar = document.getElementById('progressBar');
            const logArea = document.getElementById('logArea');
            const inputRows = document.getElementById('inputRows');
            const outputRows = document.getElementById('outputRows');
            const resultFormat = document.getElementById('resultFormat');
            const downloadLink = document.getElementById('downloadLink');

            statusArea.style.display = 'block';
            taskStatus.innerText = '提交中...';
            progressBar.style.width = '10%';

            const filters = [];
            document.querySelectorAll('#filtersContainer .filter-row').forEach(row => {
                const field = row.querySelector('input[name="filter_field"]')?.value.trim();
                const op = row.querySelector('select[name="filter_op"]')?.value;
                const value = row.querySelector('input[name="filter_value"]')?.value.trim();
                if (field && value) filters.push({ field, op, value });
            });

            const params = {
                key_cols: document.getElementById('keyCols')?.value.trim() || '',
                keep_cols: document.getElementById('keepCols')?.value.trim() || '',
                non_sum_fields: document.getElementById('nonSumFields')?.value.trim() || '',
                filters: filters,
                output_format: document.getElementById('outputFormat')?.value || 'auto',
                column_mode: document.getElementById('columnMode')?.value || 'name'
            };

            const formData = new FormData(form);
            formData.append('params', JSON.stringify(params));

            try {
                const resp = await fetch('/simplify/process', { method: 'POST', body: formData });
                const result = await resp.json();
                if (result.error) {
                    taskStatus.innerText = '错误: ' + result.error;
                    progressBar.style.width = '100%';
                    progressBar.style.background = 'var(--danger)';
                    return;
                }
                const taskId = result.task_id;
                taskStatus.innerText = '处理中...';
                progressBar.style.width = '20%';

                const poll = setInterval(async () => {
                    const sResp = await fetch('/simplify/status/' + taskId);
                    const s = await sResp.json();
                    if (s.error) { clearInterval(poll); taskStatus.innerText = '错误: ' + s.error; return; }
                    if (s.log) logArea.innerText = s.log;
                    if (s.total_input_rows) inputRows.innerText = s.total_input_rows;
                    if (s.total_output_rows) outputRows.innerText = s.total_output_rows;
                    if (s.result_format) resultFormat.innerText = s.result_format;
                    if (s.status === 'completed') {
                        clearInterval(poll);
                        taskStatus.innerText = '完成！';
                        progressBar.style.width = '100%';
                        progressBar.style.background = 'var(--success)';
                        downloadLink.innerHTML = `<a href="/simplify/download/${taskId}" class="btn btn-success">下载结果</a>`;
                    } else if (s.status === 'failed') {
                        clearInterval(poll);
                        taskStatus.innerText = '失败';
                        progressBar.style.width = '100%';
                        progressBar.style.background = 'var(--danger)';
                    } else {
                        progressBar.style.width = '50%';
                    }
                }, 2000);
            } catch (err) {
                taskStatus.innerText = '网络错误: ' + err.message;
                progressBar.style.width = '100%';
                progressBar.style.background = 'var(--danger)';
            }
        });
    });
})();

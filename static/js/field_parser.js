// field_parser.js
(function() {
    let lastResult = null;   // 保存最近分析结果（含 matched, only_a, only_b, cols_a, cols_b）

    document.getElementById('btnAnalyze').addEventListener('click', async function() {
        const textA = document.getElementById('textA').value.trim();
        const textB = document.getElementById('textB').value.trim();
        if (!textA || !textB) {
            alert('请填写两边的列名');
            return;
        }

        const btn = document.getElementById('btnAnalyze');
        btn.disabled = true;
        btn.textContent = '分析中...';

        try {
            const resp = await fetch('/field-parser/analyze', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    standard_headers: textA,
                    compare_headers: textB
                })
            });

            const result = await resp.json();
            if (resp.ok) {
                lastResult = result;   // 保存完整结果
                displayResult(result);
                document.getElementById('btnExport').style.display = 'inline-block';
            } else {
                alert('分析失败: ' + (result.message || result.error || '未知错误'));
            }
        } catch (err) {
            alert('网络请求失败: ' + err.message);
        } finally {
            btn.disabled = false;
            btn.textContent = '开始分析';
        }
    });

    document.getElementById('btnExport').addEventListener('click', async function() {
        if (!lastResult) {
            alert('请先进行分析');
            return;
        }
        // 发送完整映射结果和原始列名
        const payload = {
            matched: lastResult.matched,
            only_a: lastResult.only_a,
            only_b: lastResult.only_b,
            cols_a: lastResult.cols_a,
            cols_b: lastResult.cols_b
        };

        try {
            const resp = await fetch('/field-parser/export', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if (!resp.ok) {
                const err = await resp.json();
                alert('导出失败: ' + (err.message || '未知错误'));
                return;
            }
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = '字段映射结果.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (err) {
            alert('导出请求失败: ' + err.message);
        }
    });

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function displayResult(mapping) {
        const tbody = document.querySelector('#mappingTable tbody');
        tbody.innerHTML = '';
        mapping.matched.forEach(m => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${escapeHtml(m.standard)}</td><td>${escapeHtml(m.compare)}</td><td>${m.score}</td>`;
            tbody.appendChild(tr);
        });
        mapping.only_a.forEach(h => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${escapeHtml(h)}</td><td>-</td><td>仅标准</td>`;
            tbody.appendChild(tr);
        });
        mapping.only_b.forEach(h => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>-</td><td>${escapeHtml(h)}</td><td>仅比较</td>`;
            tbody.appendChild(tr);
        });
        document.getElementById('resultArea').style.display = 'block';
    }
})();

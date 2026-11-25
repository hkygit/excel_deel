- **templates/calculator/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}列计算 / 聚合回填{% endblock %}
{% block content %}
<h2>🧮 列计算 / 聚合回填</h2>
<form action="{{ url_for('calc.process') }}" method="post" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 上传 Excel 文件</label>
            <input type="file" name="file" accept=".xlsx,.xls" required>
        </div>
        <div class="form-row">
            <label>⚙️ 操作类型</label>
            <select name="calc_type" id="calcType">
                <option value="column">列计算</option>
                <option value="aggregate">聚合回填</option>
            </select>
        </div>
        <div class="form-row">
            <label>表头所在行号（从1开始，默认1）</label>
            <input type="number" name="header_row" value="1" min="1" style="max-width:100px;">
            <div class="hint">如果你的 Excel 顶部有标题行，请填写实际列名所在的行号</div>
        </div>
    </div>
    <div id="columnConfig" class="form-section">
        <h3>📐 列计算设置</h3>
        <div class="form-row">
            <label>表达式 (如 A+B-C)</label>
            <input type="text" name="expression" placeholder="例如 A+B">
        </div>
        <div class="form-row">
            <label>结果列 (字母或新列名)</label>
            <input type="text" name="result_col_column" placeholder="K 或 '差值'">
        </div>
        <div class="form-row">
            <label>合并依据列 (可选，逗号分隔)</label>
            <input type="text" name="merge_by_cols" placeholder="A,B">
            <div class="hint">相同组的结果列只保留第一行，其余合并。</div>
        </div>
    </div>
    <div id="aggregateConfig" class="form-section" style="display:none;">
        <h3>📊 聚合回填设置</h3>
        <div class="form-row">
            <label>分组列 (逗号分隔，支持列名或列字母)</label>
            <input type="text" name="group_cols" placeholder="A,B">
        </div>
        <div class="form-row">
            <label>聚合列 (逗号分隔)</label>
            <input type="text" name="agg_cols" placeholder="C,D">
        </div>
        <div class="form-row">
            <label>聚合函数</label>
            <select name="agg_func">
                <option value="sum">求和</option>
                <option value="mean">平均值</option>
                <option value="max">最大值</option>
                <option value="min">最小值</option>
                <option value="count">计数</option>
            </select>
        </div>
        <div class="form-row">
            <label>结果列名</label>
            <input type="text" name="result_col_aggregate" placeholder="K 或 '聚合结果'">
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始处理</button>
</form>
<script>
    const typeSelect = document.getElementById('calcType');
    const colCfg = document.getElementById('columnConfig');
    const aggCfg = document.getElementById('aggregateConfig');
    function toggle() {
        if (typeSelect.value === 'column') {
            colCfg.style.display = 'block';
            aggCfg.style.display = 'none';
        } else {
            colCfg.style.display = 'none';
            aggCfg.style.display = 'block';
        }
    }
    typeSelect.addEventListener('change', toggle);
    toggle();
</script>
{% endblock %}
    ```

- **templates/checker/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}文件校验{% endblock %}
{% block head %}
<link rel="stylesheet" href="{{ url_for('static', filename='css/checker.css') }}">
{% endblock %}
{% block content %}
<h2>✅ 文件校验</h2>
<p>上传标准文件群和比较文件群，系统将根据您指定的主键和对比列进行比较，并生成差异报告。</p>

<form id="checkerForm" enctype="multipart/form-data">
    <div class="form-section">
        <h3>📂 文件上传</h3>
        <div class="form-row">
            <label>标准文件群（可多选）</label>
            <input type="file" id="filesA" name="files_a" accept=".xlsx,.xls,.csv" multiple required>
        </div>
        <div class="form-row">
            <label>比较文件群（可多选）</label>
            <input type="file" id="filesB" name="files_b" accept=".xlsx,.xls,.csv" multiple required>
        </div>
    </div>

    <div class="form-section">
        <h3>🔑 主键设置</h3>
        <div class="form-row">
            <label>标准群主键列（逗号或空格分隔）</label>
            <input type="text" id="keyMapA" placeholder="例如 ID,日期">
        </div>
        <div class="form-row">
            <label>比较群主键列（逗号或空格分隔，默认与标准相同）</label>
            <input type="text" id="keyMapB" placeholder="留空则与标准群主键相同">
        </div>
        <div class="form-row">
            <label>主键重复处理</label>
            <select id="duplicateAction">
                <option value="error" selected>报错</option>
                <option value="keep_first">保留第一条</option>
            </select>
        </div>
    </div>

    <div class="form-section">
        <h3>📊 对比列</h3>
        <div class="parser-container" style="display: flex; gap: 20px;">
            <div class="parser-box" style="flex: 1;">
                <label>标准群对比列（每行一个或用逗号/空格分隔）</label>
                <textarea id="compareStandardText" rows="5" placeholder="例如：医疗类别&#10;入院日期&#10;出院日期&#10;医疗费总额&#10;结算时间&#10;定点医药机构名称"></textarea>
            </div>
            <div class="parser-box" style="flex: 1;">
                <label>比较群对比列（留空则与标准相同）</label>
                <textarea id="compareCompareText" rows="5" placeholder="例如：医疗类别&#10;入院日期&#10;出院日期&#10;医疗费总额&#10;结算时间&#10;定点医疗机构名称"></textarea>
            </div>
        </div>
        <div class="hint">按顺序一一对应，支持多列；比较群留空则默认与标准群列名相同。</div>
    </div>

    <div class="form-section">
        <h3>🔍 标准群筛选（可选）</h3>
        <div id="filtersContainer"></div>
        <button type="button" class="btn btn-secondary" onclick="addFilter()">+ 添加筛选条件</button>
        <div class="hint">支持多值OR：在“值”中用英文逗号分隔多个值</div>
    </div>

    <div style="text-align:center; margin: 30px 0;">
        <button type="button" class="btn" id="btnPrecheck">🔎 预检</button>
        <button type="button" class="btn btn-success" id="btnStart" disabled>▶ 开始校验</button>
    </div>
</form>

<div id="statusArea" class="status-area" style="display:none;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar" style="width:0%"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <p id="progressMessage" style="color:#666;"></p>
    <div id="precheckResult" class="precheck-result"></div>
    <div id="downloadLink" style="margin-top:15px;"></div>
</div>

<template id="filterRowTemplate">
    <div class="filter-row">
        <input type="text" name="filter_field" placeholder="字段名" class="filter-field">
        <select name="filter_op" class="filter-op">
            <option value="=">=</option><option value="!=">!=</option><option value=">">></option>
            <option value="<"><</option><option value=">=">>=</option><option value="<="><=</option>
            <option value="contains">包含</option><option value="not_contains">不包含</option>
        </select>
        <input type="text" name="filter_value" placeholder="值（多值用英文逗号）" class="filter-value">
        <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">删除</button>
    </div>
</template>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/checker.js') }}"></script>
{% endblock %}
    ```

- **templates/comparator/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}文件对比{% endblock %}
{% block content %}
<h2>🔍 文件对比</h2>
<form action="{{ url_for('compare.process') }}" method="post" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📂 文件 1</label>
            <input type="file" name="file1" accept=".xlsx,.xls" required>
            <input type="text" name="sheet1" placeholder="Sheet名称 (可选)" class="mt-2">
        </div>
        <div class="form-row">
            <label>📂 文件 2</label>
            <input type="file" name="file2" accept=".xlsx,.xls" required>
            <input type="text" name="sheet2" placeholder="Sheet名称 (可选)" class="mt-2">
        </div>
        <div class="form-row">
            <label>对比模式</label>
            <select name="compare_mode" id="compareMode">
                <option value="full">完整对比</option>
                <option value="field_match">仅字段值匹配</option>
            </select>
        </div>
    </div>
    <div id="fullFields" class="form-section">
        <div class="form-row">
            <label>主键字段 (逗号分隔)</label>
            <input type="text" name="key_fields" placeholder="学号,姓名">
        </div>
        <div class="form-row">
            <label>对比字段 (逗号分隔)</label>
            <input type="text" name="compare_fields" placeholder="学号,姓名,成绩">
        </div>
    </div>
    <div id="fieldMatchOnly" class="form-section" style="display:none;">
        <div class="form-row">
            <label>匹配字段名</label>
            <input type="text" name="match_field" placeholder="医疗机构名称">
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始对比</button>
</form>
<script>
    const modeSelect = document.getElementById('compareMode');
    const fullDiv = document.getElementById('fullFields');
    const fieldDiv = document.getElementById('fieldMatchOnly');
    modeSelect.addEventListener('change', () => {
        if (modeSelect.value === 'full') {
            fullDiv.style.display = 'block';
            fieldDiv.style.display = 'none';
        } else {
            fullDiv.style.display = 'none';
            fieldDiv.style.display = 'block';
        }
    });
    modeSelect.dispatchEvent(new Event('change'));
</script>
{% endblock %}
    ```

- **templates/convert/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}格式转换{% endblock %}
{% block content %}
<h2>🔄 文件格式转换</h2>
<p>上传文件，选择目标格式和编码，开始转换。</p>
<form id="convertForm" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 选择文件</label>
            <input type="file" id="fileInput" name="file" accept=".xlsx,.xls,.xlsm,.csv" required>
        </div>
        <div class="form-row">
            <label>目标格式</label>
            <select id="targetFormat" name="target_format">
                <option value="csv">CSV</option>
                <option value="xlsx">Excel (.xlsx)</option>
            </select>
        </div>
        <div class="form-row">
            <label>编码（仅CSV输出时有效）</label>
            <select id="encoding" name="encoding">
                <option value="utf-8-sig" selected>UTF-8 (带BOM)</option>
                <option value="utf-8">UTF-8</option>
                <option value="gbk">GBK</option>
                <option value="gb2312">GB2312</option>
            </select>
        </div>
        <div class="form-row">
            <label>分隔符（仅CSV输出时有效）</label>
            <select id="separator" name="separator">
                <option value="," selected>逗号 (,)</option>
                <option value=";">分号 (;)</option>
                <option value="	">制表符 (Tab)</option>
            </select>
        </div>
    </div>
    <button type="submit" class="btn btn-success" id="btnSubmit">开始转换</button>
</form>
<div id="statusArea" class="status-area" style="display:none;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar" style="width:0%"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <p id="progressMessage"></p>
    <div id="downloadLink" style="margin-top:15px;"></div>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/convert.js') }}"></script>
{% endblock %}
    ```

- **templates/debug.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}匹配填充调试{% endblock %}
{% block content %}
<h2>🔍 匹配填充调试信息</h2>
<div class="form-section">
    <h3>配置</h3>
    <pre>{{ config | tojson(indent=2) }}</pre>
</div>
<div class="form-section">
    <h3>信息</h3>
    <pre>{{ info | tojson(indent=2) }}</pre>
</div>
<a href="{{ url_for('match.index') }}" class="btn">返回</a>
{% endblock %}
    ```

- **templates/error.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}错误{% endblock %}
{% block content %}
<h2>⚠️ 出错了</h2>
<div class="message error" style="display:block; margin-bottom:24px;">{{ error }}</div>
<a href="{{ url_for('index') }}" class="btn">← 返回首页</a>
{% endblock %}
    ```

- **templates/field_parser/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}字段映射{% endblock %}
{% block head %}
<style>
    .parser-container { display: flex; gap: 20px; margin-bottom: 20px; }
    .parser-box { flex: 1; }
    .parser-box textarea { width: 100%; height: 200px; }
    .mapping-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    .mapping-table th, .mapping-table td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    .mapping-table .score { font-weight: bold; }
</style>
{% endblock %}
{% block content %}
<h2>🔗 字段映射分析</h2>
<p>粘贴两个文件的列名列表，系统将自动匹配相似字段。列名可使用逗号、空格、换行等分隔。</p>

<form id="parserForm">
    <div class="parser-container">
        <div class="parser-box">
            <h3>标准文件列名</h3>
            <textarea id="textA" name="text_a" placeholder="例如：编号, 姓名, 成绩"></textarea>
        </div>
        <div class="parser-box">
            <h3>比较文件列名</h3>
            <textarea id="textB" name="text_b" placeholder="例如：ID, 名字, 分数"></textarea>
        </div>
    </div>
    <button type="button" class="btn btn-primary" id="btnAnalyze">开始分析</button>
    <button type="button" class="btn btn-secondary" id="btnExport" style="display:none;">📥 导出 Excel</button>
</form>

<div id="resultArea" style="display:none; margin-top: 30px;">
    <h3>分析结果</h3>
    <table class="mapping-table" id="mappingTable">
        <thead><tr><th>标准字段</th><th>比较字段</th><th>相似度</th></tr></thead>
        <tbody></tbody>
    </table>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/field_parser.js') }}"></script>
{% endblock %}
    ```

- **templates/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}首页 - Excel 工具集{% endblock %}
{% block content %}
<h2>👋 欢迎使用 Excel 智能工具集</h2>
<p style="font-size: 1.1rem; margin-bottom: 32px;">请从导航栏选择所需功能：</p>

<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 24px;">
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🧮 列计算/聚合回填</h3>
        <p>支持多列表达式计算，或按分组聚合回填结果列，保留原有格式。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔍 文件对比</h3>
        <p>完整对比（主键匹配）或仅字段值匹配，输出差异报告。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔗 匹配填充</h3>
        <p>根据源文件匹配目标文件，填充值并保留格式。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">📋 汇总工具</h3>
        <p>合并多个文件，按唯一键汇总求和。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">✂️ 文件拆分</h3>
        <p>按列值拆分Excel为多个文件，保留合并单元格与样式。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔗 字段映射</h3>
        <p>智能匹配不同文件间的列名。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">✅ 文件校验</h3>
        <p>根据主键对比两个文件群，输出差异报告。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔄 格式转换</h3>
        <p>Excel 与 CSV 互转，支持编码选择。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">📊 数据小计</h3>
        <p>按指定主键列分组，在每组下方插入小计行，计算合计、平均值等。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">📊 车队日报</h3>
        <p>生成车队业绩日报，含上线率、目标完成率。</p>
    </div>
    <div class="form-section" style="margin:0;">
        <h3 style="margin-top:0;">🔄 行列转置</h3>
        <p>长表转宽表，取消分类列。</p>
    </div>
</div>
{% endblock %}
    ```

- **templates/matcher/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}匹配填充{% endblock %}
{% block head %}
<style>
    .button-group { display: flex; gap: 16px; margin-top: 24px; }
    .debug-btn { background: var(--warning); color: var(--dark); }
</style>
{% endblock %}
{% block content %}
<h2>🔗 Excel 数据匹配填充</h2>
<form id="matchForm" method="post" action="{{ url_for('match.process') }}" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>🎯 目标文件 (保留格式)</label>
            <input type="file" name="target_file" accept=".xlsx" required>
            <div class="hint">将保留原有合并单元格、字体等格式</div>
        </div>
        <div class="form-row">
            <label>📁 源文件 (数据源)</label>
            <input type="file" name="source_file" accept=".xlsx,.xls" required>
            <div class="hint">支持 .xlsx 和 .xls</div>
        </div>
        <div class="form-row">
            <label>源文件 Sheet 索引 (可选)</label>
            <input type="text" name="source_sheets" placeholder="例如 1,2,3">
            <div class="hint">不填默认读取第一个sheet</div>
        </div>
    </div>
    <div class="form-section">
        <div class="form-row">
            <label>目标文件键列 (字母，多列逗号分隔)</label>
            <input type="text" name="target_key_cols" placeholder="A,B" required>
        </div>
        <div class="form-row">
            <label>源文件匹配列 (字母，多列逗号分隔)</label>
            <input type="text" name="source_match_cols" placeholder="A,B" required>
        </div>
        <div class="form-row">
            <label>源文件值列 (字母)</label>
            <input type="text" name="source_value_col" placeholder="C" required>
        </div>
        <div class="form-row">
            <label>结果列 (目标文件)</label>
            <input type="text" name="result_col" placeholder="K 或 '匹配值'" required>
            <div class="hint">输入列字母覆盖该列；输入其他新建列</div>
        </div>
        <div class="form-row">
            <label>目标键重复时的处理方式</label>
            <select name="duplicate_action" id="duplicateAction">
                <option value="error" selected>报错（检测到重复键时中止）</option>
                <option value="first">只填充第一个匹配行（其余行保留原值）</option>
                <option value="all">填充所有匹配行（每行都写入相同值）</option>
            </select>
            <div class="hint">
                仅当目标文件的主键列存在重复值时生效。
                "只填充第一个"适合一个 key 只对应一条业务的场景；
                "填充所有"适合一个 key 对应多条业务记录的场景。
            </div>
        </div>
    </div>
    <div class="button-group">
        <button type="submit" class="btn btn-success">✨ 开始匹配填充</button>
        <button type="button" id="debugBtn" class="btn debug-btn">🔍 调试</button>
    </div>
</form>
<script>
    document.getElementById('debugBtn').addEventListener('click', function() {
        const form = document.getElementById('matchForm');
        form.action = '{{ url_for("match.debug") }}';
        form.target = '_blank';
        form.submit();
        form.action = '{{ url_for("match.process") }}';
        form.target = '';
    });
</script>
{% endblock %}
    ```

- **templates/splitter/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}Excel拆分工具{% endblock %}
{% block content %}
<h2>✂️ Excel拆分工具（保留合并单元格）</h2>
<p>根据指定列的唯一值将Excel拆分成多个文件，保留原格式。现在支持先筛选后拆分。</p>
<form id="splitForm" method="POST" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📂 Excel文件</label>
            <input type="file" name="file" accept=".xlsx,.xlsm,.xls" required>
            <div class="hint">支持 .xlsx, .xlsm, .xls</div>
        </div>
        <div class="form-row">
            <label>拆分列名</label>
            <input type="text" name="split_column" placeholder="例如：险种类型" required>
            <div class="hint">必须是第一行表头中的列名</div>
        </div>
        <div class="form-row">
            <label>正选值（可选）</label>
            <input type="text" name="include_values" placeholder="居民,职工">
            <div class="hint">只拆分这些值</div>
        </div>
        <div class="form-row">
            <label>反选值（可选）</label>
            <input type="text" name="exclude_values" placeholder="其他">
            <div class="hint">排除这些值</div>
        </div>
        <div class="form-row">
            <label>跳过工作表索引</label>
            <input type="text" name="skip_sheets" placeholder="1,3">
            <div class="hint">不参与拆分的sheet序号（从1开始）</div>
        </div>
    </div>

    <!-- 筛选条件区域 -->
    <div class="form-section">
        <h3>🎯 筛选条件（可选）</h3>
        <div id="advancedConditions">
            <div class="condition-row" style="display:flex; gap:8px; margin-bottom:10px;">
                <input type="text" name="condition_field_0" placeholder="字段名" style="flex:2;">
                <select name="condition_op_0" style="flex:1.5;">
                    <option value="=">=</option>
                    <option value="!=">!=</option>
                    <option value=">">></option>
                    <option value="<"><</option>
                    <option value=">=">>=</option>
                    <option value="<="><=</option>
                    <option value="包含">包含</option>
                    <option value="不包含">不包含</option>
                </select>
                <input type="text" name="condition_value_0" placeholder="值" style="flex:3;">
                <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">✖</button>
            </div>
        </div>
        <button type="button" class="btn btn-secondary" onclick="addCondition()">+ 添加条件</button>
        <div class="hint">先按这些条件筛选数据，再进行拆分。所有条件为 AND 关系。</div>
    </div>

    <button type="submit" class="btn btn-success" id="submitBtn">开始拆分</button>
</form>

<div id="statusArea" class="status-area" style="display:none;">
    <h3>处理状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">处理中...</span></p>
</div>
<script>
    let condCount = 1;
    function addCondition() {
        const container = document.getElementById('advancedConditions');
        const div = document.createElement('div');
        div.className = 'condition-row';
        div.style.cssText = 'display:flex; gap:8px; margin-bottom:10px;';
        div.innerHTML = `
            <input type="text" name="condition_field_${condCount}" placeholder="字段名" style="flex:2;">
            <select name="condition_op_${condCount}" style="flex:1.5;">
                <option value="=">=</option>
                <option value="!=">!=</option>
                <option value=">">></option>
                <option value="<"><</option>
                <option value=">=">>=</option>
                <option value="<="><=</option>
                <option value="包含">包含</option>
                <option value="不包含">不包含</option>
            </select>
            <input type="text" name="condition_value_${condCount}" placeholder="值" style="flex:3;">
            <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">✖</button>
        `;
        container.appendChild(div);
        condCount++;
    }

    const form = document.getElementById('splitForm');
    const statusArea = document.getElementById('statusArea');
    const taskStatus = document.getElementById('taskStatus');
    const progressBar = document.getElementById('progressBar');
    const submitBtn = document.getElementById('submitBtn');

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        statusArea.style.display = 'block';
        taskStatus.innerText = '提交中...';
        progressBar.style.width = '10%';
        submitBtn.disabled = true;

        const formData = new FormData(form);
        try {
            const response = await fetch('/splitter/process', {
                method: 'POST',
                body: formData
            });
            if (response.ok) {
                const contentDisposition = response.headers.get('Content-Disposition');
                let filename = '拆分结果.zip';
                if (contentDisposition) {
                    const match = contentDisposition.match(/filename="?(.+?)"?$/);
                    if (match) filename = match[1];
                }
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                taskStatus.innerText = '下载完成！';
                progressBar.style.width = '100%';
                progressBar.style.backgroundColor = '#28a745';
                setTimeout(() => {
                    statusArea.style.display = 'none';
                }, 3000);
            } else {
                let errorMsg = '拆分失败';
                try {
                    const errData = await response.json();
                    errorMsg = errData.message || errorMsg;
                } catch (e) {
                    errorMsg = `服务器错误 (${response.status})`;
                }
                taskStatus.innerText = '错误: ' + errorMsg;
                progressBar.style.width = '100%';
                progressBar.style.backgroundColor = '#dc3545';
            }
        } catch (err) {
            taskStatus.innerText = '网络错误: ' + err.message;
            progressBar.style.width = '100%';
            progressBar.style.backgroundColor = '#dc3545';
        } finally {
            submitBtn.disabled = false;
        }
    });
</script>
{% endblock %}
    ```

- **templates/subtotal/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}数据小计{% endblock %}
{% block head %}
<style>
    .override-row { display: flex; gap: 10px; margin-bottom: 8px; align-items: center; }
    .sheets-radio { margin-right: 15px; }
</style>
{% endblock %}
{% block content %}
<h2>📊 数据小计</h2>
<p>根据主键列对数据进行分组，在每组最后一行下方插入小计行。</p>
<form id="subtotalForm" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 上传 Excel 文件</label>
            <input type="file" name="file" accept=".xlsx,.xls" required>
        </div>
    </div>
    <div class="form-section">
        <h3>🔑 主键列（分组依据，支持多列和范围，如 A,B,C 或 A-C）</h3>
        <input type="text" name="key_cols" placeholder="A,B" required>
    </div>
    <div class="form-section">
        <h3>📏 维度列（在小计行显示“合计”，可选，支持多列）</h3>
        <input type="text" name="dimension_cols" placeholder="C,D">
        <div class="hint">如果指定的列不存在，将报错。</div>
    </div>
    <div class="form-section">
        <h3>🧮 计算列（支持多列和范围）</h3>
        <input type="text" name="calc_cols" placeholder="E,F,G-H" required>
    </div>
    <div class="form-section">
        <h3>⚙️ 聚合函数</h3>
        <div style="display: flex; align-items: center; gap: 10px;">
            <label>默认函数：</label>
            <select name="agg_func_default" id="aggFuncDefault">
                <option value="sum">求和</option>
                <option value="avg">平均值</option>
                <option value="min">最小值</option>
                <option value="max">最大值</option>
                <option value="count">计数</option>
            </select>
            <button type="button" class="btn btn-secondary" id="addOverrideBtn">+ 特殊列函数</button>
        </div>
        <div id="overrideContainer" style="margin-top:10px;"></div>
        <div class="hint">点击“+”为特定列设置不同的聚合函数，格式：列字母:函数，如 C:avg</div>
    </div>
    <div class="form-section">
        <h3>📄 处理Sheet</h3>
        <div>
            <label class="sheets-radio"><input type="radio" name="sheets_mode" value="1" checked> 仅第一个Sheet</label>
            <label class="sheets-radio"><input type="radio" name="sheets_mode" value="all"> 所有Sheet</label>
            <label class="sheets-radio"><input type="radio" name="sheets_mode" value="custom"> 指定Sheet索引（逗号分隔）</label>
            <input type="text" name="sheets_custom" id="sheetsCustom" placeholder="1,2,3" style="width:150px; display:none;">
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始处理</button>
</form>

<div id="statusArea" style="display:none; margin-top:20px;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <div id="downloadLink"></div>
</div>

<template id="overrideRowTemplate">
    <div class="override-row">
        <span>列：</span>
        <input type="text" class="override-col" placeholder="C" style="width:80px;">
        <span>函数：</span>
        <select class="override-func">
            <option value="sum">求和</option>
            <option value="avg">平均值</option>
            <option value="min">最小值</option>
            <option value="max">最大值</option>
            <option value="count">计数</option>
        </select>
        <button type="button" class="btn btn-danger" onclick="this.parentElement.remove()">删除</button>
    </div>
</template>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/subtotal.js') }}"></script>
{% endblock %}
    ```

- **templates/summarizer/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}汇总工具{% endblock %}
{% block content %}
<h2>📋 汇总工具</h2>
<form action="{{ url_for('summary.process') }}" method="post" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📂 上传多个Excel文件</label>
            <input type="file" name="files" accept=".xlsx,.xls" multiple required>
        </div>
        <div class="form-row">
            <label>需要保留的字段（每行一个，或逗号/空格分隔）</label>
            <textarea name="fields" rows="5" placeholder="机构名称&#10;日期&#10;金额" required></textarea>
        </div>
        <div class="form-row">
            <label>唯一键（主键，逗号、换行或空格分隔，可选）</label>
            <textarea name="unique_keys" rows="3" placeholder="例如 机构名称,日期"></textarea>
        </div>
        <div class="form-row">
            <label>求和字段（逗号、换行或空格分隔，可选）</label>
            <textarea name="sum_fields" rows="3" placeholder="例如 金额,数量"></textarea>
        </div>
        <div class="form-row">
            <label>表头所在行号（从1开始，如有标题行请填写；留空则从第一行读取）</label>
            <input type="number" name="header_row" value="" min="1" placeholder="例如 2" style="max-width:120px;">
            <div class="hint">与下方“跳过标题行数”二选一，优先使用行号</div>
        </div>
        <div class="form-row">
            <label>跳过的标题行数（第一行如果不是表头，请填写跳过行数）</label>
            <input type="number" name="skip_header_rows" value="0" min="0" style="max-width:120px;">
            <div class="hint">例如：文件第一行是标题，第二行是表头，则填写 1</div>
        </div>
    </div>
    <button type="submit" class="btn btn-success">开始汇总</button>
</form>
{% endblock %}
    ```

- **templates/team_report/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}车队日报{% endblock %}
{% block head %}
<style>
    .form-row { margin-bottom: 15px; }
    .form-row label { display: block; margin-bottom: 5px; font-weight: bold; }
    .form-row input[type="file"], .form-row input[type="text"], .form-row input[type="number"] { width: 100%; max-width: 300px; }
    .hint { color: #666; font-size: 0.9em; }
    .status-area { margin-top: 20px; }
    .progress { height: 20px; background: #eee; border-radius: 10px; overflow: hidden; }
    .progress-bar { height: 100%; background: #4361ee; width: 0%; transition: width 0.3s; }
    .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
</style>
{% endblock %}
{% block content %}
<h2>📊 生成车队日报</h2>
<form id="reportForm" enctype="multipart/form-data">
    <div class="form-section">
        <h3>📂 上传文件</h3>

        <div class="form-row">
            <label>今日司机流水文件 <span style="color:red;">*</span></label>
            <input type="file" id="todayFile" name="today_file" accept=".xlsx,.xls,.csv" required>
        </div>

        <div class="form-row">
            <label>昨日司机流水文件 （可选）</label>
            <input type="file" id="yesterdayFile" name="yesterday_file" accept=".xlsx,.xls,.csv">
        </div>

        <div class="form-row">
            <label>本月司机流水文件（可选，多选）</label>
            <input type="file" id="monthlyFiles" name="monthly_files" accept=".xlsx,.xls,.csv" multiple>
            <div class="hint">可上传多个月份的流水文件，用于统计月度累计数据</div>
        </div>

        <div class="form-row">
            <label>队伍明细映射表 （可选）</label>
            <input type="file" id="mappingFile" name="mapping_file" accept=".xlsx,.xls">
            <div class="hint">若不提供，将使用已配置的默认映射</div>
        </div>
    </div>

    <div class="form-section">
        <h3>🎯 目标值设置（留空使用默认值）</h3>
        <div class="grid-2">
            <div class="form-row">
                <label>日完单目标</label>
                <input type="number" id="dailyOrderTarget" name="daily_order_target" placeholder="默认1150" min="1">
            </div>
            <div class="form-row">
                <label>日报单目标</label>
                <input type="number" id="dailyApplyTarget" name="daily_apply_target" placeholder="默认120" min="1">
            </div>
            <div class="form-row">
                <label>月完单目标</label>
                <input type="number" id="monthlyOrderTarget" name="monthly_order_target" placeholder="默认35650" min="1">
            </div>
            <div class="form-row">
                <label>月报单目标</label>
                <input type="number" id="monthlyApplyTarget" name="monthly_apply_target" placeholder="默认3720" min="1">
            </div>
        </div>
    </div>

    <button type="submit" class="btn btn-success" id="submitBtn">开始生成</button>
</form>

<div id="statusArea" class="status-area" style="display:none;">
    <h3>任务状态</h3>
    <div class="progress"><div id="progressBar" class="progress-bar"></div></div>
    <p><strong>状态：</strong> <span id="taskStatus">等待中</span></p>
    <p id="progressMessage"></p>
    <div id="downloadLink" style="margin-top:15px;"></div>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/team_report.js') }}"></script>
{% endblock %}
    ```

- **templates/transposer/index.html**
  - language: `html`
  - content:
    ```html
{% extends "base.html" %}
{% block title %}行列转置{% endblock %}
{% block content %}
<h2>🔄 行列转置（取消分类列）</h2>
<p>将某一分类列的不同取值转换为列标题，实现长表转宽表。例如：医疗类别（住院、门诊、其他）变为列头，原数值列变为“住院a数据”、“门诊b数据”等。</p>

<form id="transposerForm" enctype="multipart/form-data">
    <div class="form-section">
        <div class="form-row">
            <label>📁 上传 Excel 文件</label>
            <input type="file" name="file" accept=".xlsx,.xls" required>
        </div>
    </div>
    <div class="form-section">
        <div class="form-row">
            <label>🔑 索引列（主键，多列用逗号或空格分隔，可选）</label>
            <input type="text" name="index_cols" placeholder="例如 病种名称 单位名称，或 A B">
            <div class="hint">若不填，将自动为每组类别生成行序号，并以此为索引。</div>
        </div>
        <div class="form-row">
            <label>📋 分类列（将被取消的列）</label>
            <input type="text" name="category_col" placeholder="例如 医疗类别" required>
        </div>
        <div class="form-row">
            <label>📊 值列（需保留的数值列，多选）</label>
            <input type="text" name="value_cols" placeholder="例如 a数据 b数据" required>
            <div class="hint">多个列名用逗号或空格分隔。</div>
        </div>
    </div>
    <button type="submit" class="btn btn-success" id="submitBtn">开始转换</button>
</form>

<div id="statusArea" class="status-area" style="display:none;">
    <p id="statusMessage"></p>
</div>
{% endblock %}
{% block scripts %}
<script src="{{ url_for('static', filename='js/transposer.js') }}"></script>
{% endblock %}
    ```

- **tests/__init__.py**
  - language: `python`
  - content:
    ```python

    ```

- **tests/conftest.py**
  - language: `python`
  - content:
    ```python
import pytest
import tempfile
import os
import shutil
from openpyxl import Workbook
import pandas as pd
from app import create_app


@pytest.fixture
def app():
    """每个测试函数独立创建一个 Flask 应用和临时目录，并在结束后清理。"""
    app = create_app()
    upload = tempfile.mkdtemp()
    app.config.update({
        "TESTING": True,
        "UPLOAD_FOLDER": upload,
        "WTF_CSRF_ENABLED": False,
    })
    yield app
    # 清理所有临时目录
    shutil.rmtree(upload, ignore_errors=True)


@pytest.fixture
def client(app):
    return app.test_client()


@pytest.fixture
def sample_excel_file():
    from io import BytesIO
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "A"
    ws["B1"] = "B"
    ws["C1"] = "C"
    ws["A2"] = 10
    ws["B2"] = 20
    ws["C2"] = 30
    ws["A3"] = 5
    ws["B3"] = 3
    ws["C3"] = 2
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@pytest.fixture
def sample_csv_header_file():
    from io import BytesIO
    content = "key,value,other\n"
    return BytesIO(content.encode("utf-8"))


@pytest.fixture
def sample_csv_data_file():
    from io import BytesIO
    content = "A,10,x\nA,20,y\nB,5,z\n"
    return BytesIO(content.encode("utf-8"))


@pytest.fixture
def sample_excel_comparison_files():
    from io import BytesIO
    df1 = pd.DataFrame({"id": [1, 2], "name": ["Alice", "Bob"], "score": [100, 90]})
    df2 = pd.DataFrame({"id": [1, 3], "name": ["Alice", "Charlie"], "score": [100, 80]})
    f1 = BytesIO()
    f2 = BytesIO()
    with pd.ExcelWriter(f1, engine="openpyxl") as w: df1.to_excel(w, index=False)
    with pd.ExcelWriter(f2, engine="openpyxl") as w: df2.to_excel(w, index=False)
    f1.seek(0); f2.seek(0)
    return f1, f2
    ```

- **tests/test_calc_no_double_normalize.py**
  - language: `python`
  - content:
    ```python
# tests/test_calc_no_double_normalize.py
"""P0-2 回归测试：calc 模块不应重复 normalize。

覆盖：
- already_normalized=True 时，processor 不调用 normalize_excel_file
- already_normalized=False（默认）时，processor 调用 normalize_excel_file
- already_normalized=True 时，原路径不会进入 temp_files（避免误删）
- already_normalized=True 且路径不是 .xlsx 时，回退到 normalize
- 端到端：通过蓝图调用时，normalize_excel_file 仅被调用一次
"""
import os

import pytest
from openpyxl import Workbook

import services.excel.calculator_processor as proc_module
from services.excel.calculator_processor import (
    ColumnCalculatorProcessor,
    AggregateBackfillProcessor,
)
from services.excel.compute_engine import ComputeEngine


# ---------------------------------------------------------------------------
# 工具
# ---------------------------------------------------------------------------
def _make_xlsx(path, n_rows: int = 3, n_cols: int = 3):
    wb = Workbook()
    ws = wb.active
    ws.append([f"c{c+1}" for c in range(n_cols)])
    for r in range(1, n_rows):
        ws.append([r * (c + 1) for c in range(n_cols)])
    wb.save(str(path))
    wb.close()
    return str(path)


class _NormalizeCounter:
    """包装 normalize_excel_file，统计调用次数。"""
    def __init__(self, monkeypatch):
        self.count = 0
        self.calls = []
        self._original = proc_module.normalize_excel_file

        def wrapper(*args, **kwargs):
            self.count += 1
            self.calls.append((args, kwargs))
            return self._original(*args, **kwargs)

        monkeypatch.setattr(
            proc_module, "normalize_excel_file", wrapper
        )


# ---------------------------------------------------------------------------
# BaseExcelProcessor._get_work_filepath 的行为
# ---------------------------------------------------------------------------
class TestGetWorkFilepathBehavior:

    def test_already_normalized_skips_normalize(self, tmp_path, monkeypatch):
        """already_normalized=True 时，normalize_excel_file 不应被调用。"""
        src = _make_xlsx(tmp_path / "input.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=src,
            header_row=1,
            expression="A+B",
            result_col_input="D",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        result = p._get_work_filepath()

        assert result == src
        assert counter.count == 0, (
            f"already_normalized=True 时不应调用 normalize_excel_file，"
            f"实际调用 {counter.count} 次"
        )

    def test_default_calls_normalize(self, tmp_path, monkeypatch):
        """默认（already_normalized=False）时，应调用 normalize_excel_file。"""
        src = _make_xlsx(tmp_path / "input.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=src,
            header_row=1,
            expression="A+B",
            result_col_input="D",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
        )
        result = p._get_work_filepath()

        assert counter.count == 1
        # P0-1 之后有效 .xlsx 会返回原路径
        assert result == src

    def test_already_normalized_path_not_in_temp_files(self, tmp_path):
        """already_normalized=True 时，原路径不应进入 temp_files（避免误删）。"""
        src = _make_xlsx(tmp_path / "input.xlsx")

        p = ColumnCalculatorProcessor(
            filepath=src,
            header_row=1,
            expression="A+B",
            result_col_input="D",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        result = p._get_work_filepath()

        assert result == src
        assert src not in p.temp_files, (
            "原路径不应进入 temp_files，否则 cleanup 会误删原文件"
        )

    def test_already_normalized_non_xlsx_falls_back(self, tmp_path, monkeypatch):
        """already_normalized=True 但扩展名不是 .xlsx → 回退到 normalize。"""
        # 故意造一个 .xls 文件
        src = tmp_path / "input.xls"
        wb = Workbook()
        wb.active.append(["A"])
        wb.active.append([1])
        wb.save(str(src))
        wb.close()

        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=str(src),
            header_row=1,
            expression="A+1",
            result_col_input="B",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        p._get_work_filepath()

        assert counter.count == 1, (
            "非 .xlsx 文件即使声明 already_normalized=True，也应回退到 normalize"
        )

    def test_already_normalized_result_is_cached(self, tmp_path, monkeypatch):
        """重复调用 _get_work_filepath 不应重复 normalize。"""
        src = _make_xlsx(tmp_path / "input.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=src,
            header_row=1,
            expression="A+B",
            result_col_input="D",
            merge_by_cols_str="",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        r1 = p._get_work_filepath()
        r2 = p._get_work_filepath()

        assert r1 == r2 == src
        assert counter.count == 0


# ---------------------------------------------------------------------------
# 两个 Processor 子类都支持 already_normalized
# ---------------------------------------------------------------------------
class TestProcessorSubclasses:

    def test_column_processor_accepts_flag(self, tmp_path, monkeypatch):
        src = _make_xlsx(tmp_path / "c.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = ColumnCalculatorProcessor(
            filepath=src, header_row=1,
            expression="A+B", result_col_input="D",
            merge_by_cols_str="", compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        assert p.already_normalized is True
        p._get_work_filepath()
        assert counter.count == 0

    def test_aggregate_processor_accepts_flag(self, tmp_path, monkeypatch):
        src = _make_xlsx(tmp_path / "a.xlsx")
        counter = _NormalizeCounter(monkeypatch)

        p = AggregateBackfillProcessor(
            filepath=src, header_row=1,
            group_cols_raw=["A"], agg_cols_raw=["B"],
            agg_func="sum", result_col_input="D",
            compute_engine=ComputeEngine(),
            already_normalized=True,
        )
        assert p.already_normalized is True
        p._get_work_filepath()
        assert counter.count == 0

    def test_default_flag_is_false(self, tmp_path):
        src = _make_xlsx(tmp_path / "d.xlsx")
        p = ColumnCalculatorProcessor(
            filepath=src, header_row=1,
            expression="A+B", result_col_input="D",
            merge_by_cols_str="", compute_engine=ComputeEngine(),
        )
        assert p.already_normalized is False


# ---------------------------------------------------------------------------
# 端到端：通过蓝图触发
# ---------------------------------------------------------------------------
class TestEndToEndViaBlueprint:
    def test_blueprint_triggers_exactly_one_normalize(
        self, client, sample_excel_file, monkeypatch
    ):
        """通过 /calc/process 调用，normalize_excel_file 应恰好被调用 1 次。

        蓝图层调用 1 次；processor 层因 already_normalized=True 应跳过。
        """
        import core.excel_helpers as helpers

        counter = {"n": 0}
        original = helpers.normalize_excel_file

        def counting(*args, **kwargs):
            counter["n"] += 1
            return original(*args, **kwargs)

        monkeypatch.setattr(helpers, "normalize_excel_file", counting)

        data = {
            "file": (sample_excel_file, "test.xlsx"),
            "calc_type": "column",
            "header_row": "1",
            "expression": "A+B",
            "result_col_column": "D",
            "merge_by_cols": "",
        }
        response = client.post(
            "/calc/process", data=data, content_type="multipart/form-data"
        )
        assert response.status_code == 200, (
            f"请求失败: {response.status_code}"
        )
        assert counter["n"] == 1, (
            f"normalize_excel_file 应被调用 1 次（蓝图），实际 {counter['n']} 次"
        )
    ```

- **tests/test_calculator.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
from openpyxl import Workbook
from services.excel.calculator import ColumnCalculator
from services.excel.aggregator import Aggregator
from core.excel_utils import ExcelFile


class TestColumnCalculator:
    def test_calculate_column(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'A'
        ws['B1'] = 'B'
        ws['C1'] = 'Result'
        ws['A2'] = 10
        ws['B2'] = 20
        ws['A3'] = 5
        ws['B3'] = 3
        test_file = tmp_path / 'test.xlsx'
        wb.save(test_file)

        with ExcelFile(test_file) as ef:
            calculator = ColumnCalculator(ef)
            calculator.calculate_column('A+B', 'C', merge_by_cols=None)
            ef.save(test_file)

        result_df = pd.read_excel(test_file)
        assert result_df.loc[0, 'Result'] == 30
        assert result_df.loc[1, 'Result'] == 8

    def test_calculate_column_with_merge(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = 'Group'
        ws['B1'] = 'Value'
        ws['C1'] = 'Result'
        ws['A2'] = 'X'
        ws['B2'] = 10
        ws['A3'] = 'X'
        ws['B3'] = 20
        ws['A4'] = 'Y'
        ws['B4'] = 5
        test_file = tmp_path / 'test.xlsx'
        wb.save(test_file)

        with ExcelFile(test_file) as ef:
            calculator = ColumnCalculator(ef)
            calculator.calculate_column('B*2', 'C', merge_by_cols=['A'])
            ef.save(test_file)

        result_df = pd.read_excel(test_file)
        assert result_df.loc[0, 'Result'] == 20
        assert pd.isna(result_df.loc[1, 'Result'])
        assert result_df.loc[2, 'Result'] == 10


class TestAggregator:
    def test_aggregate_sum(self):
        df = pd.DataFrame({'Group': ['X', 'X', 'Y'], 'Value': [10, 20, 5]})
        result = Aggregator.simple_aggregate(df, ['Group'], ['Value'], 'sum')
        assert len(result) == 2
        assert result.loc[result['Group'] == 'X', 'Value'].iloc[0] == 30
        assert result.loc[result['Group'] == 'Y', 'Value'].iloc[0] == 5

    def test_aggregate_mean(self):
        df = pd.DataFrame({'Group': ['X', 'X'], 'Value': [10, 30]})
        result = Aggregator.simple_aggregate(df, ['Group'], ['Value'], 'mean')
        assert result.loc[0, 'Value'] == 20

    def test_aggregate_by_keys(self):
        df = pd.DataFrame({
            'key': ['A', 'A', 'B'],
            'sum_col': [10, 20, 5],
            'other_col': ['first', 'second', 'other']
        })
        result = Aggregator.aggregate_by_keys(df, ['key'], ['sum_col'], ['other_col'])
        assert len(result) == 2
        assert result.loc[result['key'] == 'A', 'sum_col'].iloc[0] == 30
        assert result.loc[result['key'] == 'A', 'other_col'].iloc[0] == 'first'
        assert result.loc[result['key'] == 'B', 'sum_col'].iloc[0] == 5
    ```

- **tests/test_checker_service.py**
  - language: `python`
  - content:
    ```python
import pytest
import pandas as pd
import os
from services.checker.service import CheckerService
from core.exceptions import ValidationError
from core.config import CheckConfig  # 引入配置类


class TestCheckerService:
    @pytest.fixture
    def sample_files_a(self, tmp_path):
        df1 = pd.DataFrame({"ID": [1, 2], "Name": ["Alice", "Bob"], "Score": [100, 90]})
        df2 = pd.DataFrame({"ID": [3, 4], "Name": ["Charlie", "David"], "Score": [80, 70]})
        f1 = tmp_path / "a1.xlsx"
        f2 = tmp_path / "a2.xlsx"
        df1.to_excel(f1, index=False)
        df2.to_excel(f2, index=False)
        return [str(f1), str(f2)]

    @pytest.fixture
    def sample_files_b(self, tmp_path):
        df = pd.DataFrame({"ID": [1, 2, 5], "Name": ["Alice", "Bob", "Eve"], "Score": [100, 95, 60]})
        f = tmp_path / "b.xlsx"
        df.to_excel(f, index=False)
        return [str(f)]

    def test_execute_basic(self, sample_files_a, sample_files_b):
        service = CheckerService(
            files_a=sample_files_a,
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}],
            duplicate_action="error"
        )
        result_path = service.execute()
        assert os.path.exists(result_path)
        result = pd.read_excel(result_path, sheet_name=None)
        assert "差异数据" in result
        assert len(result["差异数据"]) == 1
        assert "标准群独有" in result
        assert len(result["标准群独有"]) == 2
        assert "比较群独有" in result
        assert len(result["比较群独有"]) == 1
        os.unlink(result_path)

    def test_duplicate_key_error(self, sample_files_a, sample_files_b, tmp_path):
        df_dup = pd.DataFrame({"ID": [1, 1], "Name": ["Alice", "Alice"], "Score": [100, 200]})
        dup_file = tmp_path / "dup.xlsx"
        df_dup.to_excel(dup_file, index=False)
        service = CheckerService(
            files_a=[str(dup_file)],
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}],
            duplicate_action="error"
        )
        with pytest.raises(ValidationError, match="标准群主键重复"):
            service.execute()

    def test_filter_in_standard(self, sample_files_a, sample_files_b):
        service = CheckerService(
            files_a=sample_files_a,
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}],
            filters_a=[{"field": "Score", "op": ">", "value": 90}],
            duplicate_action="error"
        )
        result_path = service.execute()
        result = pd.read_excel(result_path, sheet_name=None)
        assert "比较群独有" in result
        assert len(result["比较群独有"]) == 2
        os.unlink(result_path)

    def test_execute_with_sqlite_fallback(self, sample_files_a, sample_files_b, monkeypatch):
        """将 max_memory_rows 设为 0 强制使用 SQLite 回退模式"""
        # 构造一个配置使 max_memory_rows=0，触发 SQLite 路径
        low_mem_config = CheckConfig(max_memory_rows=0)
        service = CheckerService(
            files_a=sample_files_a,
            files_b=sample_files_b,
            key_map_a=["ID"], key_map_b=["ID"],
            compare_map=[{"standard": "Name", "compare": "Name"}, {"standard": "Score", "compare": "Score"}],
            duplicate_action="error",
            config=low_mem_config
        )
        result_path = service.execute()
        assert os.path.exists(result_path)
        result = pd.read_excel(result_path, sheet_name=None)
        assert "差异数据" in result
        os.unlink(result_path)
    ```

- **tests/test_common.py**
  - language: `python`
  - content:
    ```python
import os
import numpy as np
import pandas as pd
from openpyxl import Workbook
from core.validators import validate_column_letters, validate_sheet_index, validate_expression, sanitize_filename
# 直接使用核心表达式引擎，替代旧的 utils.expression 函数
from core.expression_engine import ExpressionEngine
from core.excel_utils import ExcelFile, col_letter_to_index, col_index_to_letter
from services.data.file_service import FileService
from flask import Flask

class TestValidators:
    def test_validate_column_letters(self):
        assert validate_column_letters(['A', 'AB', 'Z']) == True
        assert validate_column_letters(['a', 'Ab']) == True
        assert validate_column_letters(['A1', 'AB']) == False
        assert validate_column_letters([]) == True

    def test_validate_sheet_index(self):
        assert validate_sheet_index('1') == True
        assert validate_sheet_index(2) == True
        assert validate_sheet_index('0') == False
        assert validate_sheet_index('a') == False
        assert validate_sheet_index(None) == False

    def test_validate_expression(self):
        assert validate_expression('A+B-C') == True
        assert validate_expression('A*B/C') == True
        assert validate_expression('A**2') == True
        assert validate_expression('A+') == False
        assert validate_expression('') == False

    def test_sanitize_filename(self):
        assert sanitize_filename('a/b?c:d') == 'a_b_c_d'
        assert sanitize_filename('normal') == 'normal'

class TestExpression:
    @classmethod
    def setup_class(cls):
        cls.engine = ExpressionEngine()

    def test_extract_column_letters(self):
        # 利用引擎的 compile 方法提取变量（字母），但提取逻辑在旧函数中，这里直接复用旧实现
        import re
        def extract_column_letters(expr):
            return set(re.findall(r'[A-Z]+', expr.upper()))
        assert extract_column_letters('A+B-C') == {'A', 'B', 'C'}
        assert extract_column_letters('AA+BB') == {'AA', 'BB'}
        assert extract_column_letters('') == set()

    def test_evaluate_expression(self):
        values = {'A': 10, 'B': 20, 'C': 5}
        assert self.engine.evaluate('A+B-C', values) == 25
        assert self.engine.evaluate('A*B', values) == 200
        assert self.engine.evaluate('A/B', values) == 0.5
        result = self.engine.evaluate('A/0', values)
        assert np.isnan(result)
        result = self.engine.evaluate('X+Y', values)
        assert np.isnan(result)

class TestExcelHandler:
    def test_col_letter_conversion(self):
        assert col_letter_to_index('A') == 1
        assert col_letter_to_index('Z') == 26
        assert col_letter_to_index('AA') == 27
        assert col_index_to_letter(1) == 'A'
        assert col_index_to_letter(27) == 'AA'

    def test_excel_file_read_write(self, tmp_path):
        df = pd.DataFrame({'A': [1, 2], 'B': ['x', 'y']})
        test_file = tmp_path / 'test.xlsx'
        df.to_excel(test_file, index=False)
        with ExcelFile(test_file) as ef:
            headers = ef.get_headers()
            assert headers == ['A', 'B']
            val = ef.get_cell_value(2, 'A')
            assert val == 1
            new_file = tmp_path / 'new.xlsx'
            ef.save(new_file)
        assert new_file.exists()

    def test_copy_merge_structure(self, tmp_path):
        wb = Workbook()
        ws = wb.active
        ws.merge_cells('A1:A3')
        ws['A1'] = 'Merged'
        test_file = tmp_path / 'merge.xlsx'
        wb.save(test_file)
        with ExcelFile(test_file) as ef:
            ws = ef.active_sheet
            ef.copy_merge_structure('A', 'B', 1, 3, sheet=ws)
            merged = list(ws.merged_cells.ranges)
            assert len(merged) == 2

    def test_read_dataframe(self, tmp_path):
        df = pd.DataFrame({'X': [1, 2]})
        test_file = tmp_path / 'data.xlsx'
        df.to_excel(test_file, index=False)
        result = ExcelFile.read_dataframe(test_file)
        pd.testing.assert_frame_equal(result, df)

class TestFileService:
    def test_save_upload(self, tmp_path):
        app = Flask(__name__)
        app.config['UPLOAD_FOLDER'] = str(tmp_path)
        app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}
        with app.app_context():
            from werkzeug.datastructures import FileStorage
            import io
            fake_file = FileStorage(io.BytesIO(b'fake'), filename='test.xlsx')
            svc = FileService()
            path = svc.save_upload(fake_file, subdir='')
            assert os.path.exists(path)
            assert path.endswith('.xlsx')
            svc.cleanup([path])
            assert not os.path.exists(path)
    ```

- **tests/test_comparator.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
import tempfile
from services.excel.comparator import ExcelComparator

def create_test_excel(data, filename):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    return filename

class TestExcelComparator:
    def test_full_compare_identical(self):
        data = [{'id': 1, 'name': 'A', 'score': 100}, {'id': 2, 'name': 'B', 'score': 90}]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            create_test_excel(data, f1.name)
            create_test_excel(data, f2.name)
            result = ExcelComparator.full_compare(f1.name, f2.name, ['id'], ['name', 'score'])
            assert result['only_in_1'].empty
            assert result['only_in_2'].empty
            assert result['diff'] is None

    def test_full_compare_diff(self):
        data1 = [{'id': 1, 'name': 'A', 'score': 100}]
        data2 = [{'id': 1, 'name': 'A', 'score': 90}]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            create_test_excel(data1, f1.name)
            create_test_excel(data2, f2.name)
            result = ExcelComparator.full_compare(f1.name, f2.name, ['id'], ['score'])
            assert result['diff'] is not None
            assert len(result['diff']) == 1
            assert result['diff'].iloc[0]['文件1_score'] == 100
            assert result['diff'].iloc[0]['文件2_score'] == 90

    def test_field_match(self):
        data1 = [{'name': 'A'}, {'name': 'B'}, {'name': 'C'}]
        data2 = [{'name': 'A'}, {'name': 'B'}]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f1, \
             tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f2:
            create_test_excel(data1, f1.name)
            create_test_excel(data2, f2.name)
            only1, only2 = ExcelComparator.field_match(f1.name, f2.name, 'name')
            assert len(only1) == 1
            assert only1.iloc[0]['name'] == 'C'
            assert only2.empty
    ```

- **tests/test_condition_filter.py**
  - language: `python`
  - content:
    ```python
import pandas as pd
from core.filter_engine import ConditionFilter, Condition

class TestConditionFilter:
    def test_apply_to_dict_eq(self):
        row = {'name': 'Alice', 'age': '30'}
        cond = Condition(field='name', op='=', value='Alice')
        assert ConditionFilter.apply_to_dict(row, [cond]) == True

    def test_apply_to_dict_contains(self):
        row = {'name': 'Alice'}
        cond = Condition(field='name', op='contains', value='Ali')
        assert ConditionFilter.apply_to_dict(row, [cond]) == True

    def test_apply_to_dict_exclude(self):
        row = {'name': 'Alice'}
        cond = Condition(field='name', op='=', value='Alice', exclude=True)
        assert ConditionFilter.apply_to_dict(row, [cond]) == False

    def test_apply_to_dataframe(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Charlie'], 'age': [25, 30, 35]})
        cond = Condition(field='name', op='contains', value='Ali')
        mask = ConditionFilter.apply_to_dataframe(df, [cond])
        assert mask.tolist() == [True, False, False]

    def test_apply_to_dataframe_multiple(self):
        df = pd.DataFrame({'name': ['Alice', 'Bob', 'Charlie'], 'age': [25, 30, 35]})
        cond1 = Condition(field='age', op='>', value='30')
        cond2 = Condition(field='name', op='contains', value='Char')
        mask = ConditionFilter.apply_to_dataframe(df, [cond1, cond2])
        assert mask.tolist() == [False, False, True]
    ```

- **tests/test_convert_service.py**
  - language: `python`
  - content:
    ```python
import pytest
import pandas as pd
import tempfile
import os
from services.convert.service import ConvertService
from core.exceptions import ValidationError

class TestConvertService:
    @pytest.fixture
    def sample_xlsx(self, tmp_path):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        path = tmp_path / "test.xlsx"
        df.to_excel(path, index=False)
        return str(path)

    @pytest.fixture
    def sample_multi_sheet_xlsx(self, tmp_path):
        df1 = pd.DataFrame({"ID": [1], "Name": ["Alice"]})
        df2 = pd.DataFrame({"ID": [2], "Name": ["Bob"]})
        path = tmp_path / "multi.xlsx"
        with pd.ExcelWriter(path, engine='openpyxl') as w:
            df1.to_excel(w, sheet_name='Sheet1', index=False)
            df2.to_excel(w, sheet_name='Sheet2', index=False)
        return str(path)

    @pytest.fixture
    def sample_csv(self, tmp_path):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
        path = tmp_path / "test.csv"
        df.to_csv(path, index=False, encoding='utf-8-sig')
        return str(path)

    def test_excel_to_csv_single_sheet(self, sample_xlsx):
        service = ConvertService(file_path=sample_xlsx, target_format='csv')
        result_path = service.execute()
        assert result_path.endswith('.csv')
        df = pd.read_csv(result_path)
        assert list(df.columns) == ['A', 'B']
        assert len(df) == 2
        os.unlink(result_path)

    def test_excel_to_csv_multi_sheet(self, sample_multi_sheet_xlsx):
        service = ConvertService(file_path=sample_multi_sheet_xlsx, target_format='csv')
        result_path = service.execute()
        assert result_path.endswith('.zip')
        import zipfile
        with zipfile.ZipFile(result_path, 'r') as zf:
            names = zf.namelist()
            assert len(names) == 2
            assert any('Sheet1.csv' in n for n in names)
            assert any('Sheet2.csv' in n for n in names)
        os.unlink(result_path)

    def test_csv_to_excel(self, sample_csv):
        service = ConvertService(file_path=sample_csv, target_format='xlsx')
        result_path = service.execute()
        assert result_path.endswith('.xlsx')
        df = pd.read_excel(result_path)
        assert list(df.columns) == ['A', 'B']
        assert len(df) == 2
        os.unlink(result_path)

    def test_unsupported_conversion(self, sample_xlsx):
        service = ConvertService(file_path=sample_xlsx, target_format='xlsx')
        with pytest.raises(ValidationError, match='不支持的转换'):
            service.execute()

    def test_invalid_file(self, tmp_path):
        invalid_file = tmp_path / "invalid.xlsx"
        invalid_file.write_text("not an excel file")
        service = ConvertService(file_path=str(invalid_file), target_format='csv')
        with pytest.raises(Exception):
            service.execute()

    def test_precheck(self, sample_xlsx):
        service = ConvertService(file_path=sample_xlsx, target_format='csv')
        result = service.precheck()
        assert 'total_rows' in result
        assert result['total_rows'] == 2
    ```


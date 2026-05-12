# 视频数据趋势监测应用 — 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建 Flask + Chart.js + XLSX 应用，用户上传两个 .xlsx 数据文件，系统自动解析并持久化，提供可视化趋势分析和视频对比。

**架构:** Flask 后端直接读取/写入 `data_trend_db.xlsx` 作为数据库，前端用 Chart.js CDN 渲染图表，通过 Flask API 获取 JSON 数据。

**Tech Stack:** Python Flask, pandas, openpyxl, Chart.js CDN, HTML + CSS

---

### Task 1: 初始化项目结构和依赖

**Files:**
- Create: `d:\VibeCodingPJ\data_trend\app.py`
- Create: `d:\VibeCodingPJ\data_trend\requirements.txt`
- Create: `d:\VibeCodingPJ\data_trend\uploads\.gitkeep`

- [ ] **Step 1: 创建 `requirements.txt`**

```txt
flask==3.1.1
pandas==2.2.3
openpyxl==3.1.5
```

- [ ] **Step 2: 创建目录结构和 `uploads/.gitkeep`**

```bash
mkdir -p D:\VibeCodingPJ\data_trend\templates
mkdir -p D:\VibeCodingPJ\data_trend\static
mkdir -p D:\VibeCodingPJ\data_trend\uploads
New-Item D:\VibeCodingPJ\data_trend\uploads\.gitkeep -ItemType File -Force
```

- [ ] **Step 3: 安装依赖**

```bash
.venv\Scripts\pip.exe install flask pandas openpyxl
```

---

### Task 2: 创建数据库初始化模块（`db.py`）

**Files:**
- Create: `d:\VibeCodingPJ\data_trend\db.py`

该模块负责 `data_trend_db.xlsx` 的读写操作，作为整个应用的数据层。

- [ ] **Step 1: 创建 `db.py`**

核心功能：
1. `init_db()` — 若 xlsx 不存在则创建 3 个 Sheet 的空骨架
2. `add_video(summary_data, play_trend_df, fan_trend_df)` — 追加一期视频所有数据
3. `get_all_videos()` — 返回所有视频汇总列表（DataFrame）
4. `get_video(video_id)` — 返回单视频汇总 + 趋势数据
5. `delete_video(video_id)` — 删除某视频所有相关行
6. `get_trends_over_time()` — 按发布日期排序的各期汇总数据（用于累计趋势图）
7. `get_next_id()` — 获取下一个自增 video_id

```python
import pandas as pd
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(__file__), 'data_trend_db.xlsx')

SUMMARY_COLUMNS = [
    'video_id', '发布日期', '视频主题', '上传时间',
    'OGC_FID', '播放量', '点赞量', '评论量', '分享量', '收藏量', '弹幕量',
    '完播率', '2s跳出率', '涨粉量', '脱粉量', '粉丝播放占比'
]

TREND_PLAY_COLUMNS = ['OGC_FID', '日期', '播放量', '抖音', '抖音精选']
TREND_FAN_COLUMNS = ['OGC_FID', '日期', '涨粉量', '抖音', '抖音精选']


def init_db():
    if os.path.exists(DB_PATH):
        return
    with pd.ExcelWriter(DB_PATH, engine='openpyxl') as writer:
        pd.DataFrame(columns=SUMMARY_COLUMNS).to_excel(writer, sheet_name='视频汇总', index=False)
        pd.DataFrame(columns=TREND_PLAY_COLUMNS).to_excel(writer, sheet_name='播放量趋势', index=False)
        pd.DataFrame(columns=TREND_FAN_COLUMNS).to_excel(writer, sheet_name='涨粉量趋势', index=False)


def get_next_id():
    if not os.path.exists(DB_PATH):
        return 1
    try:
        df = pd.read_excel(DB_PATH, sheet_name='视频汇总')
        if df.empty or 'video_id' not in df.columns:
            return 1
        return int(df['video_id'].max()) + 1
    except Exception:
        return 1


def add_video(publish_date, topic, ogc_fid, summary_dict, play_trend_df, fan_trend_df):
    init_db()
    video_id = get_next_id()

    summary_row = {
        'video_id': video_id,
        '发布日期': publish_date,
        '视频主题': topic,
        '上传时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'OGC_FID': ogc_fid,
        **summary_dict
    }
    summary_row = {k: summary_row.get(k, None) for k in SUMMARY_COLUMNS}
    new_summary = pd.DataFrame([summary_row])

    existing_summary = pd.read_excel(DB_PATH, sheet_name='视频汇总')
    combined_summary = pd.concat([existing_summary, new_summary], ignore_index=True)

    play_trend_df = play_trend_df.copy()
    play_trend_df['OGC_FID'] = ogc_fid
    play_trend_df = play_trend_df[TREND_PLAY_COLUMNS]

    fan_trend_df = fan_trend_df.copy()
    fan_trend_df['OGC_FID'] = ogc_fid
    fan_trend_df = fan_trend_df[TREND_FAN_COLUMNS]

    existing_play = pd.read_excel(DB_PATH, sheet_name='播放量趋势')
    existing_fan = pd.read_excel(DB_PATH, sheet_name='涨粉量趋势')

    combined_play = pd.concat([existing_play, play_trend_df], ignore_index=True)
    combined_fan = pd.concat([existing_fan, fan_trend_df], ignore_index=True)

    with pd.ExcelWriter(DB_PATH, engine='openpyxl') as writer:
        combined_summary.to_excel(writer, sheet_name='视频汇总', index=False)
        combined_play.to_excel(writer, sheet_name='播放量趋势', index=False)
        combined_fan.to_excel(writer, sheet_name='涨粉量趋势', index=False)

    return video_id


def get_all_videos():
    init_db()
    df = pd.read_excel(DB_PATH, sheet_name='视频汇总')
    if df.empty:
        return []
    df = df.sort_values('发布日期', ascending=False)
    return df.to_dict('records')


def get_video(video_id):
    init_db()
    summary_df = pd.read_excel(DB_PATH, sheet_name='视频汇总')
    video_row = summary_df[summary_df['video_id'] == video_id]
    if video_row.empty:
        return None

    ogc_fid = int(video_row.iloc[0]['OGC_FID'])

    play_trend = pd.read_excel(DB_PATH, sheet_name='播放量趋势')
    play_trend = play_trend[play_trend['OGC_FID'] == ogc_fid].to_dict('records')

    fan_trend = pd.read_excel(DB_PATH, sheet_name='涨粉量趋势')
    fan_trend = fan_trend[fan_trend['OGC_FID'] == ogc_fid].to_dict('records')

    return {
        'summary': video_row.iloc[0].to_dict(),
        'play_trend': play_trend,
        'fan_trend': fan_trend
    }


def delete_video(video_id):
    init_db()
    summary_df = pd.read_excel(DB_PATH, sheet_name='视频汇总')
    target = summary_df[summary_df['video_id'] == video_id]
    if target.empty:
        return False

    ogc_fid = int(target.iloc[0]['OGC_FID'])

    summary_df = summary_df[summary_df['video_id'] != video_id]
    play_df = pd.read_excel(DB_PATH, sheet_name='播放量趋势')
    play_df = play_df[play_df['OGC_FID'] != ogc_fid]
    fan_df = pd.read_excel(DB_PATH, sheet_name='涨粉量趋势')
    fan_df = fan_df[fan_df['OGC_FID'] != ogc_fid]

    with pd.ExcelWriter(DB_PATH, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='视频汇总', index=False)
        play_df.to_excel(writer, sheet_name='播放量趋势', index=False)
        fan_df.to_excel(writer, sheet_name='涨粉量趋势', index=False)

    return True


def get_trends_over_time():
    init_db()
    df = pd.read_excel(DB_PATH, sheet_name='视频汇总')
    if df.empty:
        return []
    df = df.sort_values('发布日期')
    return df.to_dict('records')
```

- [ ] **Step 2: 验证数据库模块能成功初始化**

```bash
cd D:\VibeCodingPJ\data_trend
.venv\Scripts\python.exe -c "from db import init_db; init_db(); print('DB initialized')"
```

---

### Task 3: 创建 Flask 主应用（`app.py`）

**Files:**
- Modify: `d:\VibeCodingPJ\data_trend\app.py`

- [ ] **Step 1: 编写 Flask 应用，包含所有路由**

```python
import os
import pandas as pd
from flask import Flask, request, jsonify, render_template, redirect, url_for

from db import (
    init_db, add_video, get_all_videos, get_video,
    delete_video, get_trends_over_time
)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
init_db()


@app.route('/')
def dashboard():
    return render_template('dashboard.html')


@app.route('/upload')
def upload_page():
    return render_template('upload.html')


@app.route('/video/<int:video_id>')
def detail_page(video_id):
    return render_template('detail.html', video_id=video_id)


@app.route('/api/videos')
def api_videos():
    videos = get_all_videos()
    return jsonify(videos)


@app.route('/api/video/<int:video_id>')
def api_video(video_id):
    data = get_video(video_id)
    if data is None:
        return jsonify({'error': 'Video not found'}), 404
    return jsonify(data)


@app.route('/api/trends')
def api_trends():
    trends = get_trends_over_time()
    return jsonify(trends)


def safe_int(val):
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0


def safe_str(val):
    if pd.isna(val):
        return ''
    return str(val)


@app.route('/api/upload', methods=['POST'])
def api_upload():
    if 'traffic_file' not in request.files or 'fan_file' not in request.files:
        return jsonify({'error': '请同时上传流量数据和粉丝数据两个文件'}), 400

    publish_date = request.form.get('publish_date', '')
    topic = request.form.get('topic', '')

    if not publish_date or not topic:
        return jsonify({'error': '请填写发布日期和视频主题'}), 400

    traffic_file = request.files['traffic_file']
    fan_file = request.files['fan_file']

    if not traffic_file.filename.endswith('.xlsx') or not fan_file.filename.endswith('.xlsx'):
        return jsonify({'error': '请上传 .xlsx 格式的文件'}), 400

    backup_dir = app.config['UPLOAD_FOLDER']
    traffic_path = os.path.join(backup_dir, f"{publish_date}_流量数据.xlsx")
    fan_path = os.path.join(backup_dir, f"{publish_date}_粉丝数据.xlsx")
    traffic_file.save(traffic_path)
    fan_file.save(fan_path)

    try:
        traffic_summary = pd.read_excel(traffic_path, sheet_name='指标数据')
        fan_summary = pd.read_excel(fan_path, sheet_name='指标数据')
        play_trend = pd.read_excel(traffic_path, sheet_name='播放量-新增-每小时趋势数据')
        fan_trend = pd.read_excel(fan_path, sheet_name='涨粉量-新增-每小时趋势数据')
    except Exception as e:
        return jsonify({'error': f'文件解析失败，请检查文件格式: {str(e)}'}), 400

    ogc_fid = safe_int(traffic_summary.iloc[0].get('OGC_FID', 0))

    summary_dict = {
        '播放量': safe_int(traffic_summary.iloc[0].get('播放量', 0)),
        '点赞量': safe_int(traffic_summary.iloc[0].get('点赞量', 0)),
        '评论量': safe_int(traffic_summary.iloc[0].get('评论量', 0)),
        '分享量': safe_int(traffic_summary.iloc[0].get('分享量', 0)),
        '收藏量': safe_int(traffic_summary.iloc[0].get('收藏量', 0)),
        '弹幕量': safe_int(traffic_summary.iloc[0].get('弹幕量', 0)),
        '完播率': safe_str(traffic_summary.iloc[0].get('完播率', '')),
        '2s跳出率': safe_str(traffic_summary.iloc[0].get('2s跳出率', '')),
        '涨粉量': safe_int(fan_summary.iloc[0].get('涨粉量', 0)),
        '脱粉量': safe_int(fan_summary.iloc[0].get('脱粉量', 0)),
        '粉丝播放占比': safe_str(fan_summary.iloc[0].get('粉丝播放占比', '')),
    }

    play_trend = play_trend[play_trend['日期'] != '日期'] if '日期' in play_trend.columns else play_trend
    fan_trend = fan_trend[fan_trend['日期'] != '日期'] if '日期' in fan_trend.columns else fan_trend

    video_id = add_video(publish_date, topic, ogc_fid, summary_dict, play_trend, fan_trend)

    return jsonify({'success': True, 'video_id': video_id, 'topic': topic})


@app.route('/api/video/<int:video_id>/delete', methods=['POST'])
def api_delete_video(video_id):
    success = delete_video(video_id)
    if not success:
        return jsonify({'error': '视频不存在'}), 404
    return jsonify({'success': True})


if __name__ == '__main__':
    app.run(debug=True, port=5000)
```

---

### Task 4: 创建 HTML 模板 — 基础模板和样式

**Files:**
- Create: `d:\VibeCodingPJ\data_trend\templates\base.html`
- Create: `d:\VibeCodingPJ\data_trend\static\style.css`

- [ ] **Step 1: 创建 `base.html`**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{% block title %}视频数据趋势监测{% endblock %}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
    <link rel="stylesheet" href="/static/style.css">
</head>
<body>
    <nav class="nav">
        <div class="nav-inner">
            <a href="/" class="nav-logo">📊 视频趋势监测</a>
            <div class="nav-links">
                <a href="/">概览看板</a>
                <a href="/upload">上传数据</a>
            </div>
        </div>
    </nav>
    <main class="main">
        {% block content %}{% endblock %}
    </main>
</body>
</html>
```

- [ ] **Step 2: 创建 `style.css`**

```css
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f7fa; color: #333; }

.nav { background: #1a1a2e; color: #fff; padding: 0 24px; }
.nav-inner { max-width: 1200px; margin: 0 auto; display: flex; align-items: center; height: 56px; gap: 32px; }
.nav-logo { font-size: 18px; font-weight: 700; color: #fff; text-decoration: none; }
.nav-links { display: flex; gap: 20px; }
.nav-links a { color: #a0aec0; text-decoration: none; font-size: 14px; }
.nav-links a:hover { color: #fff; }

.main { max-width: 1200px; margin: 0 auto; padding: 24px; }

.cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }
.card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
.card-label { font-size: 13px; color: #888; margin-bottom: 6px; }
.card-value { font-size: 28px; font-weight: 700; color: #1a1a2e; }

.chart-container { background: #fff; border-radius: 12px; padding: 20px; margin-bottom: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
.chart-container h2 { font-size: 16px; margin-bottom: 16px; color: #444; }
.chart-container canvas { max-height: 300px; }

.table-wrap { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 14px; }
th { text-align: left; padding: 10px 12px; border-bottom: 2px solid #eee; color: #888; font-weight: 600; font-size: 12px; text-transform: uppercase; }
td { padding: 12px; border-bottom: 1px solid #f0f0f0; }
tr:hover td { background: #fafafa; }
td a { color: #2563eb; text-decoration: none; }
td a:hover { text-decoration: underline; }

.btn { display: inline-block; padding: 8px 16px; border-radius: 8px; border: none; font-size: 14px; cursor: pointer; text-decoration: none; }
.btn-primary { background: #2563eb; color: #fff; }
.btn-primary:hover { background: #1d4ed8; }
.btn-danger { background: #ef4444; color: #fff; }
.btn-danger:hover { background: #dc2626; }
.btn-sm { padding: 4px 10px; font-size: 12px; }

.form-group { margin-bottom: 16px; }
.form-group label { display: block; font-size: 14px; font-weight: 600; margin-bottom: 6px; color: #444; }
.form-group input[type="text"], .form-group input[type="date"], .form-group input[type="file"] { width: 100%; padding: 10px 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 14px; }
.form-group input[type="file"] { padding: 8px; }

.upload-page { max-width: 600px; margin: 40px auto; }
.upload-page h1 { margin-bottom: 24px; font-size: 24px; }

.empty-state { text-align: center; padding: 60px 20px; color: #888; }
.empty-state h2 { font-size: 20px; margin-bottom: 8px; color: #555; }

.alert { padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; font-size: 14px; display: none; }
.alert-error { background: #fee2e2; color: #991b1b; display: block; }
.alert-success { background: #d1fae5; color: #065f46; display: block; }

.detail-header { margin-bottom: 24px; }
.detail-header h1 { font-size: 22px; margin-bottom: 4px; }
.detail-header .meta { color: #888; font-size: 14px; }

.indicator-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin-bottom: 24px; }
.indicator-card { background: #fff; border-radius: 10px; padding: 16px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
.indicator-label { font-size: 11px; color: #888; margin-bottom: 4px; }
.indicator-value { font-size: 20px; font-weight: 700; color: #1a1a2e; }

.delete-area { margin-top: 32px; padding-top: 24px; border-top: 1px solid #eee; text-align: center; }
```

---

### Task 5: 创建上传页面模板

**Files:**
- Create: `d:\VibeCodingPJ\data_trend\templates\upload.html`

- [ ] **Step 1: 创建 `upload.html`**

```html
{% extends "base.html" %}
{% block title %}上传数据 - 视频趋势监测{% endblock %}
{% block content %}
<div class="upload-page">
    <h1>📤 上传新视频数据</h1>
    <div id="alert" class="alert"></div>
    <form id="uploadForm">
        <div class="form-group">
            <label for="publish_date">发布日期</label>
            <input type="date" id="publish_date" name="publish_date" required>
        </div>
        <div class="form-group">
            <label for="topic">视频主题</label>
            <input type="text" id="topic" name="topic" placeholder="如：Vlog#1 露营记" required>
        </div>
        <div class="form-group">
            <label for="traffic_file">流量数据 .xlsx</label>
            <input type="file" id="traffic_file" name="traffic_file" accept=".xlsx" required>
        </div>
        <div class="form-group">
            <label for="fan_file">粉丝数据 .xlsx</label>
            <input type="file" id="fan_file" name="fan_file" accept=".xlsx" required>
        </div>
        <button type="submit" class="btn btn-primary" style="width:100%;padding:12px;font-size:16px;">提交并分析</button>
    </form>
</div>

<script>
document.getElementById('uploadForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    const alert = document.getElementById('alert');
    alert.className = 'alert';
    alert.style.display = 'none';

    const formData = new FormData(this);

    try {
        const res = await fetch('/api/upload', { method: 'POST', body: formData });
        const data = await res.json();
        if (data.success) {
            alert.className = 'alert alert-success';
            alert.textContent = `✅ "${data.topic}" 上传成功！正在跳转...`;
            alert.style.display = 'block';
            setTimeout(() => { window.location.href = `/video/${data.video_id}`; }, 1500);
        } else {
            alert.className = 'alert alert-error';
            alert.textContent = '❌ ' + (data.error || '上传失败');
            alert.style.display = 'block';
        }
    } catch (err) {
        alert.className = 'alert alert-error';
        alert.textContent = '❌ 网络错误，请重试';
        alert.style.display = 'block';
    }
});
</script>
{% endblock %}
```

---

### Task 6: 创建概览看板模板

**Files:**
- Create: `d:\VibeCodingPJ\data_trend\templates\dashboard.html`

- [ ] **Step 1: 创建 `dashboard.html`**

```html
{% extends "base.html" %}
{% block title %}概览看板 - 视频趋势监测{% endblock %}
{% block content %}
<div id="emptyState" class="empty-state" style="display:none;">
    <h2>📊 还没有数据</h2>
    <p style="color:#888;margin-bottom:16px;">上传你的第一期视频数据，开始监测趋势吧！</p>
    <a href="/upload" class="btn btn-primary">上传第一期数据</a>
</div>

<div id="dashboardContent" style="display:none;">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:24px;">
        <h1 style="font-size:22px;">📊 数据趋势总览</h1>
        <a href="/upload" class="btn btn-primary">+ 上传新数据</a>
    </div>

    <div class="cards" id="summaryCards"></div>

    <div class="chart-container">
        <h2>📊 各期视频播放量对比</h2>
        <canvas id="compareChart"></canvas>
    </div>

    <div class="chart-container">
        <h2>📈 累计趋势（播放量 / 涨粉量）</h2>
        <canvas id="trendChart"></canvas>
    </div>

    <div class="table-wrap">
        <h2 style="font-size:16px;margin-bottom:16px;color:#444;">📋 视频列表</h2>
        <table>
            <thead>
                <tr>
                    <th>视频主题</th>
                    <th>发布日期</th>
                    <th>播放量</th>
                    <th>点赞</th>
                    <th>涨粉</th>
                    <th>完播率</th>
                    <th>操作</th>
                </tr>
            </thead>
            <tbody id="videoTableBody"></tbody>
        </table>
    </div>
</div>

<script>
async function loadDashboard() {
    try {
        const res = await fetch('/api/videos');
        const videos = await res.json();
        if (!videos || videos.length === 0) {
            document.getElementById('emptyState').style.display = 'block';
            return;
        }
        document.getElementById('dashboardContent').style.display = 'block';

        const total = videos.length;
        const avgPlay = Math.round(videos.reduce((s, v) => s + (v['播放量'] || 0), 0) / total);
        const totalFans = videos.reduce((s, v) => s + (v['涨粉量'] || 0), 0);
        const rates = videos.map(v => parseFloat(v['完播率']) || 0);
        const avgRate = rates.length ? (rates.reduce((a, b) => a + b, 0) / rates.length).toFixed(2) + '%' : '0%';

        document.getElementById('summaryCards').innerHTML = `
            <div class="card"><div class="card-label">总视频数</div><div class="card-value">${total}</div></div>
            <div class="card"><div class="card-label">平均播放量</div><div class="card-value">${avgPlay.toLocaleString()}</div></div>
            <div class="card"><div class="card-label">总涨粉</div><div class="card-value">${totalFans.toLocaleString()}</div></div>
            <div class="card"><div class="card-label">平均完播率</div><div class="card-value">${avgRate}</div></div>
        `;

        const sorted = [...videos].sort((a, b) => new Date(a['发布日期']) - new Date(b['发布日期']));
        const labels = sorted.map(v => v['视频主题'] || `#${v['video_id']}`);
        const plays = sorted.map(v => v['播放量'] || 0);
        const likes = sorted.map(v => v['点赞量'] || 0);
        const favorites = sorted.map(v => v['收藏量'] || 0);

        new Chart(document.getElementById('compareChart'), {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: '播放量', data: plays, backgroundColor: '#2563eb' },
                    { label: '点赞量', data: likes, backgroundColor: '#10b981' },
                    { label: '收藏量', data: favorites, backgroundColor: '#f59e0b' }
                ]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'top' } },
                scales: { y: { beginAtZero: true } }
            }
        });

        new Chart(document.getElementById('trendChart'), {
            type: 'line',
            data: {
                labels: sorted.map(v => v['发布日期']),
                datasets: [
                    { label: '播放量', data: plays, borderColor: '#2563eb', fill: false, tension: 0.3 },
                    { label: '涨粉量', data: sorted.map(v => v['涨粉量'] || 0), borderColor: '#ef4444', fill: false, tension: 0.3 }
                ]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'top' } },
                scales: { y: { beginAtZero: true } }
            }
        });

        const tbody = document.getElementById('videoTableBody');
        tbody.innerHTML = videos.map(v => `
            <tr>
                <td><a href="/video/${v['video_id']}">${v['视频主题'] || `视频 #${v['video_id']}`}</a></td>
                <td>${v['发布日期']}</td>
                <td>${(v['播放量'] || 0).toLocaleString()}</td>
                <td>${v['点赞量'] || 0}</td>
                <td>${v['涨粉量'] || 0}</td>
                <td>${v['完播率'] || '-'}</td>
                <td>
                    <a href="/video/${v['video_id']}" class="btn btn-primary btn-sm">详情</a>
                    <button class="btn btn-danger btn-sm" onclick="deleteVideo(${v['video_id']}, '${v['视频主题']}')">删除</button>
                </td>
            </tr>
        `).join('');
    } catch (err) {
        console.error('Failed to load dashboard:', err);
    }
}

async function deleteVideo(id, topic) {
    if (!confirm(`确定删除「${topic}」的所有数据？此操作不可恢复。`)) return;
    try {
        const res = await fetch(`/api/video/${id}/delete`, { method: 'POST' });
        const data = await res.json();
        if (data.success) {
            location.reload();
        } else {
            alert('删除失败: ' + (data.error || ''));
        }
    } catch (err) {
        alert('网络错误，请重试');
    }
}

loadDashboard();
</script>
{% endblock %}
```

---

### Task 7: 创建视频详情模板

**Files:**
- Create: `d:\VibeCodingPJ\data_trend\templates\detail.html`

- [ ] **Step 1: 创建 `detail.html`**

```html
{% extends "base.html" %}
{% block title %}视频详情 - 视频趋势监测{% endblock %}
{% block content %}
<div id="loading" style="text-align:center;padding:60px;color:#888;">加载中...</div>
<div id="content" style="display:none;">
    <a href="/" class="btn" style="margin-bottom:16px;color:#2563eb;">← 返回看板</a>

    <div class="detail-header">
        <h1 id="videoTitle"></h1>
        <div class="meta" id="videoMeta"></div>
    </div>

    <div class="indicator-grid" id="indicatorGrid"></div>

    <div class="chart-container">
        <h2>📈 播放量趋势（小时级）</h2>
        <canvas id="playTrendChart"></canvas>
    </div>

    <div class="chart-container">
        <h2>📈 涨粉量趋势（小时级）</h2>
        <canvas id="fanTrendChart"></canvas>
    </div>

    <div class="delete-area">
        <button class="btn btn-danger" onclick="deleteCurrentVideo()">🗑 删除此视频</button>
    </div>
</div>

<script>
const VIDEO_ID = {{ video_id }};

async function loadDetail() {
    try {
        const res = await fetch(`/api/video/${VIDEO_ID}`);
        const data = await res.json();
        if (data.error) {
            document.getElementById('loading').textContent = '❌ 视频不存在';
            return;
        }

        document.getElementById('loading').style.display = 'none';
        document.getElementById('content').style.display = 'block';

        const s = data.summary;
        document.getElementById('videoTitle').textContent = s['视频主题'] || `视频 #${s['video_id']}`;
        document.getElementById('videoMeta').textContent =
            `发布日期: ${s['发布日期']} | OGC_FID: ${s['OGC_FID']} | 上传时间: ${s['上传时间'] || '-'}`;

        const indicators = [
            ['播放量', s['播放量']], ['点赞量', s['点赞量']], ['评论量', s['评论量']], ['分享量', s['分享量']],
            ['收藏量', s['收藏量']], ['弹幕量', s['弹幕量']], ['完播率', s['完播率'] || '-'], ['2s跳出率', s['2s跳出率'] || '-'],
            ['涨粉量', s['涨粉量']], ['脱粉量', s['脱粉量']], ['粉丝播放占比', s['粉丝播放占比'] || '-']
        ];
        document.getElementById('indicatorGrid').innerHTML = indicators.map(([label, val]) => `
            <div class="indicator-card">
                <div class="indicator-label">${label}</div>
                <div class="indicator-value">${typeof val === 'number' ? val.toLocaleString() : val}</div>
            </div>
        `).join('');

        if (data.play_trend && data.play_trend.length) {
            const playLabels = data.play_trend.map(r => r['日期']);
            new Chart(document.getElementById('playTrendChart'), {
                type: 'line',
                data: {
                    labels: playLabels,
                    datasets: [
                        { label: '抖音', data: data.play_trend.map(r => r['抖音'] || 0), borderColor: '#2563eb', fill: false, tension: 0.3 },
                        { label: '抖音精选', data: data.play_trend.map(r => r['抖音精选'] || 0), borderColor: '#10b981', fill: false, tension: 0.3 }
                    ]
                },
                options: {
                    responsive: true,
                    plugins: { legend: { position: 'top' } },
                    scales: { y: { beginAtZero: true } }
                }
            });
        } else {
            document.querySelector('.chart-container:first-of-type').innerHTML = '<p style="color:#888;padding:20px;text-align:center;">暂无播放量趋势数据</p>';
        }

        if (data.fan_trend && data.fan_trend.length) {
            const fanLabels = data.fan_trend.map(r => r['日期']);
            new Chart(document.getElementById('fanTrendChart'), {
                type: 'line',
                data: {
                    labels: fanLabels,
                    datasets: [
                        { label: '抖音', data: data.fan_trend.map(r => r['抖音'] || 0), borderColor: '#ef4444', fill: false, tension: 0.3 },
                        { label: '抖音精选', data: data.fan_trend.map(r => r['抖音精选'] || 0), borderColor: '#f59e0b', fill: false, tension: 0.3 }
                    ]
                },
                options: {
                    responsive: true,
                    plugins: { legend: { position: 'top' } },
                    scales: { y: { beginAtZero: true } }
                }
            });
        } else {
            const charts = document.querySelectorAll('.chart-container');
            if (charts.length > 1) {
                charts[1].innerHTML = '<p style="color:#888;padding:20px;text-align:center;">暂无涨粉量趋势数据</p>';
            }
        }
    } catch (err) {
        document.getElementById('loading').textContent = '❌ 加载失败，请重试';
    }
}

async function deleteCurrentVideo() {
    if (!confirm('确定删除此视频的所有数据？此操作不可恢复。')) return;
    try {
        const res = await fetch(`/api/video/${VIDEO_ID}/delete`, { method: 'POST' });
        const data = await res.json();
        if (data.success) {
            window.location.href = '/';
        } else {
            alert('删除失败: ' + (data.error || ''));
        }
    } catch (err) {
        alert('网络错误，请重试');
    }
}

loadDetail();
</script>
{% endblock %}
```

---

### Task 8: 端到端验证

- [ ] **Step 1: 启动 Flask 应用**

```bash
cd D:\VibeCodingPJ\data_trend
.venv\Scripts\python.exe app.py
```

预期: `Running on http://127.0.0.1:5000`

- [ ] **Step 2: 打开浏览器访问 `http://127.0.0.1:5000`**

预期: 看到空数据引导页，显示"还没有数据"提示和"上传第一期数据"按钮

- [ ] **Step 3: 点击"上传第一期数据"，进入上传页面**

预期: 看到发布日期选择器、视频主题输入框、两个文件上传控件

- [ ] **Step 4: 选择日期、输入主题、上传两个 .xlsx 文件，点击提交**

预期: 显示上传成功提示，随后跳转到视频详情页

- [ ] **Step 5: 在详情页验证所有指标卡片和趋势图正确渲染**

预期: 11 个指标卡片数值正确，播放量趋势图显示抖音/抖音精选两条线

- [ ] **Step 6: 返回看板，验证概览数据**

预期: 统计卡片数值正确，对比柱状图和累计趋势图渲染正常，视频列表有数据

- [ ] **Step 7: 测试删除功能**

预期: 点击删除后确认弹窗，确认后回到看板，该视频数据消失

- [ ] **Step 8: 再次上传第二期数据（如有），验证多期数据的趋势累加**

预期: 看板显示多期对比，柱状图出现多个柱子

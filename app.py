import os
import sys
import webbrowser
import threading
import time
import pandas as pd
from flask import Flask, request, jsonify, render_template, redirect, url_for

from db import (
    init_db, add_video, get_all_videos, get_video,
    delete_video, get_trends_over_time
)


def _resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)


def _data_path(relative_path=''):
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(__file__)
    if relative_path:
        return os.path.join(base, relative_path)
    return base


app = Flask(__name__,
    template_folder=_resource_path('templates'),
    static_folder=_resource_path('static'))
app.config['UPLOAD_FOLDER'] = _data_path('uploads')
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


def _detect_file_type(file_storage):
    sheet_names = pd.ExcelFile(file_storage).sheet_names
    file_storage.seek(0)
    if '播放量-新增-每小时趋势数据' in sheet_names:
        return 'traffic'
    if '涨粉量-新增-每小时趋势数据' in sheet_names:
        return 'fan'
    return None


@app.route('/api/upload', methods=['POST'])
def api_upload():
    files = request.files.getlist('files')
    if len(files) != 2:
        return jsonify({'error': '请恰好上传两个 .xlsx 文件（流量数据 + 粉丝数据）'}), 400

    publish_date = request.form.get('publish_date', '')
    topic = request.form.get('topic', '')

    if not publish_date or not topic:
        return jsonify({'error': '请填写发布日期和视频主题'}), 400

    traffic_file = fan_file = None
    for f in files:
        if not f.filename.endswith('.xlsx'):
            return jsonify({'error': f'文件 "{f.filename}" 不是 .xlsx 格式'}), 400
        f.stream.seek(0)
        ftype = _detect_file_type(f.stream)
        if ftype == 'traffic':
            traffic_file = f
        elif ftype == 'fan':
            fan_file = f

    if traffic_file is None or fan_file is None:
        return jsonify({'error': '无法自动识别文件类型，请确认上传了流量数据和粉丝数据各一个'}), 400

    backup_dir = app.config['UPLOAD_FOLDER']
    traffic_path = os.path.join(backup_dir, f"{publish_date}_流量数据.xlsx")
    fan_path = os.path.join(backup_dir, f"{publish_date}_粉丝数据.xlsx")
    traffic_file.stream.seek(0)
    fan_file.stream.seek(0)
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
    threading.Thread(target=lambda: (time.sleep(1.5), webbrowser.open('http://127.0.0.1:5000')), daemon=True).start()
    app.run(debug=False, port=5000)

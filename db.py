import pandas as pd
import os
import sys
from datetime import datetime

def _get_db_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(__file__)

DB_PATH = os.path.join(_get_db_dir(), 'data_trend_db.xlsx')

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

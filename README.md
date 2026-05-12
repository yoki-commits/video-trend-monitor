# 视频趋势监测

一个用于长期监测视频账号数据变化趋势的 Web 应用。支持批量导入视频流量和粉丝数据，自动生成趋势图表，帮助分析视频内容的表现。

## 功能

- **数据导入**：支持同时上传流量数据（播放量/完播率/互动数据）和粉丝数据（涨粉量）两个 Excel 文件，系统自动识别文件类型
- **视频管理**：为每期视频记录发布时间、主题，支持查看详情和删除
- **趋势图表**：使用 Chart.js 展示视频数据的 7 天生命周期趋势
- **跨视频对比**：在首页对比所有视频的累计播放量和涨粉量
- **独立可执行**：支持打包为 .exe 文件，无需安装 Python 即可运行

## 技术栈

| 层级 | 技术 |
|------|------|
| 后端 | Python 3.13 + Flask |
| 数据处理 | pandas + openpyxl |
| 前端 | HTML5 + Chart.js (CDN) |
| 数据库 | Excel (.xlsx) |
| 打包 | PyInstaller |

## 项目结构

```
.
├── app.py              # Flask 应用主程序
├── db.py               # 数据库操作（读写 Excel）
├── run.py              # PyInstaller 打包入口
├── requirements.txt    # Python 依赖
├── static/
│   └── style.css       # 前端样式
├── templates/
│   ├── base.html       # 基础布局
│   ├── dashboard.html  # 首页/数据总览
│   ├── detail.html     # 单视频详情页
│   └── upload.html     # 数据上传页
├── 流量数据.xlsx       # 示例：流量数据模板
├── 粉丝数据.xlsx       # 示例：粉丝数据模板
└── data_trend_db.xlsx  # 运行时自动生成的数据库
```

## 安装与运行

### 方式一：源码运行

```bash
# 1. 克隆仓库
git clone https://github.com/yoki-commits/video-trend-monitor.git
cd video-trend-monitor

# 2. 创建虚拟环境（推荐）
python -m venv .venv
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # macOS/Linux

# 3. 安装依赖
pip install -r requirements.txt

# 4. 运行
python app.py
```

服务启动后自动打开浏览器，访问 http://127.0.0.1:5000

### 方式二：可执行文件运行

下载 `视频趋势监测.exe`，双击运行即可，无需安装任何环境。

> 数据库文件 `data_trend_db.xlsx` 和上传目录 `uploads/` 会自动创建在 .exe 同级目录。

## 数据上传说明

上传时需要准备两个 Excel 文件：

| 文件 | 说明 | 必需字段 |
|------|------|---------|
| 流量数据 | 播放量、完播率、互动数据 | 发布时间、播放量、完播率、点赞量、评论量、分享量、收藏量 |
| 粉丝数据 | 涨粉量趋势 | 发布时间、涨粉量 |

系统通过读取 Excel 的 Sheet 名称自动识别文件类型，无需手动标注。

## 打包为 .exe

```bash
pip install pyinstaller
pyinstaller --onefile --add-data "templates;templates" --add-data "static;static" --hidden-import openpyxl --name "视频趋势监测" run.py
```

打包后的可执行文件位于 `dist/视频趋势监测.exe`。

## 许可证

MIT

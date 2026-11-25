# Commercial Film Review Skill

> 面向 **AI Coding / Multimodal Agent**（Antigravity / Claude Code / Codex / OpenCode 等）的**品牌形象片、TVC 与产品发布片专业审片技能**，内置**量化音视频分析脚本**与**一键商业排版 Word（`.docx`）审片报告生成器**。

**English Summary**: *An evidence-based Agent Skill for reviewing commercial brand films, TVCs, and industrial/tech product launch videos (with dedicated rubrics for Solar/New-Energy and Elevator/Vertical Mobility). Runs local FFmpeg/Python analysis to extract scene cuts, ASL pacing, EBU R128 loudness, color jumps, color barcodes, and contact sheets — then outputs a timestamped P0/P1/P2 DaVinci/PR/AE fix list and exports a publication-ready `.docx` report.*

---

## 为什么做这个 Skill？

市面上绝大多数 AI 视频分析 Prompt，默认套用的是信息流短视频的话术——动不动就讲“前 2 秒黄金钩子、5 秒留存、完播率、痛点反转、爆款文案”。

但如果你审的是**企业品牌形象片、TVC 广告、展会大屏片或工业/科技制造类产品发布片**，那套逻辑不仅毫无帮助，还会把片子改得廉价、浮躁。

商业宣传片真正看重的是：
- **叙事与品牌情绪弧线**：是高级的品牌叙事，还是线性平铺的“政企汇报片 / PPT 念稿感”？
- **摄影构图与工业美学**：景别呼吸感、建筑/阵列几何秩序、材质高光控制、边缘穿帮检查。
- **剪辑点与动势节奏**：切点是否踩在音乐重拍/动作发力点、高潮段有没有单镜硬撑泄气、分段 ASL 是否张弛有度。
- **配乐与声音设计**：综合响度（LUFS）、真峰值（True Peak）、响度动态范围（LRA）、结尾有没有突然静音断崖、环境音与质感 SFX 是否空缺。
- **调色与影调连贯性**：相邻镜头白平衡/反差/黑位有没有跳色断层、航拍与地面/CG 之间是否融合、高光有没有死白裁切。
- **文字包装与品牌落版（Packshot）**：专业术语与口径准确性、角标在亮背景/高频纹理下是否隐身、字幕锚点是否统一、落版可读性。

**Commercial Film Review Skill** 把这套专业后期导演/调色师/声音指导的审片框架工程化，强制 Agent **先跑脚本拿客观数据、再完整视听核查、最后输出精确到时间码的改片清单与可交付 Word 文档**。

---

## 核心特性

### 1. 零门槛本地量化体检（`scripts/commercial_metrics.py`）
只需 `ffmpeg` / `ffprobe` + Python，自动从本地视频提取客观证据，绝不靠模型“猜时间码”：
- **剪辑与节奏统计**：自动检测镜头硬切点、总镜数、全片 **ASL（平均镜头时长）**，以及开篇 25% / 中段 50% / 尾声 25% 的分段节奏曲线。
- **广播级响度与包络测量**：基于 `loudnorm` 实测 **Integrated LUFS / True Peak (dBTP) / LRA (LU)**，并输出逐秒 RMS 响度包络（一眼抓出结尾静音断崖或全片动态过平）。
- **色彩连贯性与曝光体检**：逐镜统计均值 RGB/HEX、亮度、反差（`luma_std`）、饱和度、高光死白裁切占比（`> 96%`）与暗部死黑占比（`< 4%`），并自动标出**相邻镜头强跳色预警**。
- **可视化审片资产**：一键输出全片**色彩影调色带图 (`color_barcode.png`)** 与**带镜头号/时间码/时长/主色值的故事板拼图 (`contact_sheet.jpg`)**。

### 2. 工业与科技制造垂直行业审片细则（`references/`）
按需加载，不浪费上下文 Token：
- [`references/commercial-rubric.md`](references/commercial-rubric.md)：品牌形象片与产品发布片六维通用工业标准
- [`references/solar-new-energy.md`](references/solar-new-energy.md)：**光伏 / 储能 / 新能源专项标准**（航拍与地面色温断层、电池片/组件高光控制、工程验收片避坑、构网/参数视觉化）
- [`references/elevator-mobility.md`](references/elevator-mobility.md)：**电梯与垂直交通专项标准**（垂直几何畸变、不锈钢/镜面穿帮反光、平稳零顿挫运动语言、静谧阻尼声景）

> 💡 你可以仿照 `references/` 下的 Markdown 模板，轻松扩展自己的垂直行业细则（例如汽车、消费电子、医疗器械、建筑地产等）。

### 3. 一键导出商业排版 Word 审片报告（`scripts/export_docx_report.py`）
内置通用 JSON → `.docx` 报告生成器，直接产出可发给导演、剪辑师、调色师、客户方的正式审片文档：
- 标准 A4 页面规范、页眉页脚、藏青商业视觉体系（`微软雅黑` + `Segoe UI`）
- 左边框高亮「一、总体诊断」核心评级与亮点/硬伤卡片
- 斑马纹「二、量化体检表」，**自动内嵌全片色带图（Color Barcode）与故事板（Contact Sheet）**
- 结构化「三、六维深度拆解」
- 红 / 琥珀 / 绿三色徽章「四、按时间码排列的 P0 / P1 / P2 改片清单」（自带跨页重复表头与防行内断页）
- 「五、系列共性提示与局限说明」

---

## 工作流一览

```text
原始视频 (.mp4 / .mov)
   │
   ├─► 1. 运行 scripts/commercial_metrics.py
   │       └─► 输出 metrics.json + color_barcode.png + contact_sheet.jpg + 抽帧
   │
   ├─► 2. Agent 读图 + 完整视听审片 + 按需加载行业细则 (references/*.md)
   │       └─► （可选联动 video-close-reading 获取毫秒级音频重拍帧差）
   │
   ├─► 3. 输出结构化审片结论（P0 必改 / P1 优化 / P2 亮点保留）
   │
   └─► 4. 运行 scripts/export_docx_report.py
           └─► 一键生成带色带与故事板的专业 Word 审片报告 (.docx)
```

---

## 目录结构

```text
commercial-film-review/
├── SKILL.md                          # Skill 主指令（Agent 自动发现与调度入口）
├── README.md                         # 项目介绍与使用说明
├── LICENSE                           # MIT 开源协议
├── requirements.txt                  # Python 依赖清单
├── references/
│   ├── commercial-rubric.md          # 品牌形象片 / 产品发布片通用六维审片细则
│   ├── solar-new-energy.md           # 光伏新能源行业专项审片标准
│   └── elevator-mobility.md          # 电梯与垂直交通行业专项审片标准
├── scripts/
│   ├── commercial_metrics.py         # 音视频量化提取（切点/ASL/LUFS/跳色/色带/故事板）
│   └── export_docx_report.py         # 通用商业排版 Word (.docx) 审片报告生成器
└── examples/
    └── sample_report.json            # Word 报告导出器的标准 JSON 结构示例
```

---

## 快速安装与使用

### 1. 安装环境依赖

确保系统已安装 Python 3.10+ 并将 `ffmpeg`、`ffprobe` 加入环境变量 `PATH`，然后安装 Python 依赖：

```bash
pip install -r requirements.txt
```

### 2. 安装到你的 AI Agent

将整个文件夹放入对应 Agent 的 Skills 目录即可自动识别：

- **Antigravity**：`~/.gemini/config/skills/commercial-film-review/`
- **Claude Code**：`~/.claude/skills/commercial-film-review/`
- **Codex**：`~/.codex/skills/commercial-film-review/`

### 3. 直接在对话中触发

```text
用 /commercial-film-review 帮我审一下这条产品发布片：/path/to/film.mp4，审完直接导出 Word 审片报告。
```

### 4. 单独测试脚本（命令行）

```bash
# 1) 提取量化指标、色带与故事板
python scripts/commercial_metrics.py "/path/to/video.mp4" --out "./review_out" --scene-threshold 0.28

# 2) 用示例 JSON 测试一键生成排版 Word 报告
python scripts/export_docx_report.py examples/sample_report.json --out "./sample_review.docx"
```

---

## 致谢与可选联动（Acknowledgments）

- 本技能的所有代码、行业审片标准与 Word 报告导出器均独立编写，可**完全独立运行**。
- 当需要对高密度节奏片做毫秒级音频起音/重拍帧差（Onset/Beat Alignment）或原分辨率连续帧逐像素测色时，推荐搭配开源技能 [xiaolu-ai26/video-close-reading](https://github.com/xiaolu-ai26/video-close-reading)（MIT License）联合使用；本技能“先证据后判断、诚实标注未验证边界”的审片纪律亦受其启发。

---

## License

[MIT License](LICENSE)

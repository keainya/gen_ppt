---
name: gen-ppt
description: 'Generate PowerPoint (PPTX) presentations from structured Markdown. Use when: user asks to create a PPT, PowerPoint, slide deck, 演示文稿, PPT, or presentation from Markdown/outline/content; generating slides with cover, content pages, image gallery, text+image layout, tables; converting structured documents to .pptx format.'
argument-hint: '[markdown_file] — optional, defaults to content.md'
user-invocable: true
---

# Gen PPT — Markdown to PowerPoint

将结构化 Markdown 编译为 PowerPoint (.pptx) 演示文稿。使用 `python-pptx` 库，支持封面页、内容页、图片展示页、图配文页和结束页共 5 种页面类型，极简样式，侧重内容清晰呈现。

## When to Use

- 用户要求根据 Markdown 内容生成 PPT/演示文稿
- 需要将结构化文档（含标题、列表、表格、图片）转为幻灯片
- 快速制作演示文稿，无需复杂排版和动画

## Page Types & Markdown Syntax

### 封面页 (Cover)
- `# 标题` → PPT 标题（宋体 66pt）
- `## 演讲人` → 演讲人姓名（宋体 44pt）
- 背景图使用 `cover.png`

### 内容页 (Content)
- `## 标题` → 二级标题（宋体 44pt）
- `1. 项目` → 有序列表
- `- 项目` / `* 项目` → 无序列表
- Markdown 表格 → 三段式表格（无底色、无竖线，仅顶/中/底三条横线）
- 若页面**仅**包含一个 `## 标题`（无其他文字或表格），标题在幻灯片中水平和垂直居中
- 背景图使用 `background.png`

### 图片展示页 (Image)
- 每行一个 `![说明文字](图片路径)`，最多 6 张
- 自动布局：1 张居中、2 张左右、3 张上 2 下 1、4 张 2×2、5~6 张 3 列网格
- 超过 6 张仅取前 6 张并输出警告
- 背景图使用 `background.png`

### 图配文页 (Text + Image)（尽量少用！）
- 左侧 2/3 为文字区域（支持二级标题、列表、表格）
- 右侧 1/3 为单张图片
- 页面同时包含结构化内容和图片引用时自动识别
- 背景图使用 `background.png`

### 结束页 (End)
- 自动生成，无需在 Markdown 中编写
- 内容："谢谢"（宋体 66pt，居中）
- 背景图使用 `cover.png`

## Page Separation

使用 `---`（三个减号）作为页面分隔符。第一个 `---` 之前可放置元数据。

```markdown
<!-- cover: slides/cover.png -->
<!-- background: slides/bg.png -->

# 演示文稿标题
## 演讲人姓名

---

## 第一个内容页标题
- 要点一
- 要点二

---

![图1](images/chart1.png)
![图2](images/chart2.png)
```

## Metadata (Optional)

在第一个 `---` 之前通过 HTML 注释指定背景图：

- `<!-- cover: path/to/cover.png -->` — 封面及结束页背景图
- `<!-- background: path/to/bg.png -->` — 内容页/图片页/图配文页背景图
- 路径相对于 Markdown 文件所在目录

## Font Specifications

| 文本类型 | 中文字体 | 英文字体 | 字号 |
|----------|----------|----------|------|
| 一级标题 | 宋体 (SimSun) | Times New Roman | 66pt |
| 二级标题 | 宋体 (SimSun) | Times New Roman | 44pt |
| 正文 | 仿宋 (FangSong) | Times New Roman | 32pt |
| 表格表头 | 宋体 (SimSun) | Times New Roman | 32pt |
| 表格内容 | 仿宋 (FangSong) | Times New Roman | 32pt |

正文支持 `**加粗**` 行内语法。

## Procedure

### 1. Understand the User's Request

确定用户想要生成什么内容的 PPT。如果用户提供了具体内容大纲，按 Markdown 语法结构编写。如果用户只描述了主题，先构思内容结构。

### 2. Create the Markdown Content

编写或更新 `content.md`（或用户指定的文件），遵循以下规则：

- 使用 `#` 一级标题定义封面标题
- 使用 `##` 二级标题定义演讲人或内容页标题
- 使用 `---` 分隔不同页面
- 图片页使用 `![说明](路径)` 语法
- 表格使用标准 Markdown 表格语法
- 通过 HTML 注释指定自定义背景图

### 3. Prepare Assets (if needed)

提示用户将背景图片放置在正确位置：
- `cover.png`：与 `build_ppt.py` 同目录，用于封面和结束页
- `background.png`：与 `build_ppt.py` 同目录，用于内容页/图片页/图配文页
- 图片引用路径相对于 Markdown 文件所在目录

### 4. Run the Build Script

```bash
python build_ppt.py [markdown_file]
```

- 若未指定文件，默认读取 `content.md`
- 输出文件与输入 Markdown 同名，扩展名为 `.pptx`

### 5. Verify Output

确认生成的 `.pptx` 文件页码正确、内容完整、图片显示正常。

## Dependencies

项目需要以下 Python 包：
- `python-pptx` — PPT 生成
- `Pillow` (PIL) — 图片尺寸读取
- `lxml` — XML 操作（设置中文字体）

## File Structure

```
gen_ppt/
├── build_ppt.py       # 主构建脚本
├── package.py         # 打包脚本（生成 .exe）
├── content.md         # 默认输入文件
├── cover.png          # 封面背景（可选）
├── background.png     # 内容页背景（可选）
├── dist/              # 打包输出目录
└── README.md          # 项目文档
```

## Tips

- 内容页若只需展示一个大标题（章节页），仅写一个 `## 标题` 即可自动居中
- 图片路径支持相对路径（相对于 Markdown 文件所在目录）
- 表格采用三段式学术风格，适合数据清晰展示
- 打包为独立可执行文件：`python package.py`（Windows 生成 `.exe`，Linux 生成二进制文件）
- 背景图缺失时自动回退为纯白背景，不会报错

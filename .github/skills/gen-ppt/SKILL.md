---
name: gen-ppt
description: '将结构化 Markdown 编译成 PPT。Use when: 生成演示文稿、制作PPT、编译markdown到pptx、写content.md、修改PPT样式或字体、添加页面类型、调整build_ppt.py。'
argument-hint: '[task description for gen_ppt]'
user-invocable: true
---

# gen_ppt — 结构化 Markdown 编译为 PPT

将结构化的 `content.md` 编译为 `output.pptx`，支持 5 种页面类型及自动结束页。

## 项目文件

| 文件 | 作用 |
|------|------|
| `build_ppt.py` | 主构建脚本，读取 content.md 生成 .pptx |
| `content.md` | 结构化 Markdown 输入文件 |
| `cover.png` | 封面及结束页背景图（与 build_ppt.py 同目录） |
| `background.png` | 内容页及图片页背景图（与 build_ppt.py 同目录） |
| `src/` | 长期资源目录，存放复用图片素材 |
| `tmp/` | 临时资源目录，存放一次性使用的图片 |
| `clean.py` | 清理脚本，清除生成文件及 `tmp/` 中临时资源 |
| `package.py` | 打包脚本，生成独立可执行文件（Windows .exe / Linux binary） |

## 资源目录：src vs tmp

| 目录 | 用途 | clean.py 行为 |
|------|------|---------------|
| `src/` | 长期资源（复用的图片素材） | 保留 `.gitkeep`，其余清除 |
| `tmp/` | 临时资源（一次性使用的图片） | 保留 `.gitkeep`，其余全部清除 |

> 背景图 `cover.png` 和 `background.png` 位于项目根目录（与 `build_ppt.py` 同目录）。图片引用路径相对于项目根目录，支持任意路径（如 `src/photo.png`、`tmp/photo.png`、`images/diagram.png` 等），不限制目录。

## 运行方式

```bash
# 编译 content.md（默认）
python build_ppt.py

# 编译指定文件
python build_ppt.py path/to/slides.md
```

- 未指定文件时，默认读取 `content.md`
- 若 `content.md` 也不存在，报错提示缺少文件

## 打包为独立可执行文件

```bash
python package.py              # 为当前平台打包
python package.py --clean      # 先清理再打包
```

- Windows → `dist/build_ppt.exe`
- Linux   → `dist/build_ppt`
- 首次运行自动安装 PyInstaller
- PyInstaller 不支持跨平台编译，需在目标 OS 上运行

## content.md 结构化格式

### 页面分隔

使用三个减号 `---` 作为页面分隔符。两个 `---` 之间的内容属于同一页 PPT（第一页前可省略）。

### 封面页

使用一级标题 `# ` + 二级标题 `## `，自动识别为封面：

```markdown
# 演示文稿标题
## 演讲人姓名
---
```

### 内容页

支持二级标题 `## `、有序列表 `1. `、无序列表 `- `、纯文本、Markdown 表格：

```markdown
## 小节标题
- 要点一
- 要点二
1. 步骤一
2. 步骤二

| 列A | 列B |
|-----|-----|
| 值1 | 值2 |
```

**特殊布局**：若该页仅包含一个二级标题，无其他文字或表格，则该标题在幻灯片中水平和垂直居中显示。

```markdown
## 居中标题
---
```

### 图片展示页

图片说明文字（普通文本） + Markdown 图片引用，支持任意张图片，自动选择布局：

| 张数 | 布局 |
|------|------|
| 1 | 居中大幅展示 |
| 2 | 左右均分 |
| 3 | 上排 2 张 + 下排 1 张居中 |
| 4 | 2×2 田字格 |
| 5～6 | 3 列网格（最多 2 行） |
| ≥7 | 仅取前 6 张并警告 |

每张图可附带 alt 文本（`![说明](path)`），渲染为图片下方的小号说明。

```markdown
架构示意图
![](src/diagram.png)
```

图片也可放在 `tmp/` 目录（临时资源）：

```markdown
临时截图
![](tmp/screenshot.png)
```

多图示例：

```markdown
流程图对比
![方案A](src/flow_a.png)
![方案B](tmp/flow_b.png)
```

### 图配文页（尽可能少用）

同时包含结构化文字内容（二级标题、列表、表格等）和图片引用，自动识别为图配文页。左侧 2/3 渲染文字，右侧 1/3 渲染图片：

```markdown
## 系统架构
- 前端采用 React 框架
- 后端使用 Python FastAPI
- 数据库采用 PostgreSQL

![](src/architecture.png)
```

> 图配文页的图片同样支持 `src/` 或 `tmp/` 路径。

### 结束页

**无需在 Markdown 中编写**，PPT 末尾自动生成「谢谢」页。

## 页面类型速查

| 类型 | 检测规则 | 背景图（可通过元数据覆盖） |
|------|---------|---------------------------|
| 封面 | `# ` 一级标题 | `cover.png` |
| 内容 | 列表、表格、纯文本等 | `background.png` |
| 图片 | `![](...)` 图片引用（无结构化文字） | `background.png` |
| 图配文 | 结构化文字 + 图片引用 | `background.png` |
| 结束 | 自动追加 | `cover.png` |

背景图缺失时回退为纯白背景。

## 元数据（可选）

在第一个 `---` 之前，通过 HTML 注释指定背景图，覆盖默认值：

```markdown
<!-- cover: custom_cover.png -->
<!-- background: custom_bg.png -->

# 标题
## 演讲人
---
```

- `cover`：封面及结束页背景图
- `background`：内容页、图片页、图配文页背景图
- 路径相对于 Markdown 文件所在目录
- 未指定时使用默认值 `cover.png` / `background.png`

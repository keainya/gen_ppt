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
| `src/cover.png` | 封面及结束页背景图 |
| `src/background.png` | 内容页及图片页背景图 |

## 运行方式

```bash
./build_ppt.py
# 或
python build_ppt.py
```

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

### 图片展示页

图片说明文字（普通文本） + Markdown 图片引用：

```markdown
架构示意图
![](src/diagram.png)
```

### 图配文页

同时包含结构化文字内容（二级标题、列表、表格等）和图片引用，自动识别为图配文页。左侧 2/3 渲染文字，右侧 1/3 渲染图片：

```markdown
## 系统架构
- 前端采用 React 框架
- 后端使用 Python FastAPI
- 数据库采用 PostgreSQL

![](src/architecture.png)
```

### 结束页

**无需在 Markdown 中编写**，PPT 末尾自动生成「谢谢」页。

## 页面类型速查

| 类型 | 检测规则 | 背景图 |
|------|---------|--------|
| 封面 | `# ` 一级标题 | `src/cover.png` |
| 内容 | 列表、表格、纯文本等 | `src/background.png` |
| 图片 | `![](...)` 图片引用（无结构化文字） | `src/background.png` |
| 图配文 | 结构化文字 + 图片引用 | `src/background.png` |
| 结束 | 自动追加 | `src/cover.png` |

背景图缺失时回退为纯白背景。

# Markdown 转 Word 使用说明

## 文件结构

解压后建议保持以下结构。普通使用时，只需要关注根目录下的 `convert.bat`：

```text
md2word-win/
  convert.bat
  README.md
  _md2word/
    md2docx.py
    postprocess_docx.py
    diagram-filter.lua
    template.docx
    install-tools.ps1
    tools/
      python/
        python.exe
      pandoc/
        pandoc.exe
      mermaid/
        mmdc.cmd
        node_modules/
          .bin/
            mmdc.cmd
      plantuml/
        plantuml.jar
      node/
        node.exe
```

`_md2word` 是程序内部目录，里面放模板、脚本和可选工具。不要单独移动 `convert.bat` 或 `_md2word`，两者需要保持在同一个 `md2word-win` 目录下。

`tools` 目录中的内容：

- `tools/python/python.exe`：必需，用于运行转换程序；安装脚本会下载便携版 Python。
- `tools/pandoc/pandoc.exe`：必需，用于 Markdown 转 DOCX。
- `tools/mermaid/mmdc.cmd`：可选，用于渲染 Mermaid 图。
- `tools/plantuml/plantuml.jar`：可选，用于渲染 PlantUML 图。
- `tools/node/node.exe`：可选，安装 Mermaid CLI 时如果系统没有 Node.js，会自动下载便携版 Node.js 到这里。

Mermaid 渲染需要浏览器内核。程序会自动优先复用系统已安装的 Edge 或 Chrome；如果没有 Edge/Chrome，才需要让 Mermaid CLI / Puppeteer 下载 Chromium。

Markdown 中的图片建议和 `.md` 文件放在同一项目目录下，转换后图片会嵌入到 Word 文档中。

## 第一次准备

在 PowerShell 中进入内部目录安装 Python 和 Pandoc：

```powershell
cd path\to\md2word-win\_md2word
.\install-tools.ps1
```

如果还需要 PlantUML：

```powershell
.\install-tools.ps1 -WithPlantUml
```

如果还需要 Mermaid：

```powershell
.\install-tools.ps1 -WithMermaid
```

如果 Mermaid 和 PlantUML 都需要：

```powershell
.\install-tools.ps1 -WithDiagrams
```

## 转换命令

最简单用法：

```bat
convert.bat thesis.md
```

指定输出文件：

```bat
convert.bat thesis.md thesis.docx
```

如果要渲染 Mermaid / PlantUML：

```bat
convert.bat thesis.md thesis.docx --diagrams
```

如果希望“有工具就渲染，没有工具就保留为代码块”，使用：

```bat
convert.bat thesis.md thesis.docx --auto-diagrams
```

转换程序会自动执行必要的 DOCX 版式修正，不需要手工运行后处理脚本。

## Markdown 写法约定

一级标题会转换为章标题：

```markdown
# 引言
```

输出为：

```text
第1章 引言
```

二级标题：

```markdown
## 研究背景
```

输出为：

```text
1.1 研究背景
```

摘要建议写在 YAML 元数据中：

```markdown
---
title: "论文题目"
author: "作者"
date: "2026-06-23"
abstract: |
  这里写中文摘要内容。
---
```

程序会自动生成目录，并把摘要、目录和每个一级章节放到新页。

## 图片路径

推荐写法是使用相对于 Markdown 文件所在目录的路径：

```markdown
![系统架构图](images/architecture.png)
![实验结果](figures/result-01.jpg)
```

如果图片和 Markdown 在同一目录：

```markdown
![封面图](cover.png)
```

路径包含空格时，建议使用尖括号：

```markdown
![流程图](<images/process flow.png>)
```

Windows 绝对路径也可以使用，但不推荐，因为换电脑后容易失效：

```markdown
![截图](C:/Users/your-name/Pictures/screenshot.png)
```

网络图片也可以使用，但需要转换时能访问网络：

```markdown
![在线图片](https://example.com/image.png)
```

程序会先尝试把网络图片下载到临时目录，再交给 Pandoc 嵌入 Word；Windows 上如果 Python 下载超时，会自动用系统自带的 `curl.exe` 兜底。生成的 Word 文档中保存的是图片内容，不依赖原始网络链接继续可用。

支持的常见格式包括 `png`、`jpg/jpeg`、`gif`、`svg`。为了 Word 兼容性，论文和报告中优先使用 `png` 或 `jpg`。

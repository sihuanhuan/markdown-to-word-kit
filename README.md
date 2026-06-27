# markdown-to-word-kit

Markdown 转 Word DOCX 工具包，支持自定义模板、目录、图片、图表渲染、列表优化和页眉页脚迁移。

## 仓库内容

这个 Git 仓库只保留必要的脚本、模板和说明：

- `md2word-mac/`：Mac/Linux 轻量版本，使用系统已安装的 `python3` 和 `pandoc`。
- `md2word-win/`：Windows 版本源码和模板，不在 Git 中保存便携工具目录。
- `md2word-win/_md2word/template.docx`：Pandoc reference docx 模板。
- `md2word-win/_md2word/import-header-footer.py`：从公司 Word 文档迁移页眉页脚的脚本。

不放进 Git 的内容：

- `dist/` 发布压缩包
- `md2word-win/_md2word/tools/` Windows 便携工具目录
- 验证输出、临时文件、历史调试文件

## 发布包

发布包放在 GitHub Releases 中，不直接提交到 Git：

- `markdown-to-word-kit-v0.1.0-windows-portable.zip`：Windows 便携包，包含 Python、Pandoc、Node.js、Mermaid CLI 和已下载资源，解压后可直接运行 `convert.bat`。
- `markdown-to-word-kit-v0.1.0-windows-minimal.zip`：Windows 最小包，不包含 Python、Pandoc、Node.js 或 Mermaid CLI；适合从源码安装工具，或在已有工具环境中使用。
- `markdown-to-word-kit-v0.1.0-macos-standard.zip`：macOS 标准包，包含模板和脚本，不内置 Pandoc/Node，使用本机已安装工具。
- `markdown-to-word-kit-v0.1.0-macos-minimal.zip`：macOS 最小包，内容与标准包一致但命名明确，适合和 Windows 最小包一起分发。
- `SHA256SUMS.txt`：发布包校验值。

本地打包产物默认位于：

```text
dist/
```

## Mac 使用

```bash
cd md2word-mac
./convert.sh input.md output.docx
```

## Windows 使用

如果希望解压即用，推荐从 Release 下载 `markdown-to-word-kit-v0.1.0-windows-portable.zip`，解压后运行：

```bat
convert.bat input.md output.docx
```

如果希望包体积小，下载 `markdown-to-word-kit-v0.1.0-windows-minimal.zip`，解压后先安装工具：

```powershell
cd md2word-win\_md2word
.\install-tools.ps1 -WithMermaid
```

如果要渲染 Mermaid / PlantUML：

```bat
convert.bat input.md output.docx --diagrams
```

## 页眉页脚迁移

拿到公司 Word 模板后，可以只迁移页眉页脚，不破坏当前 Markdown 样式：

```bash
python3 md2word-mac/_md2word/import-header-footer.py company-template.docx md2word-mac/_md2word/template.docx -o template.company.docx
```

Windows 便携包中也包含同一个脚本：

```bat
cd md2word-win\_md2word
tools\python\python.exe import-header-footer.py company-template.docx template.docx -o template.company.docx
```

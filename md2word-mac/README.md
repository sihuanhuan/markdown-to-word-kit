# Markdown 转 Word Mac 使用说明

## 依赖

Mac 包不内置 Pandoc、Node 或 Mermaid CLI，默认使用系统中已经安装的工具：

- 必需：`python3`
- 必需：`pandoc`
- 可选：`mmdc`，用于渲染 Mermaid
- 可选：`plantuml`，用于渲染 PlantUML

检查命令：

```bash
python3 --version
pandoc --version
```

## 转换

```bash
cd path/to/md2word-mac
chmod +x convert.sh
./convert.sh thesis.md
```

指定输出文件：

```bash
./convert.sh thesis.md thesis.docx
```

如果本机安装了 Mermaid / PlantUML 工具，可以启用图表渲染：

```bash
./convert.sh thesis.md thesis.docx --diagrams
```

如果希望有工具就渲染、没有工具就跳过：

```bash
./convert.sh thesis.md thesis.docx --auto-diagrams
```

## 替换页眉页脚

拿到公司 Word 模板后，可以只迁移页眉页脚，不破坏当前 Markdown 样式：

```bash
python3 _md2word/import-header-footer.py company-template.docx _md2word/template.docx -o _md2word/template.company.docx
cp _md2word/template.company.docx _md2word/template.docx
```

建议先检查 `template.company.docx`，确认页眉页脚无误后再覆盖正式模板。

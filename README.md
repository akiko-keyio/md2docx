# md2docx

将 Markdown 转换为 Word `.docx`。轻量纯 Python 实现，零外部依赖。支持 Word 原生公式、图表嵌入、交叉引用。默认输出 SCI 论文提交格式，也支持自定义样式。

## How to Use

> "把 manuscript.md 转成 Word 文档"

> "把这篇论文导出为 .docx"

## Note

- 需要本机安装 Microsoft Office（公式转换依赖其样式表）。
- 默认样式为学术期刊投稿格式，可通过 `-p` 参数切换其他样式配置。

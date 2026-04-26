# HTML to DOCX Layout Sync

将 HTML 内容写入 Word（DOCX），并尽可能保留版式信息，包含：

- 页面大小（`@page size`）
- 页边距（`@page margin`）
- 字体、字号、行距
- 段前段后、首行缩进、对齐
- 标题 `h1~h6`、段落 `p/div/li`
- 行内 `strong/b`（加粗）、`em/i`（斜体）、`u`（下划线）

## 依赖

```bash
pip install python-docx beautifulsoup4
```

## 用法

```bash
python html_to_docx.py \
  --html /root/output_synced.html \
  --output /root/output_synced.docx
```

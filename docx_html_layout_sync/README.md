# DOCX 到 HTML 版式同步

这个工具会读取 `docx` 的版式信息，并把这些信息转换成 CSS 注入到 HTML 中，使 HTML 页面在尺寸、页边距、字号、行距等方面尽量与 `docx` 保持一致。

## 覆盖内容
- 页面宽高（首节）
- 页边距（上/右/下/左）
- 正文字体名、字号、行距、段前段后、默认粗细
- 标题 `h1 ~ h6`（映射到 `Heading 1 ~ Heading 6` 样式）
- `strong/b` 加粗规则

## 命令行用法

```bash
python tools/docx_html_layout_sync/sync_docx_layout_to_html.py \
  --docx template.docx \
  --html input.html \
  --output output_synced.html \
  --profile-output docx_layout_profile.json
```

python /root/docx_html_layout_sync/sync_docx_layout_to_html.py \
  --docx "/root/大模型需求描述说明_V4_04022026.docx" \
  --html "/root/排版原文测试.html \
  --output output_synced.html \
  --profile-output docx_layout_profile.json



参数说明：
- `--docx`: 版式来源 DOCX
- `--html`: 待排版 HTML
- `--output`: 输出 HTML
- `--profile-output`: 可选，导出提取到的样式 JSON

## Python 调用

```python
from tools.docx_html_layout_sync import sync_docx_layout_to_html

result = sync_docx_layout_to_html(
    docx_path="template.docx",
    html_path="input.html",
    output_html_path="output_synced.html",
    profile_output_path="docx_layout_profile.json",
)
print(result)
```


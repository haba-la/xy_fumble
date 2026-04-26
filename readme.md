# 文档排版与一致性检查工作台

一个面向 DOCX 文档处理的端到端小工具集，提供 Web 界面一键完成以下流程：

1. DOCX 转临时 HTML
2. 读取模板 DOCX 版式并同步到 HTML
3. 将同步后的 HTML 回写为 DOCX
4. 对同步后的 HTML 做文本一致性检查（LLM）

适合需要“保持模板格式 + 保留文本内容 + 批量快速处理”的场景。

## 功能概览

- 双文件上传：待修改文档 + 模板文档（均为 `.docx`）
- 自动产出：
	- 排版后 HTML（可预览）
	- 排版后 DOCX（可下载）
	- 版式画像 JSON（可下载）
	- 一致性检查 JSON（任务目录内）
- 门户支持任务隔离：每次处理生成独立 `job_id` 目录

## 处理流程

```text
source.docx + template.docx
				|
				|--(1) docx_to_html_temp/docx_to_html.py
				|       source.docx -> source_temp.html
				|
				|--(2) docx_html_layout_sync/layout_sync.py
				|       template.docx + source_temp.html -> synced_layout.html + layout_profile.json
				|
				|--(3) html_to_docx_sync/html_to_docx.py
				|       synced_layout.html -> synced_layout.docx
				|
				|--(4) html_consistency_checker/check_html_consistency.py
								synced_layout.html -> consistency_result.json
```

## 目录结构

```text
all_check/
├─ web_portal.py                              # Web 入口
├─ requirement.txt                            # Python 依赖
├─ web_portal/
│  ├─ templates/layout_portal.html            # 前端页面
│  └─ static/layout_portal.css                # 页面样式
├─ docx_to_html_temp/
│  └─ docx_to_html.py                         # DOCX -> HTML
├─ docx_html_layout_sync/
│  └─ layout_sync.py                          # 模板版式 -> HTML CSS
├─ html_to_docx_sync/
│  └─ html_to_docx.py                         # HTML -> DOCX
├─ html_consistency_checker/
│  ├─ check_html_consistency.py               # 一致性检查
│  └─ prompt_consistency.txt                  # 检查提示词
├─ rule_main_demo/
│  └─ demo.py                                 # LLM 调用封装（需提供 call_llm）
└─ runtime/portal_jobs/jobs/                  # 运行时任务产物
```

## 环境要求

- Python `3.12.x`（项目当前记录为 `3.12.3`）

## 安装依赖

在 `all_check` 目录执行：

```bash
pip install -r requirement.txt
```

如果你是全新环境，建议先创建虚拟环境再安装。

## 快速开始（Web 门户）

在 `all_check` 目录执行：

```bash
python web_portal.py --host 127.0.0.1 --port 8090
```

浏览器访问：

```text
http://127.0.0.1:8090
```

### 启动参数

- `--host`：监听地址，默认 `127.0.0.1`
- `--port`：端口，默认 `8090`
- `--debug`：开启 Flask 调试模式

## Web 接口

- `GET /`：门户页面
- `POST /api/process`：上传并执行完整流程
- `GET /preview/<job_id>`：预览同步后的 HTML
- `GET /download/<job_id>/<artifact>`：下载产物

`artifact` 可选值：

- `profile_json`
- `synced_html`
- `synced_docx`

## 运行产物

每次任务会生成：

```text
runtime/portal_jobs/jobs/<job_id>/
├─ source_upload.docx
├─ template_upload.docx
├─ source_temp.html
├─ synced_layout.html
├─ layout_profile.json
├─ synced_layout.docx
└─ consistency_result.json
```

## 一致性检查说明

- 一致性检查通过 `html_consistency_checker/check_html_consistency.py` 动态加载 `rule_main_demo/demo.py`
- `demo.py` 必须提供 `call_llm(user_prompt, system_prompt)` 函数
- 若一致性检查失败，主流程仍会返回排版结果，但响应中会包含 `consistency_error`

## 模块单独使用

### 1) DOCX -> HTML

```bash
python docx_to_html_temp/docx_to_html.py --docx input.docx --output temp.html
```

### 2) 同步 DOCX 模板版式到 HTML

```bash
python docx_html_layout_sync/sync_docx_layout_to_html.py \
	--docx template.docx \
	--html temp.html \
	--output synced.html \
	--profile-output layout_profile.json
```

### 3) HTML -> DOCX

```bash
python html_to_docx_sync/html_to_docx.py \
	--html synced.html \
	--output synced.docx \
	--template-docx template.docx
```

### 4) 一致性检查

```bash
python html_consistency_checker/check_html_consistency.py \
	--html synced.html \
	--prompt html_consistency_checker/prompt_consistency.txt \
	--output consistency_result.json
```

## 常见问题

### 1. 上传后提示仅支持 `.docx`

请确认两个上传文件后缀都为 `.docx`。

### 2. 一致性检查报错

通常是 `rule_main_demo/demo.py` 不存在、`call_llm` 未实现或其依赖缺失。

### 3. 端口占用

改用其他端口，例如：

```bash
python web_portal.py --port 8091
```

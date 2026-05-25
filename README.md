# Visio MCP

本项目是一个本地 MCP 服务，通过 Windows COM 自动化操控 Microsoft Visio 绘制流程图。它使用 stdio 传输，适合接入 Codex、Claude Desktop 或其他支持 MCP 的本地客户端。

## 功能

- 创建、打开、保存 Visio 文档
- 根据 `nodes` / `edges` 自动排版并绘制完整流程图
- 逐个添加流程图形状并连接形状
- 列出当前页面形状
- 清空当前页面
- 导出当前页面为 PDF、PNG、SVG、JPG 等 Visio 支持的格式

## 环境要求

- Windows
- 已安装并授权的 Microsoft Visio
- Python 3.10+

## 安装

```powershell
git clone https://github.com/noc228076/visio-mcp-server.git
cd visio-mcp-server
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e .
```

## 运行

```powershell
visio-mcp-server
```

也可以直接运行模块：

```powershell
python -m visio_mcp_server.server
```

## ccswitch 配置示例

把下面这段 JSON 配置加入 ccswitch，并把 `command` 中的 `<repo-path>` 替换为你本机克隆后的项目路径：

```json
{
  "type": "stdio",
  "command": "<repo-path>\\.venv\\Scripts\\python.exe",
  "args": [
    "-m",
    "visio_mcp_server.server"
  ]
}
```

## 推荐工具调用

优先使用 `visio_draw_flowchart` 一次性生成流程图。工具参数直接传入下面的 JSON，不需要外层 `params`：

```json
{
  "nodes": [
    {"id": "start", "text": "开始", "kind": "start_end"},
    {"id": "receive", "text": "接收申请", "kind": "process"},
    {"id": "valid", "text": "资料完整?", "kind": "decision"},
    {"id": "approve", "text": "审批通过", "kind": "process"},
    {"id": "fix", "text": "补充资料", "kind": "process"},
    {"id": "end", "text": "结束", "kind": "start_end"}
  ],
  "edges": [
    {"from": "start", "to": "receive"},
    {"from": "receive", "to": "valid"},
    {"from": "valid", "to": "approve", "text": "是"},
    {"from": "valid", "to": "fix", "text": "否"},
    {"from": "fix", "to": "receive"},
    {"from": "approve", "to": "end"}
  ],
  "direction": "top_to_bottom",
  "page_name": "审批流程",
  "clear_page": true,
  "auto_size_page": true
}
```

常用工具：

- `visio_get_status`
- `visio_create_document`
- `visio_open_document`
- `visio_draw_flowchart`
- `visio_add_shape`
- `visio_connect_shapes`
- `visio_list_shapes`
- `visio_save_document`
- `visio_export_page`
- `visio_clear_page`

## 说明

服务使用 Visio 页面绘图 API 直接绘制矩形、菱形、椭圆、连接线和文本，不依赖特定语言版本的 Visio stencil 名称，因此中文和英文 Office 环境都更容易运行。

坐标单位为 Visio 内部单位英寸。`visio_draw_flowchart` 会自动布局，手工使用 `visio_add_shape` 时需要提供中心点坐标 `x`、`y`。

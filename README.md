# ExcelForge

ExcelForge v2.2 现在采用单一 MCP Host 入口：`excel-mcp`。

项目不再推荐、也不再保留多 MCP 并行启动方式。所有能力都通过同一个 Host 暴露，按 `profile` 或 `bundle` 做装配：
- 一个 Host
- 一个共享 Runtime
- 一个共享 ExcelWorker / WorkbookRegistry
- 按需暴露工具，而不是拆成多个旧入口

## 当前推荐方式

### 1. 安装

```bash
uv sync
```

### 2. 启动统一 Host

```bash
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile basic_edit
```

### 3. 常用 Profile

| Profile | 适用场景 | 工具数 | 启用的 Bundle | 包含的工具域 |
| --- | --- | --- | --- | --- |
| `basic_edit` | 日常编辑（推荐新手） | 31 | foundation + edit | server / workbook / names / sheet / range |
| `calc_format` | 公式与格式处理 | 42 | foundation + edit + calc_format | 上述 + formula / format |
| `automation` | VBA 自动化与恢复 | 36 | foundation + automation + recovery | server / workbook / names / vba / recovery |
| `data_workflow` | Power Query / 数据流 | 28 | foundation + data + analysis | server / workbook / names / pq / analysis |
| `reporting` | 报表与分析 | 28 | foundation + report + analysis | server / workbook / names / chart / pivot / model / audit |
| `all` | 全量开发调试 | **64** | 全部 9 个 bundle | 全部 12 个域 |

> **如何选择 Profile？**
>
> - 只做**读写单元格、管理工作表** → `basic_edit` 够用
> - 需要**写公式、设格式** → `calc_format`
> - 需要**VBA 宏操作**（同步模块、执行宏等）→ `automation` 或 `all`
> - 需要**图表、数据透视表** → `reporting`（chart/pivot/model 当前为 experimental）
> - **不确定或开发调试** → 直接用 `all`，加载全部 64 个工具
>
> **切换 Profile 方法**：修改 MCP 客户端配置中的 `--profile` 参数值，然后重启 MCP 服务。
> 例如将 `"basic_edit"` 改为 `"all"` 即可启用所有工具。

#### 各 Profile 工具域详细对照

| 工具域 | basic_edit | calc_format | automation | data_workflow | reporting | all |
| --- | :---: | :---: | :---: | :---: | :---: | :---: |
| server (2) | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| workbook (6) | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| names (4) | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| sheet (8) | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |
| range (11) | ✅ | ✅ | ❌ | ❌ | ❌ | ✅ |
| formula (4) | ❌ | ✅ | ❌ | ❌ | ❌ | ✅ |
| format (7) | ❌ | ✅ | ❌ | ❌ | ❌ | ✅ |
| vba (8) | ❌ | ❌ | ✅ | ❌ | ❌ | ✅ |
| recovery (8) | ❌ | ❌ | ✅ | ❌ | ❌ | ✅ |
| analysis (1) | ❌ | ❌ | ❌ | ✅ | ✅ | ✅ |
| pq (5) | ❌ | ❌ | ❌ | ✅ | ❌ | ✅ |
| chart/pivot/model | ❌ | ❌ | ❌ | ❌ | ⚠️ 空 | ⚠️ 空 |

### 4. 如果只想单独装备某些功能

优先用 `profile`；如果现成 profile 不完全匹配，再用 `--enable-bundle` / `--disable-bundle` 微调。

示例：

```bash
# 只做 VBA 与恢复
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile automation

# 基础编辑 + 额外启用 VBA
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile basic_edit --enable-bundle automation

# 全量开发，但临时关闭 report
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile all --disable-bundle report
```

对应的客户端示例见：
- `mcp.example.json`
- `examples/mcp.basic_edit.example.json`
- `examples/mcp.automation.example.json`
- `examples/mcp.custom-bundles.example.json`
- `codex-mcp-snippet.toml`

## 配置文件

必须保留：
- `excel-mcp.yaml`：统一 Host 配置
- `runtime-config.yaml`：Runtime 配置
- `excelforge/gateway/profiles.yaml`：Profile 定义
- `excelforge/gateway/bundles.yaml`：Bundle 定义

已移除：
- `excel-core-mcp.yaml`
- `excel-vba-mcp.yaml`
- `excel-recovery-mcp.yaml`
- `excel-pq-mcp.yaml`
- 多 MCP 快速启动文档与旧 gateway 入口包装器

## MCP 客户端示例

### 通用 JSON

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": [
        "run",
        "python",
        "-m",
        "excelforge.gateway.host",
        "--config",
        "YOUR_PROJECT_PATH/excel-mcp.yaml",
        "--profile",
        "all"
      ],
      "cwd": "YOUR_PROJECT_PATH",
      "_comment": "将 --profile 改为所需 profile（basic_edit / calc_format / automation / all 等），然后重启 MCP 服务"
    }
  }
}
```

### Codex / TOML

```toml
[mcp_servers.excelforge]
command = "uv"
args = ["run", "python", "-m", "excelforge.gateway.host", "--config", "YOUR_PROJECT_PATH/excel-mcp.yaml", "--profile", "basic_edit"]
cwd = "YOUR_PROJECT_PATH"
startup_timeout_sec = 30
tool_timeout_sec = 180
```

## 常用命令

```bash
# 查看所有 profile
uv run python -m excelforge.gateway.host --list-profiles

# 查看所有 bundle
uv run python -m excelforge.gateway.host --list-bundles

# 打印 Runtime endpoint 便于诊断
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile basic_edit --print-runtime-endpoint

# 启动 Runtime（仅调试 Runtime 时需要）
uv run python -m excelforge.runtime --config runtime-config.yaml
```

## 架构说明

当前架构是：

```text
MCP Client
  -> excel-mcp (统一 Host)
  -> Shared Runtime
  -> ExcelWorker
  -> Excel Desktop / WorkbookRegistry / Services
```

`profile` 和 `bundle` 只影响工具暴露面，不影响 Runtime identity，不会再因为“分多个旧入口”而把 workbook 句柄拆散到不同 Runtime。

## 版本说明

- `v2.1`：稳定共享 Runtime 的 ready / warmup / health / timeout 基线
- `v2.2`：在 `v2.1` 基线上收敛为单一 Host，并引入 profile / bundle 装配

## 文档

- `设计文档V2.1 Runtime 启动预热与超时治理.md`
- `设计文档V2.2  Profile 与 Bundle 工具装配优化（修订版）.md`
- `V2.X开发记录文档.md`

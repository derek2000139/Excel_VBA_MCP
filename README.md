# ExcelForge

ExcelForge v2.3 采用单一 MCP Host 入口：`excel-mcp`。

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

| Profile         | <br /> | 适用场景              | 工具数    | 启用的 Bundle                         | 包含的工具域                                                    |
| --------------- | :----- | ----------------- | ------ | ---------------------------------- | --------------------------------------------------------- |
| `basic_edit`    | <br /> | 日常编辑（推荐新手）        | 31     | foundation + edit                  | server / workbook / names / sheet / range                 |
| `calc_format`   | <br /> | 公式与格式处理           | 42     | foundation + edit + calc\_format   | 上述 + formula / format                                     |
| `automation`    | <br /> | VBA 自动化与恢复        | 36     | foundation + automation + recovery | server / workbook / names / vba / recovery                |
| `data_workflow` | <br /> | Power Query / 数据流 | 28     | foundation + data + analysis       | server / workbook / names / pq / analysis                 |
| `reporting`     | <br /> | 报表与分析             | 28     | foundation + report + analysis     | server / workbook / names / chart / pivot / model / audit |
| `all`           | <br /> | 全量开发调试            | **64** | 全部 9 个 bundle                      | 全部 12 个域                                                  |

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

| 工具域               | basic\_edit | calc\_format | automation | data\_workflow | reporting |  all |
| ----------------- | :---------: | :----------: | :--------: | :------------: | :-------: | :--: |
| server (2)        |      ✅      |       ✅      |      ✅     |        ✅       |     ✅     |   ✅  |
| workbook (6)      |      ✅      |       ✅      |      ✅     |        ✅       |     ✅     |   ✅  |
| names (4)         |      ✅      |       ✅      |      ✅     |        ✅       |     ✅     |   ✅  |
| sheet (8)         |      ✅      |       ✅      |      ❌     |        ❌       |     ❌     |   ✅  |
| range (11)        |      ✅      |       ✅      |      ❌     |        ❌       |     ❌     |   ✅  |
| formula (4)       |      ❌      |       ✅      |      ❌     |        ❌       |     ❌     |   ✅  |
| format (7)        |      ❌      |       ✅      |      ❌     |        ❌       |     ❌     |   ✅  |
| vba (8)           |      ❌      |       ❌      |      ✅     |        ❌       |     ❌     |   ✅  |
| recovery (8)      |      ❌      |       ❌      |      ✅     |        ❌       |     ❌     |   ✅  |
| analysis (1)      |      ❌      |       ❌      |      ❌     |        ✅       |     ✅     |   ✅  |
| pq (5)            |      ❌      |       ❌      |      ❌     |        ✅       |     ❌     |   ✅  |
| chart/pivot/model |      ❌      |       ❌      |      ❌     |        ❌       |    ⚠️ 空   | ⚠️ 空 |

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
        "all",
        "--restart-runtime",
        "always"
      ],
      "cwd": "YOUR_PROJECT_PATH",
      "_comment": "将 --profile 改为所需 profile（basic_edit / calc_format / automation / all 等），--restart-runtime 改为 always（开发）或 if-stale（生产）"
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

## 日志文件

ExcelForge 会将日志写入 `~/.excelforge/logs/excelforge_YYYYMMDD.log`，包含：

- Gateway 和 Runtime 的所有操作日志
- 工具调用记录（TOOL CALL / TOOL OK / TOOL FAIL）
- Excel 进程创建和销毁记录
- 错误和警告信息

查看日志：

```powershell
# Windows
Get-Content "$env:USERPROFILE\.excelforge\logs\excelforge_20260329.log" | Select-Object -Last 50

# Linux/Mac
cat ~/.excelforge/logs/excelforge_20260329.log | tail -50
```

日志文件会自动清理，保留最近 30 天的记录。

## 依赖说明

- `psutil>=5.9.0`：用于 Excel 进程管理和僵尸进程清理（Windows/Linux/Mac 跨平台）

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
- `v2.3`：Worker 治理能力产品化、配置化、可观测化

### V2.3 新特性：Worker 健康状态

V2.3 引入了完整的 Worker 健康状态机，用于治理 Excel 实例生命周期。

#### Worker 健康状态枚举

| 状态          | 说明                        |
| ----------- | ------------------------- |
| `HEALTHY`   | Worker 正常运行，受控 Excel 实例健康 |
| `DEGRADED`  | Worker 降级中，正在执行恢复         |
| `STALE`     | 受控 Excel 实例已失效            |
| `RECYCLING` | Worker 正在回收/重建中           |
| `FAILED`    | Worker 处于失败状态             |

#### server.health 返回字段

`server.health` 接口新增完整 Worker 状态返回：

```json
{
  "excel": {
    "ready": true,
    "excel_pid": 12345,
    "version": "16.0"
  },
  "worker": {
    "state": "HEALTHY",
    "excel_pid": 12345,
    "open_workbooks": 2,
    "operation_count": 150,
    "high_risk_operation_count": 3,
    "exception_count": 0,
    "uptime_seconds": 3600.5,
    "last_exception_type": null,
    "last_recycle_reason": null,
    "last_pid_change_reason": null
  }
}
```

#### 单受控实例约束

V2.3 强化了**单受控实例约束**：

- 单个 Runtime 内默认只允许存在 **1 个受控 Excel Application 实例**
- hidden → visible 状态切换必须保持同一 PID
- `ensure_app()` 在实例有效时必须直接复用，不允许创建第二个实例

#### PID 变化原因

当 `excel_pid` 发生变化时，`last_pid_change_reason` 会记录原因：

| 原因值               | 说明               |
| ----------------- | ---------------- |
| `stale_rebuild`   | 因实例失效触发的自动重建     |
| `worker_recycle`  | 手动触发的 Worker 回收  |
| `runtime_restart` | Runtime 重启后的首次创建 |

#### rebuild() 默认行为

`rebuild()` 方法（用于手动回收 Excel 实例）默认**不自动重新打开工作簿**：

- 记录当前打开的工作簿路径
- 退出当前 Excel 实例
- 重建 COM 对象
- 下次请求时需要用户显式调用 `open_file`

如需自动 reopen，可在调用时传入 `reopen_workbooks=True`（不推荐，默认关闭）。

### --restart-runtime 策略推荐

| 场景   | 推荐策略                         | 说明                      |
| ---- | ---------------------------- | ----------------------- |
| 开发调试 | `--restart-runtime always`   | 每次启动都重启 Runtime，确保新代码生效 |
| 生产环境 | `--restart-runtime if-stale` | 仅在 Runtime 过期时重启，提升复用率  |

## Trae AI 工具截断问题与解决对策

### 问题现象

使用 Trae AI 等 MCP 客户端时，可能会遇到以下问题：

- 某些工具（如 VBA 工具）无法识别或调用
- 明明定义了工具，但调用时报错 "Tool is not available"
- `server.health` 显示正常，但实际调用失败

### 问题根因

**MCP 客户端存在工具数量限制**，当 profile 包含的工具数量超过限制时，工具列表会被截断。

例如 `all` profile 有 64 个工具，如果客户端限制为 40 个，则后半部分工具（如 VBA、recovery 等）会被截断，导致无法调用。

### 验证方法

1. 打开 Trae AI 的工具列表
2. 搜索 `vba.` 或 `backup.` 等关键字
3. 如果显示"未找到工具"，说明该工具被截断了

或者通过 Python 脚本检查 profile 实际加载的工具：

```python
from excelforge.gateway.profile_resolver import ProfileResolver, BundleRegistry

resolver = ProfileResolver()
registry = BundleRegistry()
info = resolver.get_profile_info('all')
tools = registry.get_all_tools(info['bundles'])
print(f'Total tools in all profile: {len(tools)}')
for t in sorted(tools):
    print(f'  {t}')
```

### 解决对策

#### 方案一：创建自定义 Profile（推荐）

修改 `excelforge/gateway/profiles.yaml`，将需要的工具 bundle 放在前面，并设置合理的 `tool_budget`：

```yaml
profiles:
  # ... 其他 profile ...

  # 调试测试专用 - VBA/Recovery 核心工具
  vba_first:
    description: "调试测试专用 - VBA/Recovery核心工具"
    bundles:
      - foundation     # 12工具 (server/workbook/names) ← 放最前面
      - automation     # 8工具 (VBA) ← 优先保证 VBA 可用
      - recovery       # 8工具 (rollback/snapshot/backups)
    tool_budget: 30    # 根据实际需要调整
    risk_level: medium
```

然后在 Trae AI 配置中使用 `--profile vba_first`。

#### 方案二：调整 tool\_budget

如果某个 profile 部分工具被截断，可以降低 `tool_budget` 使其刚好包含目标工具：

```yaml
automation:
  description: "自动化 - VBA 与恢复工具"
  bundles:
    - foundation
    - automation
    - recovery
  tool_budget: 28   # foundation(12) + automation(8) + recovery(8) = 28
```

#### 方案三：使用 bundle 开关

如果只想在现有 profile 基础上增加特定工具：

```bash
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile basic_edit --enable-bundle automation
```

这会在 `basic_edit` 基础上额外启用 `automation` bundle。

### 验证修复

修改配置后，重启 Trae AI MCP 服务，然后：

1. 检查工具列表是否包含目标工具（如 `vba.inspect_project`）
2. 调用 `vba.execute` 执行一个简单 VBA 宏测试

```python
# 测试 VBA 是否可用
vba.inspect_project(workbook_id="your_workbook_id")
# 如果返回 VBAProject 信息，说明工具已加载成功
```

### Profile 工具数量参考

| Profile     | 工具数  | tool\_budget | 说明             |
| ----------- | ---- | ------------ | -------------- |
| basic\_edit | 31   | 35           | 基础编辑够用         |
| automation  | 36   | 40           | VBA + Recovery |
| vba\_first  | \~28 | 30           | 调试专用           |

## 文档

- `设计文档V2.1 Runtime 启动预热与超时治理.md`
- `设计文档V2.2  Profile 与 Bundle 工具装配优化（修订版）.md`
- `V2.X开发记录文档.md`


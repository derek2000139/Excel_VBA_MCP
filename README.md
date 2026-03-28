# ExcelForge

**ExcelForge** 是一个基于 MCP (Model Context Protocol) 的 Excel 操作工具集，让 AI 助手能够安全、高效地操作 Excel 文件。

## 功能特性

- 工作簿管理 - 打开、保存、关闭 Excel 文件
- 工作表操作 - 创建、删除、重命名、查看结构
- 范围操作 - 读写数据、复制、清除、排序
- 公式支持 - 验证表达式、填充范围
- 格式设置 - 设置单元格样式、自动调整列宽
- 备份恢复 - 文件级备份、快照回滚
- VBA 读写访问 - 查看工程、扫描代码、同步模块、执行宏
- 命名范围 - 列出和读取命名区域
- 数据验证和条件格式 - 读取工作表规则

## 版本

当前版本: **v2.1.0**

## 系统要求

- Windows 10/11
- Microsoft Excel Desktop（2016/2019/365）
- Python >= 3.11
- `uv` 包管理器

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/derek2000139/Excel_VBA_MCP.git
cd ExcelForge
```

### 2. 安装依赖

```bash
uv sync --extra dev
```

### 3. 配置（重要！）

复制配置模板并修改：

```bash
copy config.example.yaml config.yaml
```

编辑 `config.yaml`，设置 `allowed_roots` 为 `*` 可以访问任意路径的 Excel：

```yaml
paths:
  allowed_roots:
    - "*"  # 允许访问所有路径
    # 或指定路径：如 "C:/Users/你的用户名/Documents"
```

### 4. 启动服务

```bash
uv run python -m excelforge --config ./config.yaml serve
```

### 5. 在 IDE 中配置 MCP

参考 `mcp.example.json` 模板配置你的 IDE（MCP 客户端）：

```json
{
  "mcpServers": {
    "ExcelForge": {
      "command": "uv",
      "args": [
        "run", "python", "-m", "excelforge",
        "--config", "你的ExcelForge项目路径/config.yaml",
        "serve"
      ],
      "cwd": "你的ExcelForge项目路径"
    }
  }
}
```

## 工具列表

v1.0.1 提供两种 Profile：
- **Default Profile**：约 30 个工具（Trae IDE 推荐）
- **Extended Profile**：约 36 个工具（全功能）

### Default Profile 工具（30个）

| 类别 | 工具数 | 主要工具 |
|------|--------|----------|
| 工作簿 | 5 | open_file, inspect, save_file, close_file, create_file |
| 工作表 | 6 | inspect_structure, create_sheet, rename_sheet, delete_sheet, set_auto_filter, get_rules |
| 范围 | 10 | read_values, write_values, clear_contents, copy_range, manage_rows/columns, sort_data, merge_cells |
| 公式 | 4 | fill_range, set_single, get_dependencies, repair_references |
| 格式 | 1 | manage |
| VBA | 7 | inspect_project, get_module_code, scan_code, sync_module, remove_module, execute, compile |

### Extended Profile 额外工具（+6）

- rollback.manage - 快照管理
- backups.manage - 备份管理
- audit.list_operations - 审计日志
- names.inspect/create/delete - 命名范围

## 路径配置（重要）

`config.yaml` 中的 `paths.allowed_roots` 控制可以访问哪些路径：

```yaml
paths:
  allowed_roots:
    - "*"           # 允许所有路径（安全风险高）
    - "C:/Users"    # 或指定具体路径
    - "D:/Work"
```

| 配置 | 说明 |
|------|------|
| `"*"` | 允许访问任意路径的 Excel 文件 |
| `"C:/Users/derek"` | 仅允许访问指定用户的目录 |
| `"D:/Project"` | 仅允许访问指定项目目录 |

**安全建议**：测试完成后，建议限制为具体工作目录，避免访问任意文件。

## VBA 安全策略

- 写入 VBA 代码必须通过安全扫描
- 默认阻止 CRITICAL 和 HIGH 风险代码（如 `Shell`, `CreateObject("WScript.Shell")` 等）
- `MsgBox` 自动替换为 `Debug.Print`
- `InputBox` 被禁用以避免弹窗阻塞

## 常见问题

### Q1: 提示"文件路径不允许"
修改 `config.yaml`，将 `allowed_roots` 改为 `"*"` 或添加目标路径。

### Q2: 提示"Python not found"
确保 Python 已安装并添加到 PATH，重启终端后尝试。

### Q3: 提示"uv: command not found"
```bash
pip install uv
```

### Q4: 宏执行失败，提示"宏不可用"
通常是 Excel Worker 状态异常。执行 `server.get_status` 检查 `excel_worker.state` 是否为 `"running"`，如果不是则重启 MCP 服务。

### Q5: Trae IDE 工具数量过多不稳定
使用 `default` profile（已配置），工具数量约 30 个。

## 版本历史

| 版本 | 日期 | 主要更新 |
|------|------|----------|
| v0.1-v0.5 | 2026-03-22~24 | 基础功能开发 |
| v1.0.0 | 2026-03-24 | 工具组配置、简化工具链、Trae兼容 |
| v1.0.1 | 2026-03-24 | 工具合并、Profile机制、Worker健康检查、通配符路径 |
| v1.0.2 | 2026-03-26 | workbook.inspect返回增加index字段、修复三元表达式为if-else、CLAUDE.md增加MCP设计决策说明 |
| **v2.0.0** | **2026-03-27** | **Runtime重架构、ExcelWorker生命周期管理、named pipe通信、snapshot/rollback/backup三大恢复机制** |
| **v2.1.0** | **2026-03-28** | **Runtime启动预热机制、Windows探活API修复、Dispatcher就绪拦截、server.health健康检查、首次请求0.43秒响应** |

## 详细文档

更多详细信息请参考 [MCP使用说明.md](MCP使用说明.md)：
- 环境配置详解
- Profile 配置说明
- 工具调用流程
- 错误码说明
- VBA 安全规则详解
- 迁移指南

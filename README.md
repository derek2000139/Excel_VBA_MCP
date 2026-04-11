# ExcelForge

基于 MCP 协议的 Excel 自动化工具服务。单一 Host 入口，通过 Profile / Bundle 按场景装配工具。

当前版本 **v2.5.0** — 18 个能力域 · 101+ 个工具 · 14 个 Bundle · 12 个 Profile（含 5 个开发 Profile）。

## 安装

### 前置要求

| 项目 | 要求 |
|------|------|
| Python | ≥ 3.11 |
| 操作系统 | Windows（需要 Excel COM 自动化） |
| Excel | Windows 版 Microsoft Excel |
| 包管理 | uv |

### 第一步：安装 uv（如尚未安装）

```powershell
pip install uv
```

### 第二步：克隆仓库

```powershell
git clone <你的仓库地址>
cd ExcelForge
```

### 第三步：配置

```powershell
cp config.example.yaml runtime-config.yaml
```

### 第四步：安装依赖

```powershell
uv sync
```

### 第五步：快速验证

```powershell
# 查看可用 Profile
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --list-profiles

# 查看可用 Bundle
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --list-bundles

# 查看当前 Profile 的工具清单
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile basic_edit --dump-tools
```

## 启动

```bash
uv sync
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile basic_edit
```

## 选择 Profile

Profile 决定暴露哪些工具。按场景选，不确定就从 `basic_edit` 开始。

| Profile | 做什么 | 工具数 |
|---------|-------|------:|
| `basic_edit` | 打开 / 读写 / Sheet 管理 | 35 |
| `calc_format` | 上述 + 公式 + 格式 | 46 |
| `automation` | VBA + 快照 / 备份 / 回滚 | 40 |
| `data_workflow` | PQ 查询 + Table + 分析审计 + workbook_ops | 33 |
| `reporting` | 导出 PDF / CSV + 分析 | 32 |
| `artifact_extract` | Package 直接解析（无需 Excel COM）| 16 |
| `artifact_transform` | artifact_extract + patch / merge / compare | 22 |
| `batch_delivery` | artifact_transform + 批量处理 | 26 |
| `all` | 全部工具（仅 CLI / 回归测试） | 101+ |

**开发 Profile**（推荐 AI 编程使用）：

| Profile | 做什么 | 工具数 |
|---------|-------|------:|
| `artifact_first` | Package 优先，COM 兜底 | 42 |
| `package_first` | 批量处理优先 | 48 |
| `batch_first` | 批量操作优先 | 50 |
| `chart_first` | 图表操作优先 | 44 |
| `rules_first` | 条件格式优先 | 42 |

切换方法：改 `--profile` 参数，重启服务。

## 微调 Bundle

Profile 不完全匹配时，用 `--enable-bundle` / `--disable-bundle` 加减：

```bash
# data_workflow + 结构编辑
--profile data_workflow --enable-bundle edit_structure  # 33 + 6 = 39

# 基础编辑 + 结构编辑
--profile basic_edit --enable-bundle edit_structure  # 35 + 6 = 41
```

可用 Bundle：

| Bundle | 工具数 | 内容 |
|--------|------:|------|
| foundation | 8 | 服务状态 + 工作簿 I/O（必选） |
| data | 8 | Table 管理 |
| analysis | 6 | 结构扫描 / 公式审计 / 分析报告 |
| workbook_ops | 6 | 另存 / 刷新 / 计算 / 导出 PDF·CSV |
| edit_basic | 7 | Sheet 创建/重命名 + Range 读写/复制 |
| edit_structure | 6 | Sheet 复制/移动/隐藏 + Range 查找替换/自动调整 |
| calc_format | 11 | 公式 + 格式 |
| automation | 8 | VBA |
| recovery | 8 | 快照 / 回滚 / 备份 |
| artifact_export | 8 | Package XML 解析（无需 Excel COM）|
| artifact_patch | 7 | Package 补丁 / 替换 / 合并 / 比较 |
| batch_ops | 3 | 批量提取 / 变换 / 比较 |
| chart_operations | 4 | 图表列表 / 详情 / 系列 / 导出 |
| format_rules | 4 | 条件格式规则应用 / 更新 / 复制 / 清除 |

## 常用命令

```bash
# 查看可用 Profile / Bundle
--list-profiles
--list-bundles

# 诊断当前 Profile 的工具清单
--dump-tools
--dump-tools-with-index

# 查看 Profile 解析过程
--dump-profile-resolution
```

完整示例：

```bash
uv run python -m excelforge.gateway.host --config excel-mcp.yaml --profile basic_edit --dump-tools
```

## MCP 客户端配置

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": [
        "run", "python", "-m", "excelforge.gateway.host",
        "--config", "YOUR_PROJECT_PATH/excel-mcp.yaml",
        "--profile", "data_workflow",
        "--restart-runtime", "if-stale"
      ],
      "cwd": "YOUR_PROJECT_PATH/"
    }
  }
}
```

- 开发环境用 `--restart-runtime always`，生产用 `if-stale`
- 更多示例见 `mcp.example.json` 及 `examples/` 目录

## 文档

| 文档 | 内容 |
|------|------|
| [工具域 Profile 参考手册](docs/ExcelForge%20V2.5%20%E2%80%94%20%E5%B7%A5%E5%85%B7%E5%9F%9F%20Profile%20%E5%8F%82%E8%80%83%E6%89%8B%E5%86%8C.md) | Tool / Bundle / Profile 完整对照矩阵与查询索引（V2.5） |
| [Trae AI 使用说明](docs/clients/trae_usage.md) | Trae 推荐 Profile、截断问题与配置示例 |
| [v2.4 变更日志](docs/changelog/v2.4.md) | 本版新增工具、Bundle 拆分、Profile 重整记录 |
| [v2.5 变更日志](docs/changelog/v2.5.md) | 新增 Package/Batch/Chart/ConditionalFormat 工具域 |
| [OOXML Package 使用指南](docs/features/OOXML_Package_%E4%BD%BF%E7%94%A8%E6%8C%87%E5%8D%97.md) | 直接解析 Excel 文件 ZIP 结构，无需 Excel COM |
| [Batch 批量处理指南](docs/features/Batch_%E6%89%B9%E9%87%8F%E5%A4%84%E7%90%86%E6%8C%87%E5%8D%97.md) | 批量提取、变换、比较多个 Excel 文件 |
| [Chart 图表指南](docs/features/Chart_%E5%9B%BE%E8%A1%A8%E6%8C%87%E5%8D%97.md) | 图表解析、导出、系列读取 |

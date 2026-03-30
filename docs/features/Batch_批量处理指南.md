# ExcelForge v2.5 Batch 批量处理指南

**版本**：v2.5.0

---

## 概述

Batch 工具用于对多个 Excel 文件进行批量操作，无需启动 Excel COM。

---

## 工具列表

| 工具 | 功能 |
|------|------|
| `package.batch_extract_parts` | 批量提取多个文件的指定部件 |
| `package.batch_transform` | 批量应用补丁变换 |
| `package.batch_compare` | 批量比较文件集 |

---

## 常用场景

### 场景 1：批量提取部件

```python
# 从多个文件提取 xl/workbook.xml
package.batch_extract_parts(
    file_paths=[
        "file1.xlsx",
        "file2.xlsx",
        "file3.xlsx"
    ],
    part_path="xl/workbook.xml",
    output_dir="./extracted"
)
```

**返回示例**：
```json
{
  "total_count": 3,
  "success_count": 3,
  "failed_count": 0,
  "success_rate": 100.0
}
```

### 场景 2：批量变换

```python
# 对多个文件应用相同的补丁
package.batch_transform(
    file_paths=[
        "report1.xlsx",
        "report2.xlsx",
        "report3.xlsx"
    ],
    patches=[
        {
            "part_path": "xl/workbook.xml",
            "action": "replace",
            "content": "..."
        }
    ],
    output_dir="./transformed"
)
```

### 场景 3：批量比较

```python
# 比较多个文件与参考文件的差异
package.batch_compare(
    file_paths=[
        "output1.xlsx",
        "output2.xlsx",
        "output3.xlsx"
    ],
    reference_file="template.xlsx"
)
```

---

## Profile 配置

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": ["run", "python", "-m", "excelforge.gateway.host", "--profile", "batch_delivery"],
      "cwd": "/path/to/ExcelForge"
    }
  }
}
```

---

## 配置选项

### dry_run

设置为 `true` 可预览操作结果，不实际执行：

```python
package.batch_transform(
    file_paths=["file1.xlsx", "file2.xlsx"],
    patches=[...],
    dry_run=True
)
```

### max_workers

控制并行执行的工作数（默认 4）：

```python
# 在 host.py 中配置
--max-workers 8
```

---

*文档版本：v2.5.0*

# ExcelForge v2.5 OOXML/Package 使用指南

**版本**：v2.5.0

---

## 概述

OOXML（Open Office XML）是 Excel `.xlsx` 文件的内部格式，本质是一个 ZIP 包，包含多个 XML 文件。ExcelForge v2.5 提供了直接读取和操作这些 XML 部件的能力，无需启动 Excel。

### 核心优势

- **无需 Excel**：纯 Python 实现，可在无 Excel 环境运行
- **批量处理**：支持大规模文件处理
- **细粒度控制**：直接读取/修改 XML 部件

---

## 执行模式

| 模式 | 说明 | 是否需要 Excel |
|------|------|---------------|
| `file_package` | 直接操作文件包（ZIP/XML） | 否 |
| `runtime` | 通过 Excel COM 操作 | 是 |

---

## 工具列表

### 读取工具（artifact_export）

| 工具 | 功能 |
|------|------|
| `package.inspect_manifest` | 检查文件包整体结构 |
| `package.list_parts` | 列出所有部件 |
| `package.get_part_xml` | 获取指定部件的 XML 内容 |
| `package.extract_part` | 提取部件到文件 |
| `package.list_media` | 列出媒体资源（图片等） |
| `package.list_custom_xml` | 列出自定义 XML 部件 |
| `package.detect_features` | 检测文件特性 |
| `package.export_manifest` | 导出完整清单报告 |

### 写入工具（artifact_patch）

| 工具 | 功能 |
|------|------|
| `package.clone_with_patch` | 克隆文件并应用补丁 |
| `package.replace_shared_strings` | 批量替换共享字符串 |
| `package.patch_defined_names` | 修补命名范围 |
| `package.update_docprops` | 更新文档属性 |
| `package.merge_template_parts` | 合并模板部件 |
| `package.remove_external_links` | 删除外部链接 |
| `package.compare_workbooks` | 比较两个工作簿 |

---

## 常用场景

### 场景 1：检查 Excel 文件包结构

```python
# 查看文件包含哪些部件
package.inspect_manifest(file_path="report.xlsx")
```

**返回示例**：
```json
{
  "file_path": "report.xlsx",
  "file_size": 15432,
  "part_count": 15,
  "parts": [
    "xl/workbook.xml",
    "xl/worksheets/sheet1.xml",
    "xl/styles.xml",
    ...
  ]
}
```

### 场景 2：检测文件特性

```python
# 检查文件是否包含宏、图表、外部链接等
package.detect_features(file_path="report.xlsx")
```

**返回示例**：
```json
{
  "has_vba": true,
  "has_external_links": false,
  "has_charts": true,
  "has_pivot_tables": false,
  "has_power_query": false
}
```

### 场景 3：提取媒体文件

```python
# 列出所有图片
package.list_media(file_path="report.xlsx")

# 提取第一个图片
package.extract_part(
    file_path="report.xlsx",
    part_path="xl/media/image1.png",
    output_path="./extracted_image.png"
)
```

### 场景 4：更新文档属性

```python
# 修改作者、标题等元信息
package.update_docprops(
    file_path="report.xlsx",
    properties={
        "creator": "自动化脚本",
        "title": "月度报告",
        "subject": "2026年3月"
    },
    output_path="report_updated.xlsx"
)
```

### 场景 5：批量替换字符串

```python
# 替换共享字符串表中的文本（如产品名称变更）
package.replace_shared_strings(
    file_path="report.xlsx",
    replacements={
        "旧产品名": "新产品名",
        "北京分部": "北方区"
    },
    output_path="report_updated.xlsx"
)
```

### 场景 6：比较两个 Excel 文件

```python
# 对比两个文件的差异
package.compare_workbooks(
    file_path1="version1.xlsx",
    file_path2="version2.xlsx"
)
```

**返回示例**：
```json
{
  "has_differences": true,
  "difference_count": 2,
  "details": {
    "modified_parts": [
      {"part": "xl/styles.xml", "size_a": 2550, "size_b": 3063},
      {"part": "xl/worksheets/sheet1.xml", "size_a": 1364, "size_b": 2157}
    ]
  }
}
```

### 场景 7：克隆并应用补丁

```python
# 自定义修改 XML 内容
package.clone_with_patch(
    file_path="report.xlsx",
    patches=[
        {
            "part_path": "xl/workbook.xml",
            "action": "replace",
            "content": "<modified>...</modified>"
        }
    ],
    output_path="report_patched.xlsx"
)
```

---

## Profile 配置

### artifact_extract（推荐用于读取）

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": ["run", "python", "-m", "excelforge.gateway.host", "--profile", "artifact_extract"],
      "cwd": "/path/to/ExcelForge"
    }
  }
}
```

### artifact_transform（用于读取+写入）

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": ["run", "python", "-m", "excelforge.gateway.host", "--profile", "artifact_transform"],
      "cwd": "/path/to/ExcelForge"
    }
  }
}
```

---

## 限制说明

1. **不支持 .xls 格式**：仅支持 `.xlsx`、`.xlsm`、`.xltx`、`.xltm`
2. **不支持带密码保护的文件**
3. **写操作默认生成新文件**：原文件不会被修改
4. **部分修改需要 Excel 重新计算**：如修改公式结果

---

## 常见问题

### Q: 为什么 package 工具不需要 Excel？

A: Excel `.xlsx` 文件本质是 ZIP 包，包含 XML 文件。Package Executor 直接读取/修改这些 XML，无需启动 Excel。

### Q: 修改后需要重新打开 Excel 吗？

A: 对于简单的属性修改（如文档属性），直接用 Excel 打开即可。对于复杂的 XML 修改，可能需要 Excel 重新计算后保存。

### Q: 支持批量处理吗？

A: 支持。使用 `batch_delivery` profile 或调用 batch 工具。

---

*文档版本：v2.5.0*

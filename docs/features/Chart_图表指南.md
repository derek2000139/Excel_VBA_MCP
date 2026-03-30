# ExcelForge v2.5 Chart 图表指南

**版本**：v2.5.0

---

## 概述

Chart 工具用于读取 Excel 文件中的图表信息，无需启动 Excel。

---

## 工具列表

| 工具 | 功能 |
|------|------|
| `chart.list_charts` | 列出工作簿中的所有图表 |
| `chart.inspect` | 检查单个图表的详细信息 |
| `chart.list_series` | 列出图表的数据系列 |
| `chart.export_spec` | 导出图表规格为 JSON |

---

## 常用场景

### 场景 1：列出所有图表

```python
# 查看文件中有哪些图表
chart.list_charts(file_path="report.xlsx")
```

**返回示例**：
```json
{
  "total_charts": 1,
  "charts": [
    {
      "chart_id": "chart1",
      "name": "销售额趋势",
      "sheet": "数据透视",
      "chart_type": "line",
      "has_title": true,
      "has_legend": true
    }
  ]
}
```

### 场景 2：检查图表详情

```python
# 查看特定图表的详细信息
chart.inspect(file_path="report.xlsx", chart_id="chart1")
```

**返回示例**：
```json
{
  "chart_id": "chart1",
  "chart_type": "bar",
  "title": "销售额趋势",
  "has_title": true,
  "has_legend": true,
  "series_count": 2,
  "series": [
    {
      "index": "0",
      "name": "产品A",
      "values_range": "'数据'!$B$2:$B$10",
      "categories_range": "'数据'!$A$2:$A$10"
    }
  ],
  "axes": [
    {"type": "catAx", "title": "月份"},
    {"type": "valAx", "title": "销售额"}
  ]
}
```

### 场景 3：导出图表规格

```python
# 导出图表定义用于备份或分析
chart.export_spec(
    file_path="report.xlsx",
    chart_id="chart1",
    output_path="./chart_spec.json"
)
```

---

## Profile 配置

Chart 工具包含在 `reporting` profile 中：

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": ["run", "python", "-m", "excelforge.gateway.host", "--profile", "reporting"],
      "cwd": "/path/to/ExcelForge"
    }
  }
}
```

或者使用 `artifact_transform` + `report` bundle：

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": ["run", "python", "-m", "excelforge.gateway.host", "--profile", "artifact_transform", "--enable-bundle", "report"],
      "cwd": "/path/to/ExcelForge"
    }
  }
}
```

---

## 图表类型

支持的图表类型：

- `bar` - 柱状图/条形图
- `line` - 折线图
- `pie` - 饼图
- `scatter` - 散点图
- `area` - 面积图
- `stock` - 股价图
- `radar` - 雷达图
- `doughnut` - 环形图

---

## 限制说明

1. **仅支持读取**：目前图表工具只支持读取，不支持创建或修改图表
2. **不支持图表样式**：暂不支持读取图表的样式/配色信息
3. **不支持图表数据表**：暂不支持读取图表下方的数据表

---

*文档版本：v2.5.0*

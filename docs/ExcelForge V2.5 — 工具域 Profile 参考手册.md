# ExcelForge V2.5 — 工具域 Profile 参考手册

**产品版本**：v2.5.0 · **工具总数**：101 · **Bundle 数**：15 · **Profile 数**：17

---

## 1. 核心概念

| 概念 | 定义 | 粒度 | 数量 |
|------|------|------|-----:|
| **Domain（能力域）** | 工具的逻辑分类命名空间。一个工具属于且仅属于一个 Domain。Domain 本身不参与装配，仅用于语义归类。 | 最细 | 22 |
| **Tool（工具）** | 最小可调用能力单元，由 Domain 命名空间 + 动作名组成，如 `workbook.open_file`。 | — | 101 |
| **Bundle（能力包）** | 一组语义相关工具的可复用集合。Bundle 是装配的基本单位，Profile 通过组合 Bundle 来确定最终工具集。一个 Bundle 可跨 Domain 取工具。 | 中间 | 15 |
| **Profile（场景配置）** | 面向用户任务场景的 Bundle 组合。Profile 回答"用户要完成什么任务"，决定运行时实际注册哪些工具。 | 最粗 | 17 |

**一句话关系**：Domain 给工具分类 → Bundle 把工具打包 → Profile 把包组合成场景。

抽象层级图
```
Profile（场景入口）
└─ Bundle（装配单元）
   └─ Tool（具体能力）
      └─ Domain（语义归属）
```
注意：严格来说，Domain 不是 Tool 的"下一级执行层"，
而是 Tool 的"语义标签 / 分类归属"。
所以上图是便于理解的记忆图。


---

## 2. 层级架构图

```
Profile 层（场景交付）
┌────────────────────────────────────────────────────────────────────────────────┐
│  basic_edit    calc_format    automation    data_workflow    reporting          │
│   15 tools      30 tools      25 tools       27 tools       24 tools            │
│                                                                                 │
│  artifact_extract  artifact_transform  batch_delivery                           │
│    16 tools         26 tools           30 tools                                 │
└───────┬─────────────┬──────────────┬─────────────┬───────────────────────────────┘
        │             │              │             │
        ▼             ▼              ▼             ▼
Bundle 层（能力复用）
┌────────────────────────────────────────────────────────────────────────────────┐
│  foundation   edit_basic   calc_format   automation   recovery   data            │
│   8 tools      7 tools    11 tools      9 tools     8 tools   13 tools          │
│                                                                                 │
│  artifact_export  artifact_patch   batch_ops   format_rules   chart            │
│    8 tools        7 tools        3 tools      4 tools      4 tools          │
│              workbook_ops   analysis   report                                │
│                6 tools      6 tools    4 tools                                │
└───────┬────────┬──────────┬────────────┬────────────┬────────────┬───────────────┘
        │        │          │            │            │            │
        ▼        ▼          ▼            ▼            ▼            ▼
Tool 层（最小能力单元）
┌────────────────────────────────────────────────────────────────────────────────┐
│  server.health  workbook.open_file  range.read_values  vba.execute              │
│  package.inspect_manifest  chart.list_charts  format.apply_conditional_rule     │
│  ...  (×101)                                                                  │
└───────┬────────────┬──────────────────┬──────────────────────────────────────────┘
        │            │                  │
        ▼            ▼                  ▼
Domain 层（逻辑分类，不参与装配）
┌────────────────────────────────────────────────────────────────────────────────┐
│  server   workbook   names   sheet_basic   range_basic   formula   format        │
│  sheet_structure   range_structure   conditional_format   analysis   vba        │
│  recovery   pq   table   workbook_ops   chart   package   package_patch   batch │
└────────────────────────────────────────────────────────────────────────────────┘
```

---

## 3. 执行模式说明

| 执行模式 | 说明 | 是否需要 Excel |
|---------|------|--------------|
| **runtime** | 通过 Excel COM 执行 | 必须 |
| **file_package** | 直接解析 Excel 文件 ZIP 结构 | 不需要 |
| **batch** | 批量处理多个文件 | 不需要 |

**V2.5 新增工具全部为 file_package 或 batch 模式**，无需启动 Excel COM。

---

## 4. 全量对照矩阵

### 4.1 正向查询：Domain → Tool → Bundle → Profile 可用性

图例：● = 默认包含　○ = 可通过 `--enable-bundle` 加入　· = 不在该 Profile 任何可选路径中

| # | Domain | Tool | Bundle | basic_ edit | calc_ format | auto- mation | data_ work | report- ing | artifact_ extract | artifact_ trans | batch_ delivery |
|--:|--------|------|--------|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| 1 | server | `server.health` | foundation | ● | ● | ● | ● | ● | ● | ● | ● |
| 2 | server | `server.get_status` | foundation | ● | ● | ● | ● | ● | ● | ● | ● |
| 3 | workbook | `workbook.open_file` | foundation | ● | ● | ● | ● | ● | ● | ● | ● |
| 4 | workbook | `workbook.create_file` | foundation | ● | ● | ● | ● | ● | ● | ● | ● |
| 5 | workbook | `workbook.save_file` | foundation | ● | ● | ● | ● | ● | ● | ● | ● |
| 6 | workbook | `workbook.close_file` | foundation | ● | ● | ● | ● | ● | ● | ● | ● |
| 7 | workbook | `workbook.list_open` | foundation | ● | ● | ● | ● | ● | ● | ● | ● |
| 8 | workbook | `workbook.inspect` | foundation | ● | ● | ● | ● | ● | ● | ● | ● |
| 9 | workbook | `workbook.save_as` | workbook_ops | ○ | ○ | ○ | ○ | ● | ○ | ○ | ○ |
| 10 | workbook | `workbook.refresh_all` | workbook_ops | ○ | ○ | ○ | ○ | ● | ○ | ○ | ○ |
| 11 | workbook | `workbook.calculate` | workbook_ops | ○ | ○ | ○ | ○ | ● | ○ | ○ | ○ |
| 12 | workbook | `workbook.list_links` | workbook_ops | ○ | ○ | ○ | ○ | ● | ○ | ○ | ○ |
| 13 | workbook | `workbook.export_pdf` | workbook_ops | ○ | ○ | ○ | ○ | ● | ○ | ○ | ○ |
| 14 | names | `names.inspect` | names | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 15 | names | `names.manage` | names | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 16 | names | `names.create` | names | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 17 | names | `names.delete` | names | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 18 | sheet_basic | `sheet.create_sheet` | edit_basic | ● | ● | ○ | ○ | ○ | ○ | ○ | ○ |
| 19 | sheet_basic | `sheet.rename_sheet` | edit_basic | ● | ● | ○ | ○ | ○ | ○ | ○ | ○ |
| 20 | sheet_basic | `sheet.inspect_structure` | edit_basic | ● | ● | ○ | ○ | ○ | ○ | ○ | ○ |
| 21 | sheet_structure | `sheet.preview_delete` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 22 | sheet_structure | `sheet.delete_sheet` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 23 | sheet_structure | `sheet.set_auto_filter` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 24 | sheet_structure | `sheet.get_conditional_formats` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 25 | sheet_structure | `sheet.get_data_validations` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 26 | sheet_structure | `sheet.copy` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 27 | sheet_structure | `sheet.move` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 28 | sheet_structure | `sheet.hide` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 29 | sheet_structure | `sheet.unhide` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 30 | range_basic | `range.read_values` | edit_basic | ● | ● | ○ | ○ | ○ | ○ | ○ | ○ |
| 31 | range_basic | `range.write_values` | edit_basic | ● | ● | ○ | ○ | ○ | ○ | ○ | ○ |
| 32 | range_basic | `range.clear_contents` | edit_basic | ● | ● | ○ | ○ | ○ | ○ | ○ | ○ |
| 33 | range_basic | `range.copy` | edit_basic | ● | ● | ○ | ○ | ○ | ○ | ○ | ○ |
| 34 | range_structure | `range.insert_rows` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 35 | range_structure | `range.delete_rows` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 36 | range_structure | `range.insert_columns` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 37 | range_structure | `range.delete_columns` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 38 | range_structure | `range.sort_data` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 39 | range_structure | `range.merge` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 40 | range_structure | `range.unmerge` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 41 | range_structure | `range.find_replace` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 42 | range_structure | `range.autofit` | edit_structure | ○ | ○ | ○ | ○ | ○ | ○ | ○ | ○ |
| 43 | formula | `formula.set_single` | calc_format | · | ● | · | · | · | · | · | · |
| 44 | formula | `formula.fill_range` | calc_format | · | ● | · | · | · | · | · | · |
| 45 | formula | `formula.get_dependencies` | calc_format | · | ● | · | · | · | · | · | · |
| 46 | formula | `formula.repair` | calc_format | · | ● | · | · | · | · | · | · |
| 47 | format | `format.set_number_format` | calc_format | · | ● | · | · | · | · | · | · |
| 48 | format | `format.set_font` | calc_format | · | ● | · | · | · | · | · | · |
| 49 | format | `format.set_fill` | calc_format | · | ● | · | · | · | · | · | · |
| 50 | format | `format.set_border` | calc_format | · | ● | · | · | · | · | · | · |
| 51 | format | `format.set_alignment` | calc_format | · | ● | · | · | · | · | · | · |
| 52 | format | `format.set_column_width` | calc_format | · | ● | · | · | · | · | · | · |
| 53 | format | `format.set_row_height` | calc_format | · | ● | · | · | · | · | · | · |
| 54 | conditional_format | `format.apply_conditional_rule` | format_rules | · | ● | · | · | · | · | · | · |
| 55 | conditional_format | `format.update_conditional_rule` | format_rules | · | ● | · | · | · | · | · | · |
| 56 | conditional_format | `format.copy_conditional_rules` | format_rules | · | ● | · | · | · | · | · | · |
| 57 | conditional_format | `format.clear_conditional_rules` | format_rules | · | ● | · | · | · | · | · | · |
| 58 | chart | `chart.list_charts` | report | · | · | · | · | ● | · | · | · |
| 59 | chart | `chart.inspect` | report | · | · | · | · | ● | · | · | · |
| 60 | chart | `chart.list_series` | report | · | · | · | · | ● | · | · | · |
| 61 | chart | `chart.export_spec` | report | · | · | · | · | ● | · | · | · |
| 62 | vba | `vba.inspect_project` | automation | · | · | ● | · | · | · | · | · |
| 63 | vba | `vba.get_module_code` | automation | · | · | ● | · | · | · | · | · |
| 64 | vba | `vba.sync_module` | automation | · | · | ● | · | · | · | · | · |
| 65 | vba | `vba.remove_module` | automation | · | · | ● | · | · | · | · | · |
| 66 | vba | `vba.execute` | automation | · | · | ● | · | · | · | · | · |
| 67 | vba | `vba.scan_code` | automation | · | · | ● | · | · | · | · | · |
| 68 | vba | `vba.compile` | automation | · | · | ● | · | · | · | · | · |
| 69 | vba | `vba.import_module` | automation | · | · | ● | · | · | · | · | · |
| 70 | vba | `vba.export_module` | automation | · | · | ● | · | · | · | · | · |
| 71 | recovery | `snapshot.manage` | recovery | · | · | ● | · | · | · | · | · |
| 72 | recovery | `snapshot.get_stats` | recovery | · | · | ● | · | · | · | · | · |
| 73 | recovery | `snapshot.cleanup` | recovery | · | · | ● | · | · | · | · | · |
| 74 | recovery | `rollback.manage` | recovery | · | · | ● | · | · | · | · | · |
| 75 | recovery | `rollback.preview_snapshot` | recovery | · | · | ● | · | · | · | · | · |
| 76 | recovery | `rollback.restore_snapshot` | recovery | · | · | ● | · | · | · | · | · |
| 77 | recovery | `backups.manage` | recovery | · | · | ● | · | · | · | · | · |
| 78 | recovery | `backups.restore` | recovery | · | · | ● | · | · | · | · | · |
| 79 | analysis | `audit.list_operations` | analysis | · | · | · | ● | ● | · | · | · |
| 80 | analysis | `analysis.scan_structure` | analysis | · | · | · | ● | ● | · | · | · |
| 81 | analysis | `analysis.scan_formulas` | analysis | · | · | · | ● | ● | · | · | · |
| 82 | analysis | `analysis.scan_links` | analysis | · | · | · | ● | ● | · | · | · |
| 83 | analysis | `analysis.scan_hidden` | analysis | · | · | · | ● | ● | · | · | · |
| 84 | analysis | `analysis.export_report` | analysis | · | · | · | ● | ● | · | · | · |
| 85 | pq | `pq.list_connections` | data | · | · | · | ● | · | · | · | · |
| 86 | pq | `pq.list_queries` | data | · | · | · | ● | · | · | · | · |
| 87 | pq | `pq.get_code` | data | · | · | · | ● | · | · | · | · |
| 88 | pq | `pq.update_query` | data | · | · | · | ● | · | · | · | · |
| 89 | pq | `pq.refresh` | data | · | · | · | ● | · | · | · | · |
| 90 | table | `table.list_tables` | data | · | · | · | ● | · | · | · | · |
| 91 | table | `table.create` | data | · | · | · | ● | · | · | · | · |
| 92 | table | `table.inspect` | data | · | · | · | ● | · | · | · | · |
| 93 | table | `table.resize` | data | · | · | · | ● | · | · | · | · |
| 94 | table | `table.rename` | data | · | · | · | ● | · | · | · | · |
| 95 | table | `table.set_style` | data | · | · | · | ● | · | · | · | · |
| 96 | table | `table.toggle_total_row` | data | · | · | · | ● | · | · | · | · |
| 97 | table | `table.delete` | data | · | · | · | ● | · | · | · | · |
| 98 | sheet | `sheet.export_csv` | workbook_ops | ○ | ○ | ○ | ○ | ● | ○ | ○ | ○ |
| 99 | package | `package.inspect_manifest` | artifact_export | · | · | · | · | · | ● | ● | ● |
| 100 | package | `package.list_parts` | artifact_export | · | · | · | · | · | ● | ● | ● |
| 101 | package | `package.get_part_xml` | artifact_export | · | · | · | · | · | ● | ● | ● |
| 102 | package | `package.extract_part` | artifact_export | · | · | · | · | · | ● | ● | ● |
| 103 | package | `package.list_media` | artifact_export | · | · | · | · | · | ● | ● | ● |
| 104 | package | `package.list_custom_xml` | artifact_export | · | · | · | · | · | ● | ● | ● |
| 105 | package | `package.detect_features` | artifact_export | · | · | · | · | · | ● | ● | ● |
| 106 | package | `package.export_manifest` | artifact_export | · | · | · | · | · | ● | ● | ● |
| 107 | package_patch | `package.clone_with_patch` | artifact_patch | · | · | · | · | · | · | ● | ● |
| 108 | package_patch | `package.replace_shared_strings` | artifact_patch | · | · | · | · | · | · | ● | ● |
| 109 | package_patch | `package.patch_defined_names` | artifact_patch | · | · | · | · | · | · | ● | ● |
| 110 | package_patch | `package.update_docprops` | artifact_patch | · | · | · | · | · | · | ● | ● |
| 111 | package_patch | `package.merge_template_parts` | artifact_patch | · | · | · | · | · | · | ● | ● |
| 112 | package_patch | `package.remove_external_links` | artifact_patch | · | · | · | · | · | · | ● | ● |
| 113 | package_patch | `package.compare_workbooks` | artifact_patch | · | · | · | · | · | · | ● | ● |
| 114 | batch | `package.batch_extract_parts` | batch_ops | · | · | · | · | · | · | · | ● |
| 115 | batch | `package.batch_transform` | batch_ops | · | · | · | · | · | · | · | ● |
| 116 | batch | `package.batch_compare` | batch_ops | · | · | · | · | · | · | · | ● |

### 4.2 反向查询：Profile → Bundle → 工具数

| Profile | 类别 | Bundle 组合 | 工具数 | 配置文件 |
|---------|------|-----------|------:|---------|
| **basic_edit** | 用户 | foundation(8) + edit_basic(7) | **15** | profiles.yaml |
| **calc_format** | 用户 | foundation(8) + edit_basic(7) + calc_format(11) + format_rules(4) | **30** | profiles.yaml |
| **automation** | 用户 | foundation(8) + automation(9) + recovery(8) | **25** | profiles.yaml |
| **data_workflow** | 用户 | foundation(8) + data(13) + analysis(6) | **27** | profiles.yaml |
| **reporting** | 用户 | foundation(8) + workbook_ops(6) + analysis(6) + report(4) | **24** | profiles.yaml |
| **artifact_extract** | 用户 | foundation(8) + artifact_export(8) | **16** | profiles.yaml |
| **artifact_transform** | 用户 | foundation(8) + artifact_export(8) + artifact_patch(7) | **23** | profiles.yaml |
| **batch_delivery** | 用户 | foundation(8) + artifact_export(8) + artifact_patch(7) + batch_ops(3) | **26** | profiles.yaml |
| **all** | 开发 | 全部 Bundle | **101** | profiles.yaml |
| **minimal** | 开发 | foundation(8) | **8** | profiles.dev.yaml |
| **edit_first** | 开发 | foundation(8) + edit_basic(7) | **15** | profiles.dev.yaml |
| **vba_first** | 开发 | foundation(8) + automation(9) | **17** | profiles.dev.yaml |
| **format_first** | 开发 | foundation(8) + calc_format(11) | **19** | profiles.dev.yaml |
| **recovery_first** | 开发 | foundation(8) + recovery(8) | **16** | profiles.dev.yaml |
| **pq_first** | 开发 | foundation(8) + data(13) | **21** | profiles.dev.yaml |
| **names_first** | 开发 | foundation(8) + names(4) | **12** | profiles.dev.yaml |
| **trae_debug** | 开发 | foundation(8) + edit_basic(7) + calc_format(11) | **26** | profiles.dev.yaml |
| **artifact_first** | 开发 | foundation(8) + artifact_export(8) | **16** | profiles.dev.yaml |
| **package_first** | 开发 | foundation(8) + artifact_export(8) + artifact_patch(7) | **23** | profiles.dev.yaml |
| **batch_first** | 开发 | foundation(8) + artifact_export(8) + batch_ops(3) | **19** | profiles.dev.yaml |
| **chart_first** | 开发 | foundation(8) + report(4) | **12** | profiles.dev.yaml |
| **rules_first** | 开发 | foundation(8) + format_rules(4) | **12** | profiles.dev.yaml |

---

## 5. Bundle 速查

### 5.1 Bundle 总览

| Bundle | 工具数 | 依赖 | 涉及 Domain | 执行模式 | 定位 |
|--------|-------:|------|-------------|----------|------|
| **foundation** | 8 | 无 | server, workbook | runtime | 必选基础：服务状态 + 工作簿 I/O |
| **names** | 4 | foundation | names | runtime | 命名范围管理 |
| **edit_basic** | 7 | foundation | sheet_basic, range_basic | runtime | 高频编辑：创建/读写/复制 |
| **edit_structure** | 18 | foundation | sheet_structure, range_structure | runtime | 结构操作：插删/合并/筛选/隐藏/查找 |
| **calc_format** | 11 | foundation | formula, format | runtime | 公式设置 + 格式调整 |
| **format_rules** | 4 | foundation | conditional_format | file_package | 条件格式规则（无需 Excel） |
| **automation** | 9 | foundation | vba | runtime | VBA 检查/同步/执行/导入导出 |
| **recovery** | 8 | foundation | recovery | runtime | 快照/回滚/备份 |
| **analysis** | 6 | foundation | analysis | runtime | 结构扫描/公式审计/报告 |
| **data** | 13 | foundation | pq, table | runtime | PQ 查询 + Table 管理 |
| **workbook_ops** | 6 | foundation | workbook_ops, sheet | runtime | 另存/刷新/计算/导出 |
| **report** | 4 | foundation | chart | file_package | 图表解析（无需 Excel） |
| **artifact_export** | 8 | foundation | package | file_package | OOXML 包提取（无需 Excel） |
| **artifact_patch** | 7 | foundation, artifact_export | package_patch | file_package | OOXML 补丁/克隆/比较 |
| **batch_ops** | 3 | foundation, artifact_export | batch | batch | 批量文件处理 |

### 5.2 各 Bundle 工具明细

#### foundation（8 工具）— 所有场景必选

| # | 工具 | 说明 | 执行模式 |
|--:|------|------|----------|
| 1 | `server.health` | 健康检查 | runtime |
| 2 | `server.get_status` | 服务状态 | runtime |
| 3 | `workbook.open_file` | 打开工作簿 | runtime |
| 4 | `workbook.create_file` | 创建工作簿 | runtime |
| 5 | `workbook.save_file` | 保存工作簿 | runtime |
| 6 | `workbook.close_file` | 关闭工作簿 | runtime |
| 7 | `workbook.list_open` | 列出已打开 | runtime |
| 8 | `workbook.inspect` | 检查元数据 | runtime |

#### names（4 工具）

| # | 工具 | 说明 |
|--:|------|------|
| 1 | `names.inspect` | 检查命名范围 |
| 2 | `names.manage` | 管理命名范围 |
| 3 | `names.create` | 创建命名范围 |
| 4 | `names.delete` | 删除命名范围 |

#### edit_basic（7 工具）

| # | 工具 | Domain | 说明 |
|--:|------|--------|------|
| 1 | `sheet.create_sheet` | sheet_basic | 创建工作表 |
| 2 | `sheet.rename_sheet` | sheet_basic | 重命名工作表 |
| 3 | `sheet.inspect_structure` | sheet_basic | 查看结构 |
| 4 | `range.read_values` | range_basic | 读取值 |
| 5 | `range.write_values` | range_basic | 写入值 |
| 6 | `range.clear_contents` | range_basic | 清空内容 |
| 7 | `range.copy` | range_basic | 复制区域 |

#### edit_structure（18 工具）

| # | 工具 | Domain | 说明 |
|--:|------|--------|------|
| 1 | `sheet.preview_delete` | sheet_structure | 预览删除影响 |
| 2 | `sheet.delete_sheet` | sheet_structure | 删除工作表 |
| 3 | `sheet.set_auto_filter` | sheet_structure | 设置自动筛选 |
| 4 | `sheet.get_conditional_formats` | sheet_structure | 获取条件格式 |
| 5 | `sheet.get_data_validations` | sheet_structure | 获取数据验证 |
| 6 | `sheet.copy` | sheet_structure | 复制工作表 |
| 7 | `sheet.move` | sheet_structure | 移动工作表 |
| 8 | `sheet.hide` | sheet_structure | 隐藏工作表 |
| 9 | `sheet.unhide` | sheet_structure | 取消隐藏 |
| 10 | `range.insert_rows` | range_structure | 插入行 |
| 11 | `range.delete_rows` | range_structure | 删除行 |
| 12 | `range.insert_columns` | range_structure | 插入列 |
| 13 | `range.delete_columns` | range_structure | 删除列 |
| 14 | `range.sort_data` | range_structure | 排序 |
| 15 | `range.merge` | range_structure | 合并单元格 |
| 16 | `range.unmerge` | range_structure | 拆分单元格 |
| 17 | `range.find_replace` | range_structure | 查找替换 |
| 18 | `range.autofit` | range_structure | 自动适应宽高 |

#### calc_format（11 工具）

| # | 工具 | Domain | 说明 |
|--:|------|--------|------|
| 1 | `formula.set_single` | formula | 设置单个公式 |
| 2 | `formula.fill_range` | formula | 填充区域公式 |
| 3 | `formula.get_dependencies` | formula | 获取依赖 |
| 4 | `formula.repair` | formula | 修复公式 |
| 5 | `format.set_number_format` | format | 数字格式 |
| 6 | `format.set_font` | format | 字体 |
| 7 | `format.set_fill` | format | 填充色 |
| 8 | `format.set_border` | format | 边框 |
| 9 | `format.set_alignment` | format | 对齐 |
| 10 | `format.set_column_width` | format | 列宽 |
| 11 | `format.set_row_height` | format | 行高 |

#### format_rules（4 工具）— V2.5 新增

| # | 工具 | Domain | 说明 | 执行模式 |
|--:|------|--------|------|----------|
| 1 | `format.apply_conditional_rule` | conditional_format | 应用条件格式规则 | file_package |
| 2 | `format.update_conditional_rule` | conditional_format | 更新条件格式规则 | file_package |
| 3 | `format.copy_conditional_rules` | conditional_format | 复制条件格式规则 | file_package |
| 4 | `format.clear_conditional_rules` | conditional_format | 清除条件格式规则 | file_package |

#### automation（9 工具）

| # | 工具 | 说明 |
|--:|------|------|
| 1 | `vba.inspect_project` | 检查 VBA 项目 |
| 2 | `vba.get_module_code` | 获取模块代码 |
| 3 | `vba.sync_module` | 同步模块 |
| 4 | `vba.remove_module` | 移除模块 |
| 5 | `vba.execute` | 执行宏 |
| 6 | `vba.scan_code` | 扫描代码 |
| 7 | `vba.compile` | 编译项目 |
| 8 | `vba.import_module` | 导入模块（V2.5 新增）|
| 9 | `vba.export_module` | 导出模块（V2.5 新增）|

#### recovery（8 工具）

| # | 工具 | 说明 |
|--:|------|------|
| 1 | `snapshot.manage` | 管理快照 |
| 2 | `snapshot.get_stats` | 快照统计 |
| 3 | `snapshot.cleanup` | 清理快照 |
| 4 | `rollback.manage` | 管理回滚 |
| 5 | `rollback.preview_snapshot` | 预览快照 |
| 6 | `rollback.restore_snapshot` | 恢复快照 |
| 7 | `backups.manage` | 管理备份 |
| 8 | `backups.restore` | 恢复备份 |

#### analysis（6 工具）

| # | 工具 | 说明 |
|--:|------|------|
| 1 | `audit.list_operations` | 审计操作记录 |
| 2 | `analysis.scan_structure` | 扫描工作簿结构 |
| 3 | `analysis.scan_formulas` | 扫描公式分布与错误 |
| 4 | `analysis.scan_links` | 扫描外部链接 |
| 5 | `analysis.scan_hidden` | 扫描隐藏元素 |
| 6 | `analysis.export_report` | 生成分析报告 |

#### data（13 工具）

| # | 工具 | Domain | 说明 |
|--:|------|--------|------|
| 1 | `pq.list_connections` | pq | 列出数据连接 |
| 2 | `pq.list_queries` | pq | 列出查询 |
| 3 | `pq.get_code` | pq | 获取查询代码 |
| 4 | `pq.update_query` | pq | 更新查询 |
| 5 | `pq.refresh` | pq | 刷新查询 |
| 6 | `table.list_tables` | table | 列出所有表格 |
| 7 | `table.create` | table | 创建表格 |
| 8 | `table.inspect` | table | 检查表格结构 |
| 9 | `table.resize` | table | 调整表格范围 |
| 10 | `table.rename` | table | 重命名表格 |
| 11 | `table.set_style` | table | 设置表格样式 |
| 12 | `table.toggle_total_row` | table | 开关总计行 |
| 13 | `table.delete` | table | 删除表格 |

#### workbook_ops（6 工具）

| # | 工具 | Domain | 说明 |
|--:|------|--------|------|
| 1 | `workbook.save_as` | workbook_ops | 另存为 |
| 2 | `workbook.refresh_all` | workbook_ops | 刷新所有连接 |
| 3 | `workbook.calculate` | workbook_ops | 重新计算 |
| 4 | `workbook.list_links` | workbook_ops | 列出外部链接 |
| 5 | `workbook.export_pdf` | workbook_ops | 导出 PDF |
| 6 | `sheet.export_csv` | workbook_ops | 导出 CSV |

#### report（4 工具）— V2.5 图表独立

| # | 工具 | Domain | 说明 | 执行模式 |
|--:|------|--------|------|----------|
| 1 | `chart.list_charts` | chart | 列出所有图表 | file_package |
| 2 | `chart.inspect` | chart | 检查图表详情 | file_package |
| 3 | `chart.list_series` | chart | 列出图表系列 | file_package |
| 4 | `chart.export_spec` | chart | 导出图表规格 | file_package |

#### artifact_export（8 工具）— V2.5 新增

| # | 工具 | Domain | 说明 | 执行模式 |
|--:|------|--------|------|----------|
| 1 | `package.inspect_manifest` | package | 检查包清单 | file_package |
| 2 | `package.list_parts` | package | 列出所有部件 | file_package |
| 3 | `package.get_part_xml` | package | 获取部件 XML | file_package |
| 4 | `package.extract_part` | package | 提取部件文件 | file_package |
| 5 | `package.list_media` | package | 列出媒体资源 | file_package |
| 6 | `package.list_custom_xml` | package | 列出自定义 XML | file_package |
| 7 | `package.detect_features` | package | 检测文件特性 | file_package |
| 8 | `package.export_manifest` | package | 导出清单报告 | file_package |

#### artifact_patch（7 工具）— V2.5 新增

| # | 工具 | Domain | 说明 | 执行模式 |
|--:|------|--------|------|----------|
| 1 | `package.clone_with_patch` | package_patch | 克隆并打补丁 | file_package |
| 2 | `package.replace_shared_strings` | package_patch | 替换共享字符串 | file_package |
| 3 | `package.patch_defined_names` | package_patch | 修补命名范围 | file_package |
| 4 | `package.update_docprops` | package_patch | 更新文档属性 | file_package |
| 5 | `package.merge_template_parts` | package_patch | 合并模板部件 | file_package |
| 6 | `package.remove_external_links` | package_patch | 删除外部链接 | file_package |
| 7 | `package.compare_workbooks` | package_patch | 比较两个工作簿 | file_package |

#### batch_ops（3 工具）— V2.5 新增

| # | 工具 | Domain | 说明 | 执行模式 |
|--:|------|--------|------|----------|
| 1 | `package.batch_extract_parts` | batch | 批量提取部件 | batch |
| 2 | `package.batch_transform` | batch | 批量变换文件 | batch |
| 3 | `package.batch_compare` | batch | 批量比较文件 | batch |

---

## 6. Profile 速查

### 6.1 用户 Profile

#### basic_edit — 基础编辑（15 工具）

**场景**：打开 → 查看结构 → 读写数据 → 基础 sheet 管理 → 保存

```
foundation ──────► 8 工具（server + workbook I/O）
edit_basic ──────► 7 工具（sheet 创建/重命名 + range 读写/复制）
                  ─────
                  15 工具   预算 20
```

#### calc_format — 公式与格式（30 工具）

**场景**：基础编辑 + 公式设置 + 格式调整 + 条件格式

```
foundation ──────► 8 工具
edit_basic ──────► 7 工具
calc_format ─────► 11 工具（formula 4 + format 7）
format_rules ────► 4 工具（conditional_format 4）
                  ─────
                  30 工具   预算 30
```

#### automation — VBA 与恢复（25 工具）

**场景**：VBA 检查/同步/执行/导入导出 + 快照/备份/回滚

```
foundation ──────► 8 工具
automation ──────► 9 工具（vba 全部，含 import/export）
recovery ────────► 8 工具（snapshot + rollback + backups）
                  ─────
                  25 工具   预算 30
```

#### data_workflow — 数据工作流（27 工具）

**场景**：PQ 管理 → 数据刷新 → Table 操作 → 分析审计

```
foundation ──────► 8 工具
data ────────────► 13 工具（pq 5 + table 8）
analysis ────────► 6 工具
                  ─────
                  27 工具   预算 30
```

#### reporting — 报表输出（24 工具）

**场景**：导出 PDF/CSV + 结构分析 + 图表检查

```
foundation ──────► 8 工具
workbook_ops ────► 6 工具（save_as + export + 刷新/计算）
analysis ────────► 6 工具
report ──────────► 4 工具（chart 图表解析，无需 Excel）
                  ─────
                  24 工具   预算 30
```

#### artifact_extract — OOXML 提取（16 工具）— V2.5 新增

**场景**：直接解析 Excel 文件 ZIP 结构，无需启动 Excel COM

```
foundation ──────► 8 工具
artifact_export ─► 8 工具（package 包提取）
                  ─────
                  16 工具   预算 20
```

#### artifact_transform — OOXML 变换（23 工具）— V2.5 新增

**场景**：文件包补丁 / 字符串替换 / 命名范围 / 文档属性 / 模板合并 / 外链删除 / 工作簿比较

```
foundation ──────► 8 工具
artifact_export ─► 8 工具
artifact_patch ──► 7 工具
                  ─────
                  23 工具   预算 30
```

#### batch_delivery — 批量交付（26 工具）— V2.5 新增

**场景**：批量提取 / 批量变换 / 批量比较

```
foundation ──────► 8 工具
artifact_export ─► 8 工具
artifact_patch ──► 7 工具
batch_ops ───────► 3 工具
                  ─────
                  26 工具   预算 30
```

### 6.2 开发 Profile（需 `--profiles-file profiles.dev.yaml`）

| Profile | 组合 | 工具数 | 用途 |
|---------|------|-------:|------|
| minimal | foundation | 8 | 最小连通性验证 |
| edit_first | foundation + edit_basic | 15 | 编辑专题 |
| vba_first | foundation + automation | 17 | VBA 专题 |
| format_first | foundation + calc_format | 19 | 格式专题 |
| recovery_first | foundation + recovery | 16 | 恢复专题 |
| pq_first | foundation + data | 21 | PQ/数据专题 |
| names_first | foundation + names | 12 | 命名管理专题 |
| trae_debug | foundation + edit_basic + calc_format | 26 | Trae 调试 |
| artifact_first | foundation + artifact_export | 16 | OOXML 提取专题 |
| package_first | foundation + artifact_export + artifact_patch | 23 | Package 变换专题 |
| batch_first | foundation + artifact_export + batch_ops | 19 | 批量处理专题 |
| chart_first | foundation + report | 12 | 图表专题 |
| rules_first | foundation + format_rules | 12 | 条件格式专题 |

### 6.3 全量 Profile

| Profile | 工具数 | 说明 |
|---------|-------:|------|
| all | 101 | 全部 Bundle，仅用于 CLI / 回归测试，不推荐在受限客户端中使用 |

---

## 7. 常用查询场景

### "我想做某件事，用哪个 Profile？"

| 我想… | Profile | 工具数 | 是否需要 Excel |
|-------|---------|-------:|---------------|
| 打开文件读写数据 | `basic_edit` | 15 | 是 |
| 设公式调格式 | `calc_format` | 30 | 是 |
| 跑 VBA 宏 | `automation` | 25 | 是 |
| 管 PQ 查询和 Table | `data_workflow` | 27 | 是 |
| 导出 PDF/CSV + 做分析 | `reporting` | 24 | 是 |
| 解析 Excel 文件结构（无需 Excel） | `artifact_extract` | 16 | **否** |
| 批量处理多个 Excel 文件 | `batch_delivery` | 26 | **否** |
| 做图表分析 | `reporting` | 24 | 部分（图表解析不需要） |
| 只验证连通性 | `minimal`（dev） | 8 | 是 |
| 全量回归测试 | `all` | 101 | 混合 |

### "某个工具属于哪个 Bundle？在哪些 Profile 中默认可用？"

| 工具 | Bundle | 默认可用的 Profile |
|------|--------|-----------------|
| `workbook.open_file` | foundation | **所有** |
| `range.read_values` | edit_basic | basic_edit, calc_format |
| `range.insert_rows` | edit_structure | 无（需 `--enable-bundle edit_structure`） |
| `formula.set_single` | calc_format | calc_format |
| `format.apply_conditional_rule` | format_rules | calc_format |
| `chart.list_charts` | report | reporting |
| `vba.execute` | automation | automation |
| `snapshot.manage` | recovery | automation |
| `pq.list_queries` | data | data_workflow |
| `table.create` | data | data_workflow |
| `analysis.scan_structure` | analysis | data_workflow, reporting |
| `workbook.export_pdf` | workbook_ops | reporting |
| `package.inspect_manifest` | artifact_export | artifact_extract, artifact_transform, batch_delivery |
| `package.clone_with_patch` | artifact_patch | artifact_transform, batch_delivery |
| `package.batch_extract_parts` | batch_ops | batch_delivery |
| `names.inspect` | names | 无（需 `--enable-bundle names`） |

### "我想给当前 Profile 加更多能力"

```bash
# 在 basic_edit 基础上加命名管理
--profile basic_edit --enable-bundle names
# 结果：15 + 4 = 19 工具

# 在 basic_edit 基础上加结构编辑
--profile basic_edit --enable-bundle edit_structure
# 结果：15 + 18 = 33 工具

# 在 calc_format 基础上加图表（无需 Excel）
--profile calc_format --enable-bundle report
# 结果：30 + 4 = 34 工具

# 在 artifact_extract 基础上加补丁能力
--profile artifact_extract --enable-bundle artifact_patch
# 结果：16 + 7 = 23 工具
```

### 客户端推荐

| 客户端 | 推荐 Profile | 避免 | 说明 |
|--------|-------------|------|------|
| Trae AI | `basic_edit` / `*_first` / `artifact_first` | `all` | UI 截断约 39 工具 |
| Workbuddy | 任意用户 Profile | — | 承载能力较好 |
| CLI / run_mcp | `all` / 任意 | — | 无限制 |
| **AI 编程（无需 Excel）** | `artifact_first` / `package_first` | — | file_package 模式 |

---

## 8. V2.5 新增功能速览

### 8.1 新增 Domain（7 个）

| Domain | 工具数 | 执行模式 | 说明 |
|--------|-------:|----------|------|
| **conditional_format** | 4 | file_package | 条件格式规则读写 |
| **chart** | 4 | file_package | 图表解析（从 report 独立）|
| **package** | 8 | file_package | OOXML 包提取 |
| **package_patch** | 7 | file_package | OOXML 补丁与克隆 |
| **batch** | 3 | batch | 批量文件处理 |

### 8.2 新增 Bundle（5 个）

| Bundle | 工具数 | 依赖 | 说明 |
|--------|-------:|------|------|
| **format_rules** | 4 | foundation | 条件格式规则 |
| **artifact_export** | 8 | foundation | OOXML 包提取 |
| **artifact_patch** | 7 | foundation, artifact_export | OOXML 补丁 |
| **batch_ops** | 3 | foundation, artifact_export | 批量处理 |
| **report** | 4 | foundation | 图表（独立自成一Bundle）|

### 8.3 新增 Profile（8 个）

| Profile | 类别 | 工具数 | 说明 |
|---------|------|-------:|------|
| **artifact_extract** | 用户 | 16 | OOXML 提取 |
| **artifact_transform** | 用户 | 23 | OOXML 变换 |
| **batch_delivery** | 用户 | 26 | 批量交付 |
| **artifact_first** | 开发 | 16 | OOXML 提取专题 |
| **package_first** | 开发 | 23 | Package 变换专题 |
| **batch_first** | 开发 | 19 | 批量处理专题 |
| **chart_first** | 开发 | 12 | 图表专题 |
| **rules_first** | 开发 | 12 | 条件格式专题 |

### 8.4 关键特性

1. **无需 Excel COM**：新增的 package/chart/conditional_format/batch 工具全部基于 file_package 或 batch 模式，直接解析 Excel 文件的 ZIP 结构
2. **批量处理**：batch_ops 支持批量提取、变换、比较多个 Excel 文件
3. **工作簿比较**：package.compare_workbooks 可以比较两个工作簿的差异

---

## 9. 预算规则速记

| 区间 | 含义 | 处理 |
|------|------|------|
| **≤ 30** | 理想 | 正常 |
| **31 ~ 39** | 可接受 | 需设计说明 |
| **≥ 40** | 受限客户端高风险 | 不作为默认场景 Profile |

---

## 10. 执行模式快速参考

| 工具类型 | 是否需要 Excel | 代表工具 |
|----------|--------------|---------|
| runtime | 必须 | `workbook.open_file`, `vba.execute`, `range.read_values` |
| file_package | 不需要 | `package.inspect_manifest`, `chart.list_charts`, `format.apply_conditional_rule` |
| batch | 不需要 | `package.batch_extract_parts`, `package.batch_transform`, `package.batch_compare` |

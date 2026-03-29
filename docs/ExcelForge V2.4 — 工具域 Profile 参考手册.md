# ExcelForge V2.4 — 工具域 Profile 参考手册

**产品版本**：v2.4.0 · **工具总数**：91 · **Bundle 数**：12 · **Profile 数**：13

---

## 1. 核心概念

| 概念 | 定义 | 粒度 | 数量 |
|------|------|------|-----:|
| **Domain（能力域）** | 工具的逻辑分类命名空间。一个工具属于且仅属于一个 Domain。Domain 本身不参与装配，仅用于语义归类。 | 最细 | 15 |
| **Tool（工具）** | 最小可调用能力单元，由 Domain 命名空间 + 动作名组成，如 `workbook.open_file`。 | — | 91 |
| **Bundle（能力包）** | 一组语义相关工具的可复用集合。Bundle 是装配的基本单位，Profile 通过组合 Bundle 来确定最终工具集。一个 Bundle 可跨 Domain 取工具。 | 中间 | 12 |
| **Profile（场景配置）** | 面向用户任务场景的 Bundle 组合。Profile 回答"用户要完成什么任务"，决定运行时实际注册哪些工具。 | 最粗 | 13 |

**一句话关系**：Domain 给工具分类 → Bundle 把工具打包 → Profile 把包组合成场景。

抽象层级图
```
Profile（场景入口）
└─ Bundle（装配单元）
   └─ Tool（具体能力）
      └─ Domain（语义归属）
```      
注意：严格来说，Domain 不是 Tool 的“下一级执行层”，
而是 Tool 的“语义标签 / 分类归属”。
所以上图是便于理解的记忆图。



---

## 2. 层级架构图

```
Profile 层（场景交付）
┌────────────────────────────────────────────────────────────────────┐
│  basic_edit    calc_format    automation    data_workflow    ...   │
│   15 tools      26 tools      23 tools       30 tools             │
└───────┬────────────┬─────────────┬──────────────┬─────────────────┘
        │            │             │              │
        ▼            ▼             ▼              ▼
Bundle 层（能力复用）
┌────────────────────────────────────────────────────────────────────┐
│  foundation  edit_basic  calc_format  automation  recovery  data  │
│   8 tools     7 tools    11 tools     7 tools    8 tools  16 tools│
│              names    edit_structure  workbook_ops   analysis     │
│             4 tools    18 tools       6 tools       6 tools      │
└───────┬────────┬──────────┬────────────┬────────────┬────────────┘
        │        │          │            │            │
        ▼        ▼          ▼            ▼            ▼
Tool 层（最小能力单元）
┌────────────────────────────────────────────────────────────────────┐
│  server.health  workbook.open_file  range.read_values  vba.execute│
│  table.create   pq.list_queries     format.set_font    ...  (×91)│
└───────┬────────────┬──────────────────┬───────────────────────────┘
        │            │                  │
        ▼            ▼                  ▼
Domain 层（逻辑分类，不参与装配）
┌────────────────────────────────────────────────────────────────────┐
│  server  workbook  names  sheet  range  formula  format  analysis │
│  vba     recovery  pq     table  chart  pivot    model            │
└────────────────────────────────────────────────────────────────────┘
```

### 关键理解

```
                   Domain 是分类标签
                        │
          ┌─────────────┼─────────────┐
          ▼             ▼             ▼
       sheet域        sheet域        sheet域
    create_sheet    copy_sheet    delete_sheet
          │             │             │
          ▼             ▼             ▼
     edit_basic    edit_structure  edit_structure    ← 同一 Domain 的工具
      (Bundle)       (Bundle)       (Bundle)          可分属不同 Bundle
          │             │             │
          └──────┬──────┘             │
                 ▼                    ▼
            basic_edit          需 --enable-bundle    ← Bundle 再组入
             (Profile)           edit_structure         不同 Profile
```

> **Domain ≠ Bundle**。Domain 只管命名，Bundle 管装配。同一个 Domain（如 `sheet`）的工具可以被拆进不同 Bundle（如 `edit_basic` 和 `edit_structure`），按使用频率分层。

---

## 3. 全量对照矩阵

### 3.1 正向查询：Domain → Tool → Bundle → Profile 可用性

图例：● = 默认包含　○ = 可通过 `--enable-bundle` 加入　· = 不在该 Profile 任何可选路径中

| # | Domain | Tool | Bundle | basic_ edit | calc_ format | auto- mation | data_ work | report- ing |
|--:|--------|------|--------|:---:|:---:|:---:|:---:|:---:|
| 1 | server | `server.health` | foundation | ● | ● | ● | ● | ● |
| 2 | server | `server.get_status` | foundation | ● | ● | ● | ● | ● |
| 3 | workbook | `workbook.open_file` | foundation | ● | ● | ● | ● | ● |
| 4 | workbook | `workbook.create_file` | foundation | ● | ● | ● | ● | ● |
| 5 | workbook | `workbook.save_file` | foundation | ● | ● | ● | ● | ● |
| 6 | workbook | `workbook.close_file` | foundation | ● | ● | ● | ● | ● |
| 7 | workbook | `workbook.list_open` | foundation | ● | ● | ● | ● | ● |
| 8 | workbook | `workbook.inspect` | foundation | ● | ● | ● | ● | ● |
| 9 | workbook | `workbook.save_as` | workbook_ops | ○ | ○ | ○ | ○ | ● |
| 10 | workbook | `workbook.refresh_all` | workbook_ops | ○ | ○ | ○ | ○ | ● |
| 11 | workbook | `workbook.calculate` | workbook_ops | ○ | ○ | ○ | ○ | ● |
| 12 | workbook | `workbook.list_links` | workbook_ops | ○ | ○ | ○ | ○ | ● |
| 13 | workbook | `workbook.export_pdf` | workbook_ops | ○ | ○ | ○ | ○ | ● |
| 14 | names | `names.inspect` | names | ○ | ○ | ○ | ○ | ○ |
| 15 | names | `names.manage` | names | ○ | ○ | ○ | ○ | ○ |
| 16 | names | `names.create` | names | ○ | ○ | ○ | ○ | ○ |
| 17 | names | `names.delete` | names | ○ | ○ | ○ | ○ | ○ |
| 18 | sheet | `sheet.create_sheet` | edit_basic | ● | ● | ○ | ○ | ○ |
| 19 | sheet | `sheet.rename_sheet` | edit_basic | ● | ● | ○ | ○ | ○ |
| 20 | sheet | `sheet.inspect_structure` | edit_basic | ● | ● | ○ | ○ | ○ |
| 21 | sheet | `sheet.export_csv` | workbook_ops | ○ | ○ | ○ | ○ | ● |
| 22 | sheet | `sheet.preview_delete` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 23 | sheet | `sheet.delete_sheet` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 24 | sheet | `sheet.set_auto_filter` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 25 | sheet | `sheet.get_conditional_formats` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 26 | sheet | `sheet.get_data_validations` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 27 | sheet | `sheet.copy` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 28 | sheet | `sheet.move` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 29 | sheet | `sheet.hide` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 30 | sheet | `sheet.unhide` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 31 | range | `range.read_values` | edit_basic | ● | ● | ○ | ○ | ○ |
| 32 | range | `range.write_values` | edit_basic | ● | ● | ○ | ○ | ○ |
| 33 | range | `range.clear_contents` | edit_basic | ● | ● | ○ | ○ | ○ |
| 34 | range | `range.copy` | edit_basic | ● | ● | ○ | ○ | ○ |
| 35 | range | `range.insert_rows` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 36 | range | `range.delete_rows` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 37 | range | `range.insert_columns` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 38 | range | `range.delete_columns` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 39 | range | `range.sort_data` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 40 | range | `range.merge` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 41 | range | `range.unmerge` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 42 | range | `range.find_replace` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 43 | range | `range.autofit` | edit_structure | ○ | ○ | ○ | ○ | ○ |
| 44 | formula | `formula.set_single` | calc_format | · | ● | · | · | · |
| 45 | formula | `formula.fill_range` | calc_format | · | ● | · | · | · |
| 46 | formula | `formula.get_dependencies` | calc_format | · | ● | · | · | · |
| 47 | formula | `formula.repair` | calc_format | · | ● | · | · | · |
| 48 | format | `format.set_number_format` | calc_format | · | ● | · | · | · |
| 49 | format | `format.set_font` | calc_format | · | ● | · | · | · |
| 50 | format | `format.set_fill` | calc_format | · | ● | · | · | · |
| 51 | format | `format.set_border` | calc_format | · | ● | · | · | · |
| 52 | format | `format.set_alignment` | calc_format | · | ● | · | · | · |
| 53 | format | `format.set_column_width` | calc_format | · | ● | · | · | · |
| 54 | format | `format.set_row_height` | calc_format | · | ● | · | · | · |
| 55 | vba | `vba.inspect_project` | automation | · | · | ● | · | · |
| 56 | vba | `vba.get_module_code` | automation | · | · | ● | · | · |
| 57 | vba | `vba.sync_module` | automation | · | · | ● | · | · |
| 58 | vba | `vba.execute` | automation | · | · | ● | · | · |
| 59 | vba | `vba.scan_code` | automation | · | · | ● | · | · |
| 60 | vba | `vba.compile` | automation | · | · | ● | · | · |
| 61 | vba | `vba.remove_module` | automation | · | · | ● | · | · |
| 62 | recovery | `snapshot.manage` | recovery | · | · | ● | · | · |
| 63 | recovery | `snapshot.get_stats` | recovery | · | · | ● | · | · |
| 64 | recovery | `snapshot.cleanup` | recovery | · | · | ● | · | · |
| 65 | recovery | `rollback.manage` | recovery | · | · | ● | · | · |
| 66 | recovery | `rollback.preview_snapshot` | recovery | · | · | ● | · | · |
| 67 | recovery | `rollback.restore_snapshot` | recovery | · | · | ● | · | · |
| 68 | recovery | `backups.manage` | recovery | · | · | ● | · | · |
| 69 | recovery | `backups.restore` | recovery | · | · | ● | · | · |
| 70 | analysis | `audit.list_operations` | analysis | · | · | · | ● | ● |
| 71 | analysis | `analysis.scan_structure` | analysis | · | · | · | ● | ● |
| 72 | analysis | `analysis.scan_formulas` | analysis | · | · | · | ● | ● |
| 73 | analysis | `analysis.scan_links` | analysis | · | · | · | ● | ● |
| 74 | analysis | `analysis.scan_hidden` | analysis | · | · | · | ● | ● |
| 75 | analysis | `analysis.export_report` | analysis | · | · | · | ● | ● |
| 76 | pq | `pq.list_connections` | data | · | · | · | ● | · |
| 77 | pq | `pq.list_queries` | data | · | · | · | ● | · |
| 78 | pq | `pq.get_code` | data | · | · | · | ● | · |
| 79 | pq | `pq.update_query` | data | · | · | · | ● | · |
| 80 | pq | `pq.refresh` | data | · | · | · | ● | · |
| 81 | pq | `pq.inspect_query` | data | · | · | · | ● | · |
| 82 | pq | `pq.refresh_all` | data | · | · | · | ● | · |
| 83 | pq | `pq.get_refresh_status` | data | · | · | · | ● | · |
| 84 | table | `table.list_tables` | data | · | · | · | ● | · |
| 85 | table | `table.create` | data | · | · | · | ● | · |
| 86 | table | `table.inspect` | data | · | · | · | ● | · |
| 87 | table | `table.resize` | data | · | · | · | ● | · |
| 88 | table | `table.rename` | data | · | · | · | ● | · |
| 89 | table | `table.set_style` | data | · | · | · | ● | · |
| 90 | table | `table.toggle_total_row` | data | · | · | · | ● | · |
| 91 | table | `table.delete` | data | · | · | · | ● | · |

### 3.2 反向查询：Profile → Bundle → 工具数

| Profile | 类别 | Bundle 组合 | 工具数 | 预算 | 配置文件 |
|---------|------|-----------|------:|----:|---------|
| **basic_edit** | 用户 | foundation(8) + edit_basic(7) | **15** | 20 | profiles.yaml |
| **calc_format** | 用户 | foundation(8) + edit_basic(7) + calc_format(11) | **26** | 30 | profiles.yaml |
| **automation** | 用户 | foundation(8) + automation(7) + recovery(8) | **23** | 25 | profiles.yaml |
| **data_workflow** | 用户 | foundation(8) + data(16) + analysis(6) | **30** | 35 | profiles.yaml |
| **reporting** | 用户 | foundation(8) + workbook_ops(6) + analysis(6) | **20** | 25 | profiles.yaml |
| **minimal** | 开发 | foundation(8) | **8** | 10 | profiles.dev.yaml |
| **vba_first** | 开发 | foundation(8) + automation(7) | **15** | 16 | profiles.dev.yaml |
| **format_first** | 开发 | foundation(8) + calc_format(11) | **19** | 20 | profiles.dev.yaml |
| **recovery_first** | 开发 | foundation(8) + recovery(8) | **16** | 18 | profiles.dev.yaml |
| **pq_first** | 开发 | foundation(8) + data(16) | **24** | 25 | profiles.dev.yaml |
| **edit_first** | 开发 | foundation(8) + edit_basic(7) | **15** | 16 | profiles.dev.yaml |
| **trae_debug** | 开发 | foundation(8) + edit_basic(7) | **15** | 16 | profiles.dev.yaml |
| **all** | 开发 | 全部 Bundle | **91** | 无限制 | profiles.yaml |

---

## 4. Bundle 速查

### 4.1 Bundle 总览

| Bundle | 工具数 | 依赖 | 涉及 Domain | 定位 |
|--------|-------:|------|-------------|------|
| **foundation** | 8 | 无 | server, workbook | 必选基础：服务状态 + 工作簿 I/O |
| **names** | 4 | foundation | names | 命名范围管理 |
| **edit_basic** | 7 | foundation | sheet, range | 高频编辑：创建/读写/复制 |
| **edit_structure** | 18 | foundation | sheet, range | 结构操作：插删/合并/筛选/隐藏/查找 |
| **calc_format** | 11 | foundation | formula, format | 公式设置 + 格式调整 |
| **automation** | 7 | foundation | vba | VBA 检查/同步/执行 |
| **recovery** | 8 | foundation | recovery | 快照/回滚/备份 |
| **analysis** | 6 | foundation | analysis | 结构扫描/公式审计/报告 |
| **data** | 16 | foundation | pq, table | PQ 查询 + Table 管理 |
| **workbook_ops** | 6 | foundation | workbook, sheet | 另存/刷新/计算/导出 |
| **report** | 0 | foundation | chart, pivot, model | v2.5+ 扩充 |
| **artifact_export** | 0 | foundation | — | v2.5+ 扩充 |

### 4.2 各 Bundle 工具明细

#### foundation（8 工具）— 所有场景必选

| # | 工具 | 说明 |
|--:|------|------|
| 1 | `server.health` | 健康检查 |
| 2 | `server.get_status` | 服务状态 |
| 3 | `workbook.open_file` | 打开工作簿 |
| 4 | `workbook.create_file` | 创建工作簿 |
| 5 | `workbook.save_file` | 保存工作簿 |
| 6 | `workbook.close_file` | 关闭工作簿 |
| 7 | `workbook.list_open` | 列出已打开 |
| 8 | `workbook.inspect` | 检查元数据 |

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
| 1 | `sheet.create_sheet` | sheet | 创建工作表 |
| 2 | `sheet.rename_sheet` | sheet | 重命名工作表 |
| 3 | `sheet.inspect_structure` | sheet | 查看结构 |
| 4 | `range.read_values` | range | 读取值 |
| 5 | `range.write_values` | range | 写入值 |
| 6 | `range.clear_contents` | range | 清空内容 |
| 7 | `range.copy` | range | 复制区域 |

#### edit_structure（18 工具）

| # | 工具 | Domain | 说明 |
|--:|------|--------|------|
| 1 | `sheet.preview_delete` | sheet | 预览删除影响 |
| 2 | `sheet.delete_sheet` | sheet | 删除工作表 |
| 3 | `sheet.set_auto_filter` | sheet | 设置自动筛选 |
| 4 | `sheet.get_conditional_formats` | sheet | 获取条件格式 |
| 5 | `sheet.get_data_validations` | sheet | 获取数据验证 |
| 6 | `sheet.copy` | sheet | 复制工作表 |
| 7 | `sheet.move` | sheet | 移动工作表 |
| 8 | `sheet.hide` | sheet | 隐藏工作表 |
| 9 | `sheet.unhide` | sheet | 取消隐藏 |
| 10 | `range.insert_rows` | range | 插入行 |
| 11 | `range.delete_rows` | range | 删除行 |
| 12 | `range.insert_columns` | range | 插入列 |
| 13 | `range.delete_columns` | range | 删除列 |
| 14 | `range.sort_data` | range | 排序 |
| 15 | `range.merge` | range | 合并单元格 |
| 16 | `range.unmerge` | range | 拆分单元格 |
| 17 | `range.find_replace` | range | 查找替换 |
| 18 | `range.autofit` | range | 自动适应宽高 |

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

#### automation（7 工具）

| # | 工具 | 说明 |
|--:|------|------|
| 1 | `vba.inspect_project` | 检查 VBA 项目 |
| 2 | `vba.get_module_code` | 获取模块代码 |
| 3 | `vba.sync_module` | 同步模块 |
| 4 | `vba.execute` | 执行宏 |
| 5 | `vba.scan_code` | 扫描代码 |
| 6 | `vba.compile` | 编译项目 |
| 7 | `vba.remove_module` | 移除模块 |

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

#### data（16 工具）

| # | 工具 | Domain | 说明 |
|--:|------|--------|------|
| 1 | `pq.list_connections` | pq | 列出数据连接 |
| 2 | `pq.list_queries` | pq | 列出查询 |
| 3 | `pq.get_code` | pq | 获取查询代码 |
| 4 | `pq.update_query` | pq | 更新查询 |
| 5 | `pq.refresh` | pq | 刷新查询 |
| 6 | `pq.inspect_query` | pq | 查询详细信息 |
| 7 | `pq.refresh_all` | pq | 刷新所有查询 |
| 8 | `pq.get_refresh_status` | pq | 获取刷新状态 |
| 9 | `table.list_tables` | table | 列出所有表格 |
| 10 | `table.create` | table | 创建表格 |
| 11 | `table.inspect` | table | 检查表格结构 |
| 12 | `table.resize` | table | 调整表格范围 |
| 13 | `table.rename` | table | 重命名表格 |
| 14 | `table.set_style` | table | 设置表格样式 |
| 15 | `table.toggle_total_row` | table | 开关总计行 |
| 16 | `table.delete` | table | 删除表格 |

#### workbook_ops（6 工具）

| # | 工具 | Domain | 说明 |
|--:|------|--------|------|
| 1 | `workbook.save_as` | workbook | 另存为 |
| 2 | `workbook.refresh_all` | workbook | 刷新所有连接 |
| 3 | `workbook.calculate` | workbook | 重新计算 |
| 4 | `workbook.list_links` | workbook | 列出外部链接 |
| 5 | `workbook.export_pdf` | workbook | 导出 PDF |
| 6 | `sheet.export_csv` | sheet | 导出 CSV |

---

## 5. Profile 速查

### 5.1 用户 Profile

#### basic_edit — 基础编辑（15 工具）

**场景**：打开 → 查看结构 → 读写数据 → 基础 sheet 管理 → 保存

```
foundation ──────► 8 工具（server + workbook I/O）
edit_basic ──────► 7 工具（sheet 创建/重命名 + range 读写/复制）
                  ─────
                  15 工具   预算 20
```

#### calc_format — 公式与格式（26 工具）

**场景**：基础编辑 + 公式设置 + 格式调整

```
foundation ──────► 8 工具
edit_basic ──────► 7 工具
calc_format ─────► 11 工具（formula 4 + format 7）
                  ─────
                  26 工具   预算 30
```

#### automation — VBA 与恢复（23 工具）

**场景**：VBA 检查/同步/执行 + 快照/备份/回滚

```
foundation ──────► 8 工具
automation ──────► 7 工具（vba 全部）
recovery ────────► 8 工具（snapshot + rollback + backups）
                  ─────
                  23 工具   预算 25
```

#### data_workflow — 数据工作流（30 工具）

**场景**：PQ 管理 → 数据刷新 → Table 操作 → 分析审计

```
foundation ──────► 8 工具
data ────────────► 16 工具（pq 8 + table 8）
analysis ────────► 6 工具
                  ─────
                  30 工具   预算 35
```

#### reporting — 报表输出（20 工具）

**场景**：导出 PDF/CSV + 结构分析 + 未来图表/透视表

```
foundation ──────► 8 工具
workbook_ops ────► 6 工具（save_as + export + 刷新/计算）
analysis ────────► 6 工具
                  ─────
                  20 工具   预算 25
```

### 5.2 开发 Profile（需 `--profiles-file profiles.dev.yaml`）

| Profile | 组合 | 工具数 | 预算 | 用途 |
|---------|------|-------:|----:|------|
| minimal | foundation | 8 | 10 | 最小连通性验证 |
| edit_first | foundation + edit_basic | 15 | 16 | 编辑专题 |
| vba_first | foundation + automation | 15 | 16 | VBA 专题 |
| recovery_first | foundation + recovery | 16 | 18 | 恢复专题 |
| format_first | foundation + calc_format | 19 | 20 | 格式专题 |
| pq_first | foundation + data | 24 | 25 | PQ/数据专题 |
| trae_debug | foundation + edit_basic | 15 | 16 | Trae 调试 |

### 5.3 全量 Profile

| Profile | 工具数 | 说明 |
|---------|-------:|------|
| all | 91 | 全部 Bundle，仅用于 CLI / 回归测试，不推荐在受限客户端中使用 |

---

## 6. 常用查询场景

### "我想做某件事，用哪个 Profile？"

| 我想… | Profile | 工具数 |
|-------|---------|-------:|
| 打开文件读写数据 | `basic_edit` | 15 |
| 设公式调格式 | `calc_format` | 26 |
| 跑 VBA 宏 | `automation` | 23 |
| 管 PQ 查询和 Table | `data_workflow` | 30 |
| 导出 PDF/CSV + 做分析 | `reporting` | 20 |
| 只验证连通性 | `minimal`（dev） | 8 |
| 全量回归测试 | `all` | 91 |

### "某个工具属于哪个 Bundle？在哪些 Profile 中默认可用？"

| 工具 | Bundle | 默认可用的 Profile |
|------|--------|-----------------|
| `workbook.open_file` | foundation | **所有** |
| `range.read_values` | edit_basic | basic_edit, calc_format |
| `range.insert_rows` | edit_structure | 无（需 `--enable-bundle edit_structure`） |
| `formula.set_single` | calc_format | calc_format |
| `vba.execute` | automation | automation |
| `snapshot.manage` | recovery | automation |
| `pq.list_queries` | data | data_workflow |
| `table.create` | data | data_workflow |
| `analysis.scan_structure` | analysis | data_workflow, reporting |
| `workbook.export_pdf` | workbook_ops | reporting |
| `names.inspect` | names | 无（需 `--enable-bundle names`） |

### "我想给当前 Profile 加更多能力"

```bash
# 在 basic_edit 基础上加命名管理
--profile basic_edit --enable-bundle names
# 结果：15 + 4 = 19 工具

# 在 basic_edit 基础上加结构编辑
--profile basic_edit --enable-bundle edit_structure
# 结果：15 + 18 = 33 工具

# 在 calc_format 基础上加结构编辑 + 命名
--profile calc_format --enable-bundle edit_structure --enable-bundle names
# 结果：26 + 18 + 4 = 48 工具（仅建议 CLI 使用）
```

### 客户端推荐

| 客户端 | 推荐 Profile | 避免 | 说明 |
|--------|-------------|------|------|
| Trae AI | `basic_edit` / `*_first` / `trae_debug` | `all` | UI 截断约 39 工具 |
| Workbuddy | 任意用户 Profile | — | 承载能力较好 |
| CLI / run_mcp | `all` / 任意 | — | 无限制 |

---

## 7. 预算规则速记

| 区间 | 含义 | 处理 |
|------|------|------|
| **≤ 30** | 理想 | 正常 |
| **31 ~ 39** | 可接受 | 需设计说明 |
| **≥ 40** | 受限客户端高风险 | 不作为默认场景 Profile |
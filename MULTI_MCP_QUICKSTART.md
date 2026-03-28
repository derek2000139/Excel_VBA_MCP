# ExcelForge V2 Multi-MCP Quickstart

## 1) Start Runtime

```bash
uv run python -m excelforge.runtime --config ./runtime-config.yaml
```

Runtime writes lock metadata to `./.runtime_data_v2/runtime.lock`.

## 2) Start Gateways

Core:

```bash
uv run python -m excelforge.gateway.core --config ./excel-core-mcp.yaml
```

VBA:

```bash
uv run python -m excelforge.gateway.vba --config ./excel-vba-mcp.yaml
```

Recovery:

```bash
uv run python -m excelforge.gateway.recovery --config ./excel-recovery-mcp.yaml
```

PQ (placeholder):

```bash
uv run python -m excelforge.gateway.pq --config ./excel-pq-mcp.yaml
```

## 3) Tool Split

- `excel-core-mcp`: workbook/sheet/range/formula/format/server
- `excel-vba-mcp`: vba tools
- `excel-recovery-mcp`: rollback/backups/audit/names
- `excel-pq-mcp`: pq tools (feature placeholder)

## 4) Notes

- `workbook_id` is created only by Runtime (`workbook.open` / `workbook.create`).
- All gateways pass through `actor_id` for cross-gateway audit traceability.
- V2 uses a new data directory: `./.runtime_data_v2`.

from __future__ import annotations

from dataclasses import dataclass

from excelforge.config import AppConfig, load_config
from excelforge.persistence.audit_repo import AuditRepository
from excelforge.persistence.backup_repo import BackupRepository
from excelforge.persistence.cleanup import CleanupService
from excelforge.persistence.db import Database
from excelforge.persistence.snapshot_repo import SnapshotRepository
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.services.audit_service import AuditService
from excelforge.services.backup_service import BackupService
from excelforge.services.format_service import FormatService
from excelforge.services.formula_service import FormulaService
from excelforge.services.named_range_service import NamedRangeService
from excelforge.services.operation_service import OperationService
from excelforge.services.range_service import RangeService
from excelforge.services.rollback_service import RollbackService
from excelforge.services.server_service import ServerService
from excelforge.services.sheet_service import SheetService
from excelforge.services.snapshot_service import SnapshotService
from excelforge.services.vba_service import VbaService
from excelforge.services.workbook_service import WorkbookService


@dataclass
class RuntimeServices:
    config: AppConfig
    db: Database
    worker: ExcelWorker
    operation_service: OperationService
    server_service: ServerService
    workbook_service: WorkbookService
    sheet_service: SheetService
    range_service: RangeService
    formula_service: FormulaService
    format_service: FormatService
    vba_service: VbaService
    named_range_service: NamedRangeService
    rollback_service: RollbackService
    snapshot_service: SnapshotService
    backup_service: BackupService
    audit_service: AuditService

    def shutdown(self) -> None:
        self.worker.stop(wait_seconds=15)


def create_runtime_services(config_path: str | None = None) -> RuntimeServices:
    config = load_config(config_path)
    db = Database(config)
    db.init_schema()

    worker = ExcelWorker(config)

    audit_repo = AuditRepository(db)
    snapshot_repo = SnapshotRepository(db)
    backup_repo = BackupRepository(db)

    snapshot_service = SnapshotService(config, snapshot_repo)
    backup_service = BackupService(
        config,
        backup_repo,
        workbook_registry=worker.context.registry,
        snapshot_service=snapshot_service,
    )
    workbook_service = WorkbookService(config, worker, snapshot_service)
    sheet_service = SheetService(config, worker, snapshot_service, backup_service)
    range_service = RangeService(config, worker, snapshot_service, backup_service)
    formula_service = FormulaService(config, worker, snapshot_service)
    format_service = FormatService(config, worker)
    vba_service = VbaService(config, worker, backup_service)
    named_range_service = NamedRangeService(config, worker, backup_service)
    rollback_service = RollbackService(config, worker, snapshot_repo, snapshot_service, backup_service)
    audit_service = AuditService(config, audit_repo)

    cleanup_service = CleanupService(config, audit_repo, snapshot_repo, backup_repo)
    operation_service = OperationService(config, audit_service, cleanup_service)
    operation_service.run_cleanup_on_startup()

    server_service = ServerService(config, worker, snapshot_service, backup_service)

    return RuntimeServices(
        config=config,
        db=db,
        worker=worker,
        operation_service=operation_service,
        server_service=server_service,
        workbook_service=workbook_service,
        sheet_service=sheet_service,
        range_service=range_service,
        formula_service=formula_service,
        format_service=format_service,
        vba_service=vba_service,
        named_range_service=named_range_service,
        rollback_service=rollback_service,
        snapshot_service=snapshot_service,
        backup_service=backup_service,
        audit_service=audit_service,
    )

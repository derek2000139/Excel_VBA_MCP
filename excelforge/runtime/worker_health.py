from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Optional

from dataclasses import dataclass, field


class WorkerHealth(Enum):
    HEALTHY = "HEALTHY"
    DEGRADED = "DEGRADED"
    STALE = "STALE"
    RECYCLING = "RECYCLING"
    FAILED = "FAILED"


@dataclass
class WorkerMetrics:
    started_at: datetime = field(default_factory=datetime.utcnow)
    excel_pid: Optional[int] = None
    open_workbooks: int = 0
    operation_count: int = 0
    high_risk_operation_count: int = 0
    exception_count: int = 0
    last_exception_at: Optional[datetime] = None
    last_exception_type: Optional[str] = None
    last_recycle_reason: Optional[str] = None
    last_pid_change_reason: Optional[str] = None
    state: WorkerHealth = WorkerHealth.HEALTHY

    @property
    def uptime_seconds(self) -> float:
        return (datetime.utcnow() - self.started_at).total_seconds()

    def record_operation(self, high_risk: bool = False) -> None:
        self.operation_count += 1
        if high_risk:
            self.high_risk_operation_count += 1

    def record_exception(self, exc_type: str) -> None:
        self.exception_count += 1
        self.last_exception_at = datetime.utcnow()
        self.last_exception_type = exc_type

    def to_dict(self) -> dict:
        return {
            "state": self.state.value,
            "excel_pid": self.excel_pid,
            "open_workbooks": self.open_workbooks,
            "operation_count": self.operation_count,
            "high_risk_operation_count": self.high_risk_operation_count,
            "exception_count": self.exception_count,
            "uptime_seconds": self.uptime_seconds,
            "last_exception_type": self.last_exception_type,
            "last_recycle_reason": self.last_recycle_reason,
            "last_pid_change_reason": self.last_pid_change_reason,
        }

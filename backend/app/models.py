# backend/app/models.py
from __future__ import annotations

from datetime import datetime
from typing import Any, Optional

from sqlalchemy import (
    String,
    Integer,
    Boolean,
    DateTime,
    JSON,
    Identity,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column


class Base(DeclarativeBase):
    """Project base model."""


# -----------------------------
# public.operation_log
# -----------------------------
class OperationLog(Base):
    __tablename__ = "operation_log"

    # id int4 GENERATED ALWAYS AS IDENTITY ...
    id: Mapped[int] = mapped_column(
        Integer,
        Identity(always=True),
        primary_key=True,
    )

    # message_dttm timestamp NOT NULL
    message_dttm: Mapped[datetime] = mapped_column(
        DateTime,
        nullable=False,
        # Uncomment if you want DB to auto-fill with NOW():
        # server_default=func.now(),
    )

    # message json NOT NULL
    message: Mapped[dict[str, Any]] = mapped_column(
        JSON,
        nullable=False,
    )

    def __repr__(self) -> str:
        return f"<OperationLog id={self.id} message_dttm={self.message_dttm!r}>"


# -----------------------------
# workshop.job_cards
# -----------------------------
class JobCard(Base):
    __tablename__ = "job_cards"

    # drive_id varchar PRIMARY KEY
    drive_id: Mapped[str] = mapped_column(
        String, primary_key=True
    )

    # "name" varchar NOT NULL
    name: Mapped[str] = mapped_column(String, nullable=False)

    # creation_dttm timestamp NOT NULL
    creation_dttm: Mapped[datetime] = mapped_column(
        DateTime, nullable=False
        # server_default=func.now(),
    )

    # part_number varchar NOT NULL
    part_number: Mapped[str] = mapped_column(String, nullable=False)

    # serial_number varchar NULL
    serial_number: Mapped[Optional[str]] = mapped_column(String, nullable=True)

    # modified_dttm timestamp NOT NULL
    modified_dttm: Mapped[datetime] = mapped_column(
        DateTime, nullable=False
        # server_default=func.now(),
    )

    # operations json NULL
    operations: Mapped[Optional[dict[str, Any]]] = mapped_column(
        JSON, nullable=True
    )

    # project varchar NULL
    project: Mapped[Optional[str]] = mapped_column(String, nullable=True)

    # is_active bool NOT NULL
    is_active: Mapped[bool] = mapped_column(Boolean, nullable=False)

    def __repr__(self) -> str:
        return f"<JobCard drive_id={self.drive_id!r} part_number={self.part_number!r}>"

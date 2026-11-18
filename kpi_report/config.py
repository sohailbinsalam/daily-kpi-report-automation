"""Loads config.toml into small typed objects."""
from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Kpi:
    column: str
    label: str
    unit: str
    agg: str = "weighted"          # "weighted" or "sum"
    better: str = "higher"         # "higher" or "lower"
    target: float | None = None
    decimals: int = 2

    def meets_target(self, value: float) -> bool | None:
        """True / False against the target, None when the KPI has no target."""
        if self.target is None or value != value:  # NaN check
            return None
        return bool(value >= self.target) if self.better == "higher" else bool(value <= self.target)

    def is_improvement(self, delta: float) -> bool:
        return bool(delta >= 0) if self.better == "higher" else bool(delta <= 0)

    def fmt(self, value: float) -> str:
        if value != value:
            return "n/a"
        return f"{value:,.{self.decimals}f}"


@dataclass
class Config:
    root: Path
    paths: dict[str, Path]
    report: dict
    email: dict
    kpis: list[Kpi] = field(default_factory=list)

    @property
    def kpi_columns(self) -> list[str]:
        return [k.column for k in self.kpis]


def load_config(path: str | Path = "config.toml") -> Config:
    path = Path(path).resolve()
    with path.open("rb") as f:
        raw = tomllib.load(f)

    root = path.parent
    paths = {name: (root / rel) for name, rel in raw["paths"].items()}
    kpis = [Kpi(**k) for k in raw["kpi"]]
    for k in kpis:
        if k.agg not in ("weighted", "sum"):
            raise ValueError(f"{k.column}: agg must be 'weighted' or 'sum'")
        if k.better not in ("higher", "lower"):
            raise ValueError(f"{k.column}: better must be 'higher' or 'lower'")
    return Config(root=root, paths=paths, report=raw["report"], email=raw["email"], kpis=kpis)

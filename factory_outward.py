from __future__ import annotations

from factory_inward import (FactoryInwardResult, add_factory_inward_sheet,
                            default_positions)


def add_factory_outward_sheet(wb, file_bytes, positions, gen_date
                              ) -> FactoryInwardResult:
    """Append the 'Factory Outward' pivot sheet to ``wb``."""
    return add_factory_inward_sheet(
        wb, file_bytes, positions, gen_date, kind="Outward",
        default_trans="DELIVERY CHALLAN", sheet_name="Factory Outward")


def process(file_bytes: bytes, positions=None,
            gen_date: str = "") -> tuple[bytes, FactoryInwardResult]:
    """Build a single-sheet Factory Outward workbook. Returns (bytes, result)."""
    from report_common import new_workbook, workbook_bytes
    if positions is None:
        positions = default_positions()
    wb = new_workbook()
    result = add_factory_outward_sheet(wb, file_bytes, positions, gen_date)
    return workbook_bytes(wb), result

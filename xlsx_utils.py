import os
import tempfile
from typing import Callable

import xlsxwriter


def write_workbook_atomically(
    output_path: str, build_workbook: Callable[[xlsxwriter.Workbook], None]
) -> str:
    output_dir = os.path.dirname(output_path) or "."
    os.makedirs(output_dir, exist_ok=True)

    file_descriptor, temp_path = tempfile.mkstemp(
        suffix=".xlsx", prefix=".tmp-", dir=output_dir
    )
    os.close(file_descriptor)

    workbook = None
    try:
        workbook = xlsxwriter.Workbook(temp_path)
        build_workbook(workbook)
        workbook.close()
        workbook = None
        os.replace(temp_path, output_path)
        return output_path
    except Exception:
        if workbook is not None:
            workbook.close()
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise

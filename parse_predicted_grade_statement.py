import pymupdf
import re
from dataclasses import dataclass
from typing import List

from xlsx_utils import write_workbook_atomically


@dataclass
class PredictedSubjectResult:
    name: str
    exam_board: str
    level: str
    predicted_grade: str
    exam_date: str


@dataclass
class PredictedGradeRecord:
    candidate_name: str
    candidate_english_name: str
    group: str
    subjects: List[PredictedSubjectResult]


@dataclass
class PredictedGradeRectCoefficients:
    name: tuple = (0.1252, 0.2543, 0.5654, 0.2629)
    group: tuple = (0.1252, 0.2971, 0.5654, 0.3057)
    table_start_y: float = 0.4829
    row_margin: float = 0.008
    row_height: float = 0.0371


class PredictedGradeExtractor:
    def __init__(self, dpi: int = 300):
        """
        dpi: Resolution for OCR (higher = better accuracy but slower)
        """
        self.dpi = dpi
        self.coeffs = PredictedGradeRectCoefficients()

    def extract(
        self, pdf_path: str, progress_callback=None
    ) -> list[PredictedGradeRecord]:
        print("Processing document:", pdf_path)
        doc = pymupdf.open(pdf_path)
        try:
            total_pages = len(doc)
            records = []
            for page_num, page in enumerate(doc, 1):
                if progress_callback:
                    progress_callback(page_num, total_pages)
                page_rect = page.rect

                def make_rect(coeff):
                    return pymupdf.Rect(
                        page_rect.x1 * coeff[0],
                        page_rect.y1 * coeff[1],
                        page_rect.x1 * coeff[2],
                        page_rect.y1 * coeff[3],
                    )

                is_electronic = len(page.get_textbox(page.rect)) > 0

                if not is_electronic:
                    print(f"Non-electronic document on page {page_num}, using OCR")
                    textpage = page.get_textpage_ocr(
                        language="eng",
                        dpi=self.dpi,
                        full=True,
                    )
                else:
                    textpage = page.get_textpage()

                name_rect = make_rect(self.coeffs.name)
                group_rect = make_rect(self.coeffs.group)

                name = textpage.extractTextbox(name_rect).strip()
                group = textpage.extractTextbox(group_rect).strip()

                raw_name = name.strip().split(":")[-1].strip()
                group = group.strip().split(":")[-1].strip()

                # Clean up name - title case
                name = raw_name.split("(")[0].strip()
                english_name = ""
                if "(" in raw_name and ")" in raw_name:
                    english_name = raw_name.split("(", 1)[1].rsplit(")", 1)[0].strip()

                subjects: list[PredictedSubjectResult] = []
                row_start_y = self.coeffs.table_start_y

                while row_start_y < 1.0:
                    row_rect = pymupdf.Rect(
                        page_rect.x1 * 0,
                        page_rect.y1 * (row_start_y + self.coeffs.row_margin),
                        page_rect.x1 * 1,
                        page_rect.y1
                        * (
                            row_start_y
                            + self.coeffs.row_height
                            - self.coeffs.row_margin
                        ),
                    )

                    row_text = textpage.extractTextbox(row_rect).strip()

                    print("Extracted row text:")
                    print(row_text)

                    if len(row_text) == 0:
                        break

                    items = [
                        item.strip()
                        for item in re.split(r"\s{2,}|\t|\n", row_text)
                        if item.strip() != "" and item.strip() != "Board"
                    ]

                    if len(items) >= 5:
                        subject_name = items[0]
                        exam_board = items[1]
                        level = items[2]
                        predicted_grade = items[3]
                        exam_date = self._format_date(items[4])

                        subjects.append(
                            PredictedSubjectResult(
                                name=subject_name,
                                exam_board=exam_board,
                                level=level,
                                predicted_grade=predicted_grade,
                                exam_date=exam_date,
                            )
                        )
                    elif len(items) >= 2 and "Subject" not in items[0]:
                        if len(items) >= 3:
                            subjects.append(
                                PredictedSubjectResult(
                                    name=items[0],
                                    exam_board=items[1] if len(items) > 1 else "CIE",
                                    level=items[2] if len(items) > 2 else "GCE A Level",
                                    predicted_grade=items[3] if len(items) > 3 else "",
                                    exam_date=self._format_date(items[4])
                                    if len(items) > 4
                                    else "",
                                )
                            )

                    row_start_y += self.coeffs.row_height

                if not subjects:
                    raise ValueError(f"No subjects extracted on page {page_num}")

                records.append(
                    PredictedGradeRecord(
                        candidate_name=name,
                        candidate_english_name=english_name,
                        group=group,
                        subjects=subjects,
                    )
                )

            return records
        finally:
            doc.close()

    def _format_date(self, date_str: str) -> str:
        """
        Format date from malformed format like 'June ,2026' to 'June 2026'.
        Uses regex to extract month and year separately.
        """
        # Extract month (word characters) and year (4 digits)
        month_match = re.search(r"[A-Za-z]+", date_str)
        year_match = re.search(r"\d{4}", date_str)

        month = month_match.group() if month_match else ""
        year = year_match.group() if year_match else ""

        return f"{month} {year}".strip()

    def extract_all(
        self,
        pdf_paths: list[str],
    ) -> list[PredictedGradeRecord]:
        """
        Extract data from multiple PDF files.
        """
        records = []
        for pdf_path in pdf_paths:
            try:
                record = self.extract(pdf_path)
                records.extend(record)
            except Exception as e:
                print(f"Error processing {pdf_path}: {e}")
        return records

    def write_to_xlsx(
        self,
        records: list[PredictedGradeRecord],
        output_path: str = "predicted_grades.xlsx",
    ) -> str:
        """
        Write predicted grade records to an xlsx file.
        """

        def build_workbook(workbook):
            worksheet = workbook.add_worksheet()

            headers = [
                "Name",
                "English Name",
                "Group",
                "Subject",
                "Exam Board",
                "Level",
                "Predicted Grade",
                "Exam Date",
            ]
            for col, header in enumerate(headers):
                worksheet.write(0, col, header)

            row = 1
            for record in records:
                for subject in record.subjects:
                    worksheet.write(row, 0, record.candidate_name)
                    worksheet.write(row, 1, record.candidate_english_name)
                    worksheet.write(row, 2, record.group)
                    worksheet.write(row, 3, subject.name)
                    worksheet.write(row, 4, subject.exam_board)
                    worksheet.write(row, 5, subject.level)
                    worksheet.write(row, 6, subject.predicted_grade)
                    worksheet.write(row, 7, subject.exam_date)
                    row += 1

        return write_workbook_atomically(output_path, build_workbook)


if __name__ == "__main__":
    extractor = PredictedGradeExtractor(dpi=300)

    for pdf_path in ["test/predicted_grades/example-perdict-grade.pdf"]:
        results = extractor.extract(pdf_path)

        for result in results:
            print(f"\n{'=' * 60}")
            print(f"File: {pdf_path}")
            print(
                f"Candidate: {result.candidate_name} ({result.candidate_english_name})"
            )
            print(f"Group: {result.group}")
            print("Subjects:")
            for sub in result.subjects:
                print(
                    f"  • {sub.name}: {sub.predicted_grade} ({sub.level}) - {sub.exam_date}"
                )

        xlsx_file = extractor.write_to_xlsx(results)
        print(f"\nWritten to: {xlsx_file}")

import fitz  # PyMuPDF
import re
from dataclasses import dataclass
from typing import List, Optional

from xlsx_utils import write_workbook_atomically


@dataclass
class SubjectResult:
    name: str
    grade: str
    syllabus_code: Optional[str] = None
    percentage_mark: Optional[int] = None
    level: Optional[str] = None


@dataclass
class ExamRecord:
    candidate_name: str
    exam_date: str
    school: str
    document_type: str
    subjects: List[SubjectResult]
    candidate_number: Optional[str] = None


@dataclass
class RectCoefficients:
    title: tuple = (0, 0.041, 1, 0.08)
    exam_kind: tuple = (0, 0.231, 1, 0.257)
    name: tuple = (0, 0.115, 0.5, 0.166)
    dob: tuple = (0.646, 0.115, 0.787, 0.166)
    id_num: tuple = (0.787, 0.115, 1, 0.166)
    center_name: tuple = (0, 0.166, 0.5, 0.217)
    series: tuple = (0.646, 0.166, 0.787, 0.217)


class ElectronicRectCoefficients:
    # All y0 and y1 values increased by 0.085
    title: tuple = (0, 0.126, 1, 0.165)
    exam_kind: tuple = (0, 0.403, 1, 0.419)
    name: tuple = (0, 0.2, 0.5, 0.251)
    dob: tuple = (0.646, 0.2, 0.738, 0.251)
    id_num: tuple = (0.738, 0.2, 1, 0.251)
    center_name: tuple = (0, 0.261, 0.5, 0.302)
    series: tuple = (0.646, 0.261, 0.787, 0.302)


def format_str_from_ocr(string: str) -> str:
    return " ".join(
        [
            i
            for i in map(
                lambda x: x.lower() if len(x) <= 3 else (x[0].upper() + x[1:].lower()),
                string.strip().strip(".").strip(":").strip().split(" "),
            )
        ]
    )


class CambridgeOCRExtractor:
    def __init__(self, dpi: int = 300):
        """
        dpi: Resolution for OCR (higher = better accuracy but slower)
        """
        self.dpi = dpi

    def extract(self, pdf_path: str, progress_callback=None) -> list[ExamRecord]:
        print("Processing document:", pdf_path)
        doc = fitz.open(pdf_path)
        try:
            total_pages = len(doc)
            records = []
            for page_num, page in enumerate(doc, 1):
                if progress_callback:
                    progress_callback(page_num, total_pages)
                page_rect = page.rect

                common_coeffs = RectCoefficients()
                electronic_coeffs = ElectronicRectCoefficients()

                def make_rect(coeff):
                    return fitz.Rect(
                        page_rect.x1 * coeff[0],
                        page_rect.y1 * coeff[1],
                        page_rect.x1 * coeff[2],
                        page_rect.y1 * coeff[3],
                    )

                is_electronic = len(page.get_textbox(page.rect)) > 0

                if not is_electronic:
                    print(f"Non-electronic document on page {page_num}, using OCR")
                    page = page.get_textpage_ocr(
                        language="eng",
                        dpi=self.dpi,
                        full=True,
                    )
                else:
                    page = page.get_textpage()

                title_rect = make_rect(
                    common_coeffs.title
                    if not is_electronic
                    else electronic_coeffs.title
                )
                exam_kind_rect = make_rect(
                    common_coeffs.exam_kind
                    if not is_electronic
                    else electronic_coeffs.exam_kind
                )
                name_rect = make_rect(
                    common_coeffs.name if not is_electronic else electronic_coeffs.name
                )
                dob_rect = make_rect(
                    common_coeffs.dob if not is_electronic else electronic_coeffs.dob
                )
                id_num = make_rect(
                    common_coeffs.id_num
                    if not is_electronic
                    else electronic_coeffs.id_num
                )
                center_name_rect = make_rect(
                    common_coeffs.center_name
                    if not is_electronic
                    else electronic_coeffs.center_name
                )
                series_rect = make_rect(
                    common_coeffs.series
                    if not is_electronic
                    else electronic_coeffs.series
                )

                title = page.extractTextbox(title_rect).strip()
                exam_kind = page.extractTextbox(exam_kind_rect).strip()
                name = page.extractTextbox(name_rect).strip()
                dob = page.extractTextbox(dob_rect).strip()
                id_num = page.extractTextbox(id_num).strip()
                center_name = page.extractTextbox(center_name_rect).strip()
                exam_date = page.extractTextbox(series_rect).strip()

                if (
                    len(exam_kind) == 0
                    or len(name) == 0
                    or len(dob) == 0
                    or len(id_num) == 0
                    or len(center_name) == 0
                    or len(exam_date) == 0
                ):
                    if len(exam_kind) == 0:
                        print(f"Could not extract exam kind on page {page_num}")
                    if len(name) == 0:
                        print(f"Could not extract name on page {page_num}")
                    if len(dob) == 0:
                        print(f"Could not extract DOB on page {page_num}")
                    if len(id_num) == 0:
                        print(f"Could not extract ID number on page {page_num}")
                    if len(center_name) == 0:
                        print(f"Could not extract center name on page {page_num}")
                    if len(exam_date) == 0:
                        print(f"Could not extract series on page {page_num}")

                    print(f"Invalid document on page {page_num}, skipping")
                    continue

                # assert len(exam_kind) > 0, "Could not extract exam kind"
                # assert len(name) > 0, "Could not extract name"
                # assert len(dob) > 0, "Could not extract DOB"
                # assert len(id_num) > 0, "Could not extract ID number"
                # assert len(center_name) > 0, "Could not extract center name"
                # assert len(exam_date) > 0, "Could not extract series"

                name = self._extract_value(name, "name", page_num)
                name = " ".join(
                    [
                        part[0].upper() + part[1:].lower()
                        for part in name.split(" ")
                        if part
                    ]
                )
                dob = self._extract_value(dob, "DOB", page_num)
                id_num = self._extract_value(id_num, "ID number", page_num)
                center_name = format_str_from_ocr(
                    self._extract_value(center_name, "center name", page_num)
                )
                exam_date = self._extract_value(exam_date, "series", page_num)

                # print("=======")
                # print("Title: ", title)
                # print("Exam kind: ", exam_kind)
                # print("Name: ", name)
                # print("DOB: ", dob)
                # print("ID num: ", id_num)
                # print("Center name: ", center_name)
                # print("Series: ", exam_date)

                line_height = 0.005
                line_space = 0.013
                subject_start_pos = 0.4336 if "Electronic" in title else 0.2747

                digit_regex = re.compile(r"\d+")

                subjects: list[SubjectResult] = []
                reached_end = False
                item_num = 0
                while not reached_end and subject_start_pos < 1:
                    line_rect = make_rect(
                        (0, subject_start_pos, 1, subject_start_pos + line_height)
                    )
                    line_text = page.extractTextbox(line_rect).strip()
                    subject_start_pos += line_height + line_space

                    if len(line_text) == 0:
                        reached_end = True
                    else:
                        items = re.split(r"[\—\-\n;,.\/\\\=\+]+", line_text)

                        if "Syllabus" in line_text:
                            item_num = len(items)
                            continue

                        if line_text.startswith("With"):
                            if not subjects:
                                raise ValueError(
                                    f"Found continuation grade before any subject on page {page_num}"
                                )
                            last_item = subjects.pop()
                            minor_grade_match = digit_regex.search(line_text)
                            if minor_grade_match:
                                last_item.grade += f"-{minor_grade_match.group()}"
                            subjects.append(last_item)
                            continue

                        syllabus_code = items[0].strip()
                        subject_name = ""
                        qualification = ""
                        grade = ""
                        pum = 0

                        if item_num == 4:
                            subject_name = format_str_from_ocr(" ".join(items[1:-2]))
                            grade = self._parse_grade(items[-2].strip())
                            qualification = exam_kind
                            pum_match = digit_regex.search(items[-1].strip())
                            if pum_match:
                                pum = int(pum_match.group())
                        else:
                            subject_name = format_str_from_ocr(" ".join(items[1:-3]))
                            qualification = format_str_from_ocr(items[-3])
                            grade = self._parse_grade(items[-2].strip())
                            pum_match = digit_regex.search(items[-1].strip())
                            if pum_match:
                                pum = int(pum_match.group())

                        if not subject_name or not grade or not syllabus_code:
                            raise ValueError(
                                f"Incomplete subject row on page {page_num}: {line_text}"
                            )

                        subjects.append(
                            SubjectResult(
                                subject_name, grade, syllabus_code, pum, qualification
                            )
                        )

                if not subjects:
                    raise ValueError(f"No subjects extracted on page {page_num}")

                records.append(
                    ExamRecord(
                        name, exam_date, center_name, "statement", subjects, id_num
                    )
                )

            return records
        finally:
            doc.close()

    def extract_all(
        self,
        pdf_paths: list[str],  # List of PDF file paths to process
    ) -> list[ExamRecord]:  # Return type: List of ExamRecord objects
        """
        Extract data from multiple PDF files.
        Returns a list of ExamRecord objects.
        Args:
            pdf_paths: A list of file paths pointing to PDF files to be processed
        Returns:
            A list of ExamRecord objects containing extracted data from the PDFs
        """
        records = []
        for pdf_path in pdf_paths:
            try:
                record = self.extract(pdf_path)
                records.extend(record)
            except Exception as e:
                print(f"Error processing {pdf_path}: {e}")
        return records

    def _parse_grade(self, grade_line: str) -> str:
        """
        Parse grade from OCR text. Official format is X(Y), like A(a) or A*(a*).
        Use outer part (X) as primary grade, fall back to inner part (Y) if outer is missing.
        """
        grade_line = grade_line.strip()
        if not grade_line:
            return ""

        match = re.match(r"^([A-Z*]{1,2})\(([a-zA-Z*]{1,2})\)\s*$", grade_line)
        if match:
            return match.group(1)

        paren_match = re.search(r"\(([a-zA-Z*]{1,2})\)", grade_line)
        if paren_match:
            inner = paren_match.group(1)
            before_paren = grade_line[: paren_match.start()].strip()
            if before_paren and re.match(r"^[A-Z*]{1,2}$", before_paren):
                return before_paren
            return inner.upper()

        grade_line = grade_line.upper()
        if len(grade_line) == 1:
            return grade_line
        if grade_line[0] != "A":
            return grade_line[0]
        else:
            if grade_line[1] == "*":
                return grade_line[0:2]
            else:
                return grade_line[0]

    def _extract_value(self, raw_text: str, field_name: str, page_num: int) -> str:
        parts = [part.strip() for part in raw_text.split("\n") if part.strip()]
        if len(parts) < 2:
            raise ValueError(
                f"Could not extract {field_name} value on page {page_num}: {raw_text!r}"
            )
        return parts[1]

    def write_to_xlsx(
        self, records: list[ExamRecord], output_path: str = "cie_results.xlsx"
    ) -> str:
        """
        Write multiple exam records to a single xlsx file.
        """

        def build_workbook(workbook):
            worksheet = workbook.add_worksheet()

            headers = [
                "Candidate Name",
                "Candidate Number",
                "School",
                "Exam Date",
                "Document Type",
                "Subject Name",
                "Subject Grade",
                "Subject Level",
                "Syllabus Code",
                "Percentage Mark",
            ]
            for col, header in enumerate(headers):
                worksheet.write(0, col, header)

            row = 1
            for record in records:
                for subject in record.subjects:
                    worksheet.write(row, 0, record.candidate_name)
                    worksheet.write(row, 1, record.candidate_number or "")
                    worksheet.write(row, 2, record.school)
                    worksheet.write(row, 3, record.exam_date)
                    worksheet.write(row, 4, record.document_type)
                    worksheet.write(row, 5, subject.name)
                    worksheet.write(row, 6, subject.grade)
                    worksheet.write(row, 7, subject.level or "")
                    worksheet.write(row, 8, subject.syllabus_code or "")
                    worksheet.write(row, 9, subject.percentage_mark or "")
                    row += 1

        return write_workbook_atomically(output_path, build_workbook)


if __name__ == "__main__":
    extractor = CambridgeOCRExtractor(dpi=300)

    for pdf_path in ["test/statements/statement-combined.pdf"]:
        results = extractor.extract(pdf_path)

        for result in results:
            print(f"\n{'=' * 60}")
            print(f"File: {pdf_path}")
            print(f"Document Type: {result.document_type}")
            print(f"Candidate: {result.candidate_name}")
            print(f"Candidate No: {result.candidate_number or 'N/A'}")
            print(f"School: {result.school}")
            print(f"Exam Date: {result.exam_date}")
            print("Subjects:")
            for sub in result.subjects:
                details = f"  • {sub.name}: {sub.grade}"
                if sub.level:
                    details += f" ({sub.level})"
                if sub.percentage_mark:
                    details += f" - {sub.percentage_mark}%"
                if sub.syllabus_code:
                    details += f" [{sub.syllabus_code}]"
                print(details)

        xlsx_file = extractor.write_to_xlsx(results)
        print(f"Written to: {xlsx_file}")

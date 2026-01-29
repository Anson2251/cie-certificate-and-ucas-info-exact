import fitz  # PyMuPDF
import re
import xlsxwriter
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional


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


class CambridgeOCRExtractor:
    def __init__(self, dpi: int = 300):
        """
        dpi: Resolution for OCR (higher = better accuracy but slower)
        """
        self.dpi = dpi

    def extract(self, pdf_path: str, language: str = "eng") -> ExamRecord:
        """
        Extract data from a single PDF using official get_textpage_ocr() API.
        """
        doc = fitz.open(pdf_path)
        page = doc[0]

        textpage = page.get_textpage_ocr(
            language=language,
            dpi=self.dpi,
            full=True,
        )

        text = textpage.extractText()

        textpage = None
        doc.close()

        if "Electronic Statement of Results" in text:
            return self._parse_statement(text)
        elif (
            "General Certificate of Education" in text or "This certifies that" in text
        ):
            return self._parse_certificate(text)
        else:
            try:
                return self._parse_statement(text)
            except:
                return self._parse_certificate(text)

    def extract_all(
        self, pdf_paths: list[str], language: str = "eng"
    ) -> list[ExamRecord]:
        """
        Extract data from multiple PDF files.
        Returns a list of ExamRecord objects.
        """
        records = []
        for pdf_path in pdf_paths:
            try:
                record = self.extract(pdf_path, language)
                records.append(record)
            except Exception as e:
                print(f"Error processing {pdf_path}: {e}")
        return records

    def _parse_statement(self, text: str) -> ExamRecord:
        lines = [line.strip() for line in text.split("\n") if line.strip()]

        name = next(
            (
                l
                for l in lines
                if l.isupper()
                and len(l.split()) == 2
                and l not in ["GCE AS & A LEVEL", "SYLLABUS TITLE"]
            ),
            "Unknown",
        )
        exam_date = next(
            (
                l
                for l in lines
                if re.search(
                    r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}",
                    l,
                )
            ),
            "Unknown",
        )
        school = next((l for l in lines if "College" in l or "School" in l), "Unknown")
        cand_num = next(
            (l.replace(" ", "") for l in lines if "CN" in l and "/" in l), None
        )

        subjects = []
        for i, line in enumerate(lines):
            if re.match(r"^\d{4,5}$", line) and i + 3 < len(lines):
                code, title, grade, score = (
                    line,
                    lines[i + 1],
                    lines[i + 2],
                    lines[i + 3],
                )
                if score.isdigit():
                    subjects.append(
                        SubjectResult(
                            syllabus_code=code,
                            name=title.title(),
                            grade=grade,
                            percentage_mark=int(score),
                        )
                    )

        return ExamRecord(name, exam_date, school, "statement", subjects, cand_num)

    def _parse_grade(self, grade_line: str) -> str:
        """
        Parse grade from OCR text. Official format is X(Y), like A(a) or A*(a*).
        Use outer part (X) as primary grade, fall back to inner part (Y) if outer is missing.
        """
        grade_line = grade_line.strip()

        match = re.match(r"^([A-Z]{1,2})\(([a-zA-Z*]{1,2})\)\s*$", grade_line)
        if match:
            return match.group(1)

        paren_match = re.search(r"\(([a-zA-Z*]{1,2})\)", grade_line)
        if paren_match:
            inner = paren_match.group(1)
            before_paren = grade_line[: paren_match.start()].strip()
            if before_paren and re.match(r"^[A-Z*]{1,2}$", before_paren):
                return before_paren
            return inner.upper()

        return grade_line.upper()

    def _parse_certificate(self, text: str) -> ExamRecord:
        alevel = True
        if text.find("International General Certificate of Secondary Education"):
            alevel = False
        name_match = re.search(
            r"This certifies that in the [^\n]+\n([A-Z][A-Z\s]+?)\n\s*of\s+([^\n]+)",
            text,
        )
        name = name_match.group(1).strip() if name_match else "Unknown"
        school = name_match.group(2).strip() if name_match else "Unknown"

        date_match = re.search(r"([A-Za-z]+\s+\d{4})\s+examination series", text)
        exam_date = date_match.group(1) if date_match else "Unknown"

        cand_match = re.search(r"Candidate Number[.:]\s*([A-Z0-9/]+)", text)
        candidate_num = cand_match.group(1).strip() if cand_match else None

        subjects = []

        section_match = re.search(
            r"shown:\s*(.*?)\s*SYLLABUSES\s+REPORTED", text, re.DOTALL | re.IGNORECASE
        )
        if not section_match:
            return ExamRecord(
                name,
                exam_date,
                school.replace("of ", "", 1),
                "certificate",
                subjects,
                candidate_num,
            )

        section_text = section_match.group(1)

        lines = [line.strip() for line in section_text.splitlines() if line.strip()]

        while lines and lines[0].upper() in ["SYLLABUS", "GRADE"]:
            lines.pop(0)

        i = 0
        while i + 2 < len(lines):
            subject_name = lines[i]

            if (alevel):
                level_line = lines[i + 1]
                grade_line = lines[i + 2]
            else:
                grade_line = lines[i + 1]
                level_line = "IGCSE"

            is_valid = (
                ("Advanced" in level_line
                and ("Level" in level_line or "Subsidiary" in level_line) or level_line == "IGCSE")
                and subject_name
                and grade_line
            )

            if is_valid:
                level = ("A Level" if "Subsidiary" not in level_line else "AS Level") if alevel else level_line
                grade = self._parse_grade(grade_line)

                subjects.append(
                    SubjectResult(name=subject_name, grade=grade, level=level)
                )
                i += (3 if alevel else 2)
            else:
                i += 1

        return ExamRecord(
            candidate_name=name,
            exam_date=exam_date,
            school=school.replace("of ", "", 1),
            document_type="certificate",
            subjects=subjects,
            candidate_number=candidate_num,
        )
        
    
    def write_to_xlsx(
        self, records: list[ExamRecord], output_path: str = "cie_results.xlsx"
    ) -> str:
        """
        Write multiple exam records to a single xlsx file.
        """
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet()

        headers = [
            "candidate_name",
            "candidate_number",
            "school",
            "exam_date",
            "document_type",
            "subject_name",
            "subject_grade",
            "subject_level",
            "syllabus_code",
            "percentage_mark",
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

        workbook.close()
        return output_path


if __name__ == "__main__":
    extractor = CambridgeOCRExtractor(dpi=300)

    for pdf_path in [
        "test/cs/General Certificate of Education-2024-June-IG.pdf",
    ]:
        try:
            result = extractor.extract(pdf_path)

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

            xlsx_file = extractor.write_to_xlsx([result])
            print(f"Written to: {xlsx_file}")

        except Exception as e:
            print(f"Error with {pdf_path}: {e}")

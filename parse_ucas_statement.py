import re
from typing_extensions import Callable
import pymupdf
import xlsxwriter
from dataclasses import dataclass
from more_itertools import peekable


PDF_PAGE_DIM = (612.0, 792.0)
MARGIN_TOP = 60
MARGIN_BOTTOM = 40


@dataclass
class EducationEntry:
    school_name: str
    qualification_category: str
    subject_name: str
    subject_grade: str
    subject_date: str
    subject_awarding_org: str
    subject_country: str

@dataclass
class UCASData:
    name: str
    group: str
    education: list[EducationEntry]
    personal_statement: tuple[str, str, str] = ("", "", "")  # (question 1, question 2, question 3)


class UCASExtractor:
    def __init__(self, pdf_path: str = "./test/ucas/example-ucas.pdf"):
        self.pdf_path = pdf_path

    def extract(self, progress_callback: Callable | None = None) -> UCASData:
        doc = pymupdf.open(self.pdf_path)
        total_pages = doc.page_count
        print("Number of pages: ", total_pages)

        if progress_callback:
            progress_callback(1, total_pages)

        pdf_dim = (doc[0].rect.width, doc[0].rect.height)
        print("Dimension: ", pdf_dim)

        name, class_group = self._find_name_class(doc)

        education_section = self._get_education_section(doc)[0]
        education_title_page = education_section[0]
        education_title_rect = education_section[1][0]

        employment_section = self._get_employment_section(doc)[0]
        employment_title_page = employment_section[0]
        employment_title_rect = employment_section[1][0]

        print(education_title_page, employment_title_page)

        if progress_callback:
            progress_callback(min(2, total_pages), total_pages)

        education_raw_text = self._get_text_between(
            doc,
            education_title_page,
            employment_title_page,
            pymupdf.Point(education_title_rect.x0, education_title_rect.y1 + 60),
            pymupdf.Point(pdf_dim[0], employment_title_rect.y0),
        )

        personal_statement_section = self._get_personal_statement_section(doc)[0]
        personal_statement_title_page = personal_statement_section[0]
        personal_statement_title_rect = personal_statement_section[1][0]

        choices_section = self._get_choices_section(doc)[0]
        choices_title_page = choices_section[0]
        choices_title_rect = choices_section[1][0]

        personal_statement_raw_text = self._get_text_between(
            doc,
            personal_statement_title_page,
            choices_title_page,
            pymupdf.Point(personal_statement_title_rect.x0, personal_statement_title_rect.y1 + 24),
            pymupdf.Point(pdf_dim[0], choices_title_rect.y0) - 20,
        )

        personal_statement = self._parse_personal_statement(personal_statement_raw_text)
        education_info = self._parse_education_info(education_raw_text)

        if progress_callback:
            progress_callback(total_pages, total_pages)

        return UCASData(
            name=name,
            group=class_group,
            education=education_info,
            personal_statement=personal_statement,
        )

    def _find_rects_in_pdf(self, doc, search_string) -> list[tuple[int, list[pymupdf.Rect]]]:
        results: list[tuple[int, list[pymupdf.Rect]]] = []
        for page in doc:
            areas = page.search_for(search_string)
            if areas:
                results.append((page.number, areas))
        return results

    def _get_sections(self, doc: pymupdf.Document, title: str):
        return [
            j
            for j in filter(
                lambda p: len([i for i in p[1] if i.height > 16]) > 0,
                self._find_rects_in_pdf(doc, title),
            )
        ]

    def _get_education_section(self, doc: pymupdf.Document):
        titles = self._get_sections(doc, "Education")
        assert len(titles) > 0, "Education title not found"
        return titles

    def _get_personal_statement_section(self, doc: pymupdf.Document):
        titles = self._get_sections(doc, "Personal statement")
        assert len(titles) > 0, "Personal statement title not found"
        return titles

    def _get_choices_section(self, doc: pymupdf.Document):
        titles = self._get_sections(doc, "Choices")
        assert len(titles) > 0, "Choices title not found"
        return titles

    def _get_employment_section(self, doc: pymupdf.Document):
        titles = self._get_sections(doc, "Employment")
        assert len(titles) > 0, "Employment title not found"
        return titles

    def _get_text_between(
        self,
        doc: pymupdf.Document,
        page_start: int,
        page_end: int,
        pos_start: pymupdf.Point,
        pos_end: pymupdf.Point,
    ):
        result = ""
        for page_num in range(page_start, page_end + 1):
            page = doc[page_num]
            if page_num == page_start and page_num == page_end:
                rect = pymupdf.Rect(pos_start.x, pos_start.y, pos_end.x, pos_end.y)
            elif page_num == page_start:
                rect = pymupdf.Rect(
                    page.rect.x0,
                    pos_start.y,
                    page.rect.x1,
                    page.rect.y1 - MARGIN_BOTTOM,
                )
            elif page_num == page_end:
                rect = pymupdf.Rect(page.rect.x0, MARGIN_TOP, pos_end.x, pos_end.y)
            else:
                rect = pymupdf.Rect(
                    page.rect.x0, MARGIN_TOP, page.rect.x1, page.rect.y1 - MARGIN_BOTTOM
                )
            result += page.get_textbox(rect)
            result += "\n"
        return result.strip()

    def _find_name_class(self, doc: pymupdf.Document):
        info_rect = pymupdf.Rect(
            0,
            0,
            PDF_PAGE_DIM[0] * 0.5,
            50,
        )
        raw_text = doc[0].get_textbox(info_rect)

        assert len(raw_text) > 0, "No header text found in first page"

        lines = raw_text.split("\n")
        name_match = re.search(r"^([A-Za-z]+\s[A-Za-z]+)", lines[0])
        class_match = re.search(r"Group:\s*([^;]+)", lines[-1])

        name = ""
        class_group = ""

        if name_match:
            name = name_match.group(1)

        if class_match:
            class_group = class_match.group(1)

        return name, class_group

    def _parse_personal_statement(self, raw_text: str) -> tuple[str, str, str]:
        questions = [
            "Why do you want to study this course or subject?",
            "How have your qualifications and studies helped you to prepare for this course or subject?",
            "What else have you done to prepare outside of education, and why are these experiences useful?",
        ]

        sections = []
        lines = [i.strip() for i in raw_text.split("\n") if (len(i.strip()) > 0)]

        question_indices = [
            i for i, line in enumerate(lines) if any(q in line for q in questions)
        ]
        question_indices.append(len(lines))

        for idx in range(len(question_indices) - 1):
            start_index = question_indices[idx]
            end_index = question_indices[idx + 1]

            section_lines = lines[start_index + 1 : end_index]
            section_text = " ".join(section_lines).strip()
            sections.append(section_text)

        while len(sections) < 3:
            sections.append("")

        return sections[0], sections[1], sections[2]

    def _parse_education_info(self, raw_text: str) -> list[EducationEntry]:
        lines = [i.strip() for i in raw_text.split("\n") if (len(i.strip()) > 0)]

        qualification_end_index = lines.index("Unique Learner Number (ULN):")

        indices = peekable(
            map(
                lambda x: x - 1,
                [i for i, x in enumerate(lines) if x == "National centre number:"],
            )
        )
        results = []

        for i in indices:
            end_index = indices.peek(-1)
            end_index = end_index if end_index >= 0 else qualification_end_index

            school_name = lines[i]
            qualification_category = ""
            subject_name = ""
            subject_grade = ""
            subject_date = ""
            subject_awarding_org = ""
            subject_country = ""

            qualification_lines = peekable(lines[i + 4 : end_index])
            for j in qualification_lines:
                if j.startswith("Grade:") or j.startswith("Result:"):
                    subject_grade = j.split(":")[1].strip()
                elif j.startswith("Qualification date:"):
                    subject_date = j.split(":")[1].strip()
                    next_line = qualification_lines.peek("")
                    if len(next_line) == 4 and next_line.isnumeric():
                        subject_date += f" {next(qualification_lines)}"
                elif j.startswith("Awarding organisation:"):
                    subject_awarding_org = j.split(":")[1].strip()
                elif j.startswith("Country:"):
                    subject_country = j.split(":")[1].strip()
                elif ":" not in j:
                    is_valid_subject = (
                        subject_name != ""
                        and subject_date != ""
                        and subject_awarding_org != ""
                    )
                    if is_valid_subject:
                        results.append(
                            EducationEntry(
                                school_name=school_name,
                                qualification_category=qualification_category,
                                subject_name=subject_name,
                                subject_grade=subject_grade if subject_grade else "N/A",
                                subject_date=subject_date,
                                subject_awarding_org=subject_awarding_org,
                                subject_country=subject_country,
                            )
                        )
                        subject_name = ""
                        subject_grade = ""
                        subject_date = ""
                        subject_country = ""

                    if ":" not in qualification_lines.peek():
                        qualification_category = j.strip()
                    else:
                        subject_name = j.strip()

        return results

    def write_to_xlsx(
        self, all_data: list[UCASData], output_path: str = "ucas_results.xlsx"
    ) -> str:
        """
        Write multiple UCAS data records to a single xlsx file.
        """
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet("Education")

        headers = [
            "Name",
            "Group",
            "School Name",
            "Qualification",
            "Subject",
            "Grade",
            "Date",
            "Awarding Organisation",
            "Country"
        ]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        row = 1
        for data in all_data:
            for entry in data.education:
                worksheet.write(row, 0, data.name)
                worksheet.write(row, 1, data.group)
                worksheet.write(row, 2, entry.school_name)
                worksheet.write(row, 3, entry.qualification_category)
                worksheet.write(row, 4, entry.subject_name)
                worksheet.write(row, 5, entry.subject_grade)
                worksheet.write(row, 6, entry.subject_date)
                worksheet.write(row, 7, entry.subject_awarding_org)
                worksheet.write(row, 8, entry.subject_country)
                row += 1

        worksheet = workbook.add_worksheet("Personal Statement")
        headers = [
            "Name",
            "Group",
            "Question 1",
            "Question 2",
            "Question 3"
        ]
        row = 1
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        for data in all_data:
            worksheet.write(row, 0, data.name)
            worksheet.write(row, 1, data.group)
            worksheet.write(row, 2, data.personal_statement[0])
            worksheet.write(row, 3, data.personal_statement[1])
            worksheet.write(row, 4, data.personal_statement[2])
            row += 1

        workbook.close()
        return output_path


if __name__ == "__main__":
    extractor = UCASExtractor()
    result = extractor.extract()
    extractor.write_to_xlsx([result])

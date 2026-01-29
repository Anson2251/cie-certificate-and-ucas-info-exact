import pymupdf
import xlsxwriter
from dataclasses import dataclass
from datetime import datetime
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
    first_name: str
    last_name: str
    education: list[EducationEntry]


class UCASExtractor:
    def __init__(self, pdf_path: str = "./example-ucas.pdf"):
        self.pdf_path = pdf_path

    def extract(self) -> UCASData:
        doc = pymupdf.open(self.pdf_path)
        print("Number of pages: ", doc.page_count)

        pdf_dim = (doc[0].rect.width, doc[0].rect.height)
        print("Dimension: ", pdf_dim)

        first_name, last_name = self._find_name(doc, pdf_dim)

        education_section = self._get_education_section(doc)[0]
        education_title_page = education_section[0]
        education_title_rect = education_section[1][0]

        employment_section = self._get_employment_section(doc)[0]
        employment_title_page = employment_section[0]
        employment_title_rect = employment_section[1][0]

        print(education_title_page, employment_title_page)

        education_raw_text = self._get_text_between(
            doc,
            education_title_page,
            employment_title_page,
            pymupdf.Point(education_title_rect.x0, education_title_rect.y1 + 60),
            pymupdf.Point(pdf_dim[0], employment_title_rect.y0),
        )

        education_info = self._parse_education_info(education_raw_text)

        return UCASData(
            first_name=first_name,
            last_name=last_name,
            education=education_info,
        )

    def _find_rects_in_pdf(self, doc, search_string):
        results = []
        for page in doc:
            areas = page.search_for(search_string)
            if areas:
                results.append((page.number, areas))
        return results

    def _get_sections(self, doc, title):
        return [
            j
            for j in filter(
                lambda p: len([i for i in p[1] if i.height > 16]) > 0,
                self._find_rects_in_pdf(doc, title),
            )
        ]

    def _get_education_section(self, doc):
        titles = self._get_sections(doc, "Education")
        assert len(titles) > 0, "Education title not found"
        return titles

    def _get_employment_section(self, doc):
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

    def _find_name(self, doc: pymupdf.Document, dim):
        (first_name_page, first_name_rects) = self._find_rects_in_pdf(
            doc, "First and middle name(s)"
        )[0]
        (last_name_page, last_name_rects) = self._find_rects_in_pdf(doc, "Last name")[0]

        assert first_name_page != -1 and len(first_name_rects) > 0, (
            "Student first name not found"
        )
        assert last_name_page != -1 and len(last_name_rects) > 0, (
            "Student last name not found"
        )

        first_name_rect = pymupdf.Rect(
            first_name_rects[0].x1,
            first_name_rects[0].y0,
            dim[0],
            first_name_rects[0].y1,
        )
        last_name_rect = pymupdf.Rect(
            last_name_rects[0].x1, last_name_rects[0].y0, dim[0], last_name_rects[0].y1
        )

        first_name = doc[first_name_page].get_textbox(first_name_rect)
        last_name = doc[last_name_page].get_textbox(last_name_rect)

        return first_name, last_name

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
        self, data: UCASData, output_path: str = "ucas_results.xlsx"
    ) -> str:
        """
        Write UCAS data to xlsx file.
        """
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet()

        headers = [
            "name",
            "school_name",
            "qualification_category",
            "subject_name",
            "subject_grade",
            "subject_date",
            "subject_awarding_org",
            "subject_country",
        ]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        full_name = f"{data.first_name} {data.last_name}".strip()

        for row, entry in enumerate(data.education, start=1):
            worksheet.write(row, 0, full_name)
            worksheet.write(row, 1, entry.school_name)
            worksheet.write(row, 2, entry.qualification_category)
            worksheet.write(row, 3, entry.subject_name)
            worksheet.write(row, 4, entry.subject_grade)
            worksheet.write(row, 5, entry.subject_date)
            worksheet.write(row, 6, entry.subject_awarding_org)
            worksheet.write(row, 7, entry.subject_country)

        workbook.close()
        return output_path

    def write_combined_to_xlsx(
        self, all_data: list[UCASData], output_path: str = "ucas_results.xlsx"
    ) -> str:
        """
        Write multiple UCAS data records to a single xlsx file.
        """
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet()

        headers = [
            "name",
            "school_name",
            "qualification_category",
            "subject_name",
            "subject_grade",
            "subject_date",
            "subject_awarding_org",
            "subject_country",
        ]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        row = 1
        for data in all_data:
            full_name = f"{data.first_name} {data.last_name}".strip()
            for entry in data.education:
                worksheet.write(row, 0, full_name)
                worksheet.write(row, 1, entry.school_name)
                worksheet.write(row, 2, entry.qualification_category)
                worksheet.write(row, 3, entry.subject_name)
                worksheet.write(row, 4, entry.subject_grade)
                worksheet.write(row, 5, entry.subject_date)
                worksheet.write(row, 6, entry.subject_awarding_org)
                worksheet.write(row, 7, entry.subject_country)
                row += 1

        workbook.close()
        return output_path


if __name__ == "__main__":
    extractor = UCASExtractor()
    result = extractor.extract()

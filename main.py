import pymupdf
import xlsxwriter
from datetime import datetime
from more_itertools import peekable

PDF_PAGE_DIM = (612.0, 792.0)
MARGIN_TOP = 60
MARGIN_BOTTOM = 40


def find_rects_in_pdf(doc, search_string) -> list[tuple[int, list[pymupdf.Rect]]]:
    results = []
    for page in doc:
        areas = page.search_for(search_string)
        if areas:
            results.append((page.number, areas))
    return results


def get_sections(doc, title):
    return [
        j
        for j in filter(
            lambda p: len([i for i in p[1] if i.height > 16]) > 0,
            find_rects_in_pdf(doc, title),
        )
    ]


def get_text_between(
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
                page.rect.x0, pos_start.y, page.rect.x1, page.rect.y1 - MARGIN_BOTTOM
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


def main():
    print(pymupdf.__doc__)
    doc = pymupdf.open("./example-ucas.pdf")
    print("Number of pages: ", doc.page_count)
    PDF_PAGE_DIM = (doc[0].rect.width, doc[0].rect.height)
    print("Dimension: ", PDF_PAGE_DIM)
    name = " ".join(find_name(doc, PDF_PAGE_DIM))
    education_titles = get_sections(doc, "Education")
    employment_titles = get_sections(doc, "Employment")

    assert len(education_titles) > 0, "Education title not found"
    assert len(employment_titles) > 0, "Employment title not found"

    education_title_page = education_titles[0][0]
    education_title_rect = education_titles[0][1][0]

    employment_title_page = employment_titles[0][0]
    employment_title_rect = employment_titles[0][1][0]

    print(education_title_page, employment_title_page)

    education_raw_text = get_text_between(
        doc,
        education_title_page,
        employment_title_page,
        pymupdf.Point(education_title_rect.x0, education_title_rect.y1 + 60),
        pymupdf.Point(PDF_PAGE_DIM[0], employment_title_rect.y0),
    )
    education_info = parse_education_info(education_raw_text)
    first_name, last_name = find_name(doc, PDF_PAGE_DIM)
    write_education_info_to_xlsx(education_info, name)


def parse_education_info(raw_text: str):
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
            elif ":" not in j :
                is_valid_subject = (
                    subject_name != ""
                    and subject_date != ""
                    and subject_awarding_org != ""
                )
                if is_valid_subject:
                    results.append(
                        {
                            "school_name": school_name,
                            "qualification_category": qualification_category,
                            "subject_name": subject_name,
                            "subject_grade": subject_grade if subject_grade else "N/A",
                            "subject_date": subject_date,
                            "subject_awarding_org": subject_awarding_org,
                            "subject_country": subject_country,
                        }
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


def find_name(doc: pymupdf.Document, dim):
    (first_name_page, first_name_rects) = find_rects_in_pdf(
        doc, "First and middle name(s)"
    )[0]
    (last_name_page, last_name_rects) = find_rects_in_pdf(doc, "Last name")[0]

    assert first_name_page != -1 and len(first_name_rects) > 0, (
        "Student first name not found"
    )
    assert last_name_page != -1 and len(last_name_rects) > 0, (
        "Student last name not found"
    )

    first_name_rect = pymupdf.Rect(
        first_name_rects[0].x1, first_name_rects[0].y0, dim[0], first_name_rects[0].y1
    )
    last_name_rect = pymupdf.Rect(
        last_name_rects[0].x1, last_name_rects[0].y0, dim[0], last_name_rects[0].y1
    )

    first_name = doc[first_name_page].get_textbox(first_name_rect)
    last_name = doc[last_name_page].get_textbox(last_name_rect)

    return first_name, last_name


def write_education_info_to_xlsx(education_info: list[dict], current_name: str):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{current_name}_{timestamp}.xlsx"
    workbook = xlsxwriter.Workbook(filename)
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

    info = map(lambda i: {**i, "name": current_name}, education_info)

    for row, info in enumerate(info, start=1):
        for col, key in enumerate(headers):
            worksheet.write(row, col, info.get(key, ""))

    workbook.close()
    return filename


if __name__ == "__main__":
    main()

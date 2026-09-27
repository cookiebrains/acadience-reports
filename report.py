import csv
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt


GREEN = RGBColor(30, 115, 64)
YELLOW = RGBColor(220, 165, 0)
RED = RGBColor(175, 50, 50)

BELOW_BENCHMARK = ("below benchmark", "well below benchmark")
AT_OR_ABOVE = ("at benchmark", "above benchmark")


def get_homeroom_teacher(class_name):
    """Get the first teacher from a name such as Grade 2 (Broderick, Merrill)."""
    if "(" not in class_name or ")" not in class_name:
        raise ValueError(f"Cannot find homeroom teacher in: {class_name!r}")

    teachers = class_name.split("(", 1)[1].split(")", 1)[0]
    homeroom_teacher = teachers.split(",", 1)[0].strip()

    if not homeroom_teacher:
        raise ValueError(f"Homeroom teacher is blank in: {class_name!r}")

    return homeroom_teacher


def categorize_students(students, beginning_status):
    categories = {
        "at_above_growing": [],
        "at_above_slow": [],
        "below_fast": [],
        "below_typical": [],
        "below_slow": [],
    }

    for student in sorted(
        students,
        key=lambda row: (
            row["Student Last Name"].lower(),
            row["Student First Name"].lower(),
        ),
    ):
        name = f"{student['Student First Name']} {student['Student Last Name']}"
        status = student["Reading Composite Status"].strip().lower()
        growth = student["Reading Composite Pathway"].strip().lower()
        student_id = student["Student Number"].strip()
        boy_status = beginning_status.get(student_id)

        if boy_status in BELOW_BENCHMARK and status in AT_OR_ABOVE:
            name += " ⭐"

        if status in AT_OR_ABOVE and growth in (
            "typical", "above typical", "well above typical"
        ):
            categories["at_above_growing"].append(name)

        elif status in AT_OR_ABOVE and growth in (
            "below typical", "well below typical"
        ):
            categories["at_above_slow"].append(name)

        elif status in BELOW_BENCHMARK and growth in (
            "above typical", "well above typical"
        ):
            categories["below_fast"].append(name)

        elif status in BELOW_BENCHMARK and growth == "typical":
            categories["below_typical"].append(name)

        elif status in BELOW_BENCHMARK and growth in (
            "below typical", "well below typical"
        ):
            categories["below_slow"].append(name)

        else:
            raise ValueError(
                f"Could not categorize {name}: {status} / {growth}"
            )

    return categories


def add_list(slide, benchmark, growth, students_in_group, color, x, y, width):
    # Heading
    heading_box = slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(width), Inches(0.65)
    )
    heading_box.text = benchmark + "\n" + growth

    benchmark_run = heading_box.text_frame.paragraphs[0].runs[0]
    benchmark_run.font.size = Pt(16)
    benchmark_run.font.bold = True

    growth_run = heading_box.text_frame.paragraphs[1].runs[0]
    growth_run.font.size = Pt(13)
    growth_run.font.bold = True
    growth_run.font.color.rgb = color

    # Divider
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y + 0.62), Inches(width), Inches(0.015)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(210, 220, 225)
    line.line.fill.background()

    # Student names
    if len(students_in_group) > 15:
        columns = 3
    elif len(students_in_group) > 8:
        columns = 2
    else:
        columns = 1

    rows_per_column = (len(students_in_group) + columns - 1) // columns
    column_width = width / columns

    for index, name in enumerate(students_in_group):
        column = index // rows_per_column
        row = index % rows_per_column

        name_box = slide.shapes.add_textbox(
            Inches(x + column * column_width),
            Inches(y + 0.78 + row * (0.25 if columns == 1 else 0.24)),
            Inches(column_width),
            Inches(0.20),
        )
        name_box.text = name
        name_box.text_frame.paragraphs[0].runs[0].font.size = Pt(
            13 if columns == 1 else 12
        )


def make_slide(prs, teacher, grade, students, beginning_status):
    categories = categorize_students(students, beginning_status)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.35), Inches(4), Inches(1)
    )
    title.text = teacher

    teacher_run = title.text_frame.paragraphs[0].runs[0]
    teacher_run.font.size = Pt(20)
    teacher_run.font.bold = True

    grade_line = title.text_frame.add_paragraph()
    grade_line.text = f"Grade {grade}"
    grade_line.runs[0].font.size = Pt(16)

    # At or above benchmark: two sections across the top
    add_list(
        slide, "At or above benchmark", "Typical or better growth",
        categories["at_above_growing"], GREEN, 0.5, 1.4, 4.8
    )
    add_list(
        slide, "At or above benchmark", "Below typical growth",
        categories["at_above_slow"], YELLOW, 5.7, 1.4, 4.8
    )

    # Below benchmark: three sections across the bottom
    add_list(
        slide, "Below benchmark", "Above or well above typical growth",
        categories["below_fast"], GREEN, 0.5, 4.2, 3.0
    )
    add_list(
        slide, "Below benchmark", "Typical growth",
        categories["below_typical"], RED, 4.0, 4.2, 3.0
    )
    add_list(
        slide, "Below benchmark", "Below or well below typical growth",
        categories["below_slow"], RED, 7.5, 4.2, 3.0
    )

    footnote = slide.shapes.add_textbox(
        Inches(0.5), Inches(8.05), Inches(10), Inches(0.3)
    )
    footnote.text = (
        "⭐ Starfish: moved from below benchmark at BOY "
        "to at or above benchmark at MOY."
    )
    footnote.text_frame.paragraphs[0].runs[0].font.size = Pt(10)


def main():
    csv_path = Path(__file__).with_name("moy_three_class_test.csv")
    with csv_path.open(newline="", encoding="utf-8-sig") as file:
        all_rows = list(csv.DictReader(file))

    beginning_status = {}
    groups = {}

    for row in all_rows:
        period = row["Benchmark Period"].strip().lower()
        student_id = row["Student Number"].strip()

        if period == "beginning":
            beginning_status[student_id] = (
                row["Reading Composite Status"].strip().lower()
            )

        elif period == "middle":
            homeroom_teacher = get_homeroom_teacher(
                row["Class Name"].strip()
            )
            grade = row["Student Grade Level"].strip()

            key = (homeroom_teacher, grade)
            groups.setdefault(key, []).append(row)

    prs = Presentation()
    prs.slide_width = Inches(11)
    prs.slide_height = Inches(8.5)

    for (teacher, grade), students in sorted(groups.items()):
        make_slide(prs, teacher, grade, students, beginning_status)
        print(f"{teacher}, Grade {grade}: {len(students)} students")

    output_path = Path(__file__).with_name("practice.pptx")
    prs.save(output_path)
    print(f"Created {len(prs.slides)} slides: {output_path.name}")


if __name__ == "__main__":
    main()
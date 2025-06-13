import wikipedia
import re
from pptx import Presentation
from pptx.util import Pt

def get_filtered_sections(title):
    try:
        page = wikipedia.page(title)
        content = page.content
        # Only top-level sections
        sections = re.findall(r'^==\s*([^=]+?)\s*==$', content, flags=re.MULTILINE)
        filtered = [s.strip() for s in sections 
                    if s.lower() not in ('references', 'external links', 'see also', 'notes')]
        return filtered
    except Exception as e:
        print(f"Section error: {e}")
        return []

def create_custom_ppt(main_title, selected_sections):
    prs = Presentation()
    # Title Slide
    prs.slides.add_slide(prs.slide_layouts[0]).shapes.title.text = main_title

    for section in selected_sections:
        try:
            content = wikipedia.page(main_title).section(section)
            if not content:
                continue

            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = section

            points = re.findall(r'[^.!?]*[.!?]', content)[:5]

            tf = slide.placeholders[1].text_frame
            tf.clear()

            for point in points:
                p = tf.add_paragraph()
                p.text = point.strip()
                p.level = 0
                for run in p.runs:
                    run.font.size = Pt(20)
        except Exception as e:
            print(f"Error in {section}: {e}")

    output_file = f"{main_title.replace(' ', '_')}_presentation.pptx"
    prs.save(output_file)
    print(f"PPT saved as {output_file}")

def main():
    main_title = input("Enter the Wikipedia topic: ").strip()
    sections = get_filtered_sections(main_title)
    if not sections:
        print("No sections found.")
        return

    print("\nAvailable subtopics/sections:")
    for idx, section in enumerate(sections):
        print(f"{idx+1}. {section}")

    print("\nEnter the numbers of the sections you want to include, separated by commas (e.g., 1,3,5):")
    selected = input().strip()
    try:
        indices = [int(x)-1 for x in selected.split(",") if x.strip().isdigit()]
        selected_sections = [sections[i] for i in indices if 0 <= i < len(sections)]
    except Exception as e:
        print("Invalid input. Please enter valid numbers separated by commas.")
        return

    if not selected_sections:
        print("No valid sections selected.")
        return

    create_custom_ppt(main_title, selected_sections)

if __name__ == "__main__":
    main()

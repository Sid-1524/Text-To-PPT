import wikipedia
import re
from pptx import Presentation
from pptx.util import Pt

MAX_BULLET_CHARS = 120  # Maximum characters per bullet point
MAX_TOTAL_CHARS = 500   # Maximum total characters per slide content

def clean_input(user_input):
    cleaned = re.sub(r'[^a-zA-Z0-9 ]', '', user_input)
    return re.sub(r'\s+', ' ', cleaned).strip()

def get_wikipedia_page(title):
    try:
        return wikipedia.page(title, auto_suggest=False)
    except wikipedia.exceptions.PageError:
        results = wikipedia.search(title)
        if not results:
            return None
        print("Did you mean:")
        for i, res in enumerate(results[:5], 1):
            print(f"{i}. {res}")
        choice = input("Enter number or 0 to cancel: ").strip()
        if choice.isdigit() and 0 < int(choice) <= len(results):
            return wikipedia.page(results[int(choice)-1], auto_suggest=False)
        return None
    except wikipedia.exceptions.DisambiguationError as e:
        print("Multiple matches found:")
        for i, opt in enumerate(e.options[:5], 1):
            print(f"{i}. {opt}")
        choice = input("Enter number or 0 to cancel: ").strip()
        if choice.isdigit() and 0 < int(choice) <= len(e.options):
            return wikipedia.page(e.options[int(choice)-1], auto_suggest=False)
        return None

def get_valid_sections(page_content):
    sections = re.findall(r'^==\s*([^=]+?)\s*==$', page_content, flags=re.MULTILINE)
    return [s.strip() for s in sections 
            if s.lower() not in ('references', 'external links', 'see also', 'notes')]

def create_presentation(main_title, sections):
    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = main_title
    
    for section in sections:
        content = wikipedia.page(main_title, auto_suggest=False).section(section)
        if not content:
            continue
            
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = section
        
        # Extract first 5 sentences, but enforce max length
        sentences = re.findall(r'[^.!?]*[.!?]', content)
        points = []
        total_chars = 0
        for sent in sentences:
            sent = sent.strip()
            if not sent:
                continue
            if len(sent) > MAX_BULLET_CHARS:
                sent = sent[:MAX_BULLET_CHARS].rstrip() + "..."
            if total_chars + len(sent) > MAX_TOTAL_CHARS:
                break
            points.append(sent)
            total_chars += len(sent)
            if len(points) == 5:
                break
        
        text_frame = slide.placeholders[1].text_frame
        text_frame.clear()
        for point in points:
            p = text_frame.add_paragraph()
            p.text = point
            p.level = 0
            p.font.size = Pt(20)
    
    filename = f"{main_title.replace(' ', '_')}_presentation.pptx"
    prs.save(filename)
    return filename

def main():
    raw_title = input("Enter Wikipedia topic: ")
    clean_title = clean_input(raw_title)
    
    page = get_wikipedia_page(clean_title)
    if not page:
        print("Error: Could not find Wikipedia page")
        return
        
    sections = get_valid_sections(page.content)
    if not sections:
        print("No valid sections found")
        return
    
    print("\nAvailable sections:")
    for i, section in enumerate(sections, 1):
        print(f"{i}. {section}")
    
    selected = input("\nEnter section numbers to include (comma-separated): ")
    try:
        indices = [int(i)-1 for i in selected.split(',') if i.strip().isdigit()]
        selected_sections = [sections[i] for i in indices if 0 <= i < len(sections)]
    except:
        print("Invalid selection")
        return
    
    if not selected_sections:
        print("No sections selected")
        return
    
    output_file = create_presentation(page.title, selected_sections)
    print(f"\nPresentation saved as {output_file}")

if __name__ == "__main__":
    main()

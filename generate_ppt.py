import wikipedia
import re
from pptx import Presentation
from pptx.util import Pt

def clean_input(user_input):
    """Remove special characters and sanitize input"""
    cleaned = re.sub(r'[^a-zA-Z0-9 ]', '', user_input)
    return re.sub(r'\s+', ' ', cleaned).strip()

def get_wikipedia_page(title):
    """Robust page retrieval with user-guided fallback"""
    try:
        # First try exact match without auto-suggest
        return wikipedia.page(title, auto_suggest=False)
    except wikipedia.exceptions.PageError:
        # Fallback to search with user selection
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
    """Extract main sections from Wikipedia content"""
    sections = re.findall(r'^==\s*([^=]+?)\s*==$', page_content, flags=re.MULTILINE)
    return [s.strip() for s in sections 
            if s.lower() not in ('references', 'external links', 'see also', 'notes')]

def create_presentation(main_title, sections):
    """Generate PowerPoint with selected sections"""
    prs = Presentation()
    
    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = main_title
    
    # Content slides
    for section in sections:
        content = wikipedia.page(main_title, auto_suggest=False).section(section)
        if not content:
            continue
            
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = section
        
        # Extract first 5 sentences
        points = re.findall(r'[^.!?]*[.!?]', content)[:5]
        
        # Format content
        text_frame = slide.placeholders[1].text_frame
        text_frame.clear()
        for point in points:
            p = text_frame.add_paragraph()
            p.text = point.strip()
            p.level = 0
            p.font.size = Pt(20)
    
    filename = f"{main_title.replace(' ', '_')}_presentation.pptx"
    prs.save(filename)
    return filename

def main():
    """Main execution flow"""
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

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
        return filtered[:7]
    except Exception as e:
        print(f"Section error: {e}")
        return []

def create_complete_ppt(main_title):
    prs = Presentation()
    
    # Title Slide
    prs.slides.add_slide(prs.slide_layouts[0]).shapes.title.text = main_title
    
    sections = get_filtered_sections(main_title)
    print(f"Processing sections: {sections}")
    
    for section in sections:
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
                # Set font size to 20
                for run in p.runs:
                    run.font.size = Pt(20)
                    
        except Exception as e:
            print(f"Error in {section}: {e}")
    
    output_file = f"{main_title.replace(' ', '_')}_presentation.pptx"
    prs.save(output_file)
    return output_file

# Run the function
create_complete_ppt("Climate Change")

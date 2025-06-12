<img src="https://r2cdn.perplexity.ai/pplx-full-logo-primary-dark%402x.png" class="logo" width="120"/>

## README.md

# Wikipedia to PowerPoint Generator

This project is an **AI-powered tool** that automatically generates a structured PowerPoint presentation from a given Wikipedia topic. It extracts the main sections from the Wikipedia article, summarizes each section into five key points, and creates a 7-slide presentation with properly formatted content. The font size for all bullet points is set to 20 for clarity and consistency.

---

## Features

- **Automatic Section Extraction:** Fetches the top-level sections (e.g., "Terminology", "Impacts") from a Wikipedia article.
- **Content Summarization:** Extracts the first five sentences from each section as bullet points.
- **PowerPoint Generation:** Creates a `.pptx` file with a title slide and up to 7 content slides, each titled with the section name.
- **Custom Formatting:** Sets the font size of all slide content to 20 for readability.
- **Easy to Use:** Just specify your main topic and run the script.

---

## Installation

1. **Clone the repository:**

```bash
git clone https://github.com/Sid-1524/text-to-ppt.git
cd wikipedia-to-ppt
```

2. **Install dependencies:**

```bash
pip install wikipedia python-pptx
```


---

## Usage

1. **Edit the script (if needed):**
    - Change the topic in the last line of the script:

```python
create_complete_ppt("Climate Change")
```

    - Replace `"Climate Change"` with your desired Wikipedia topic.
2. **Run the script:**

```bash
python your_script_name.py
```

3. **Output:**
    - A PowerPoint file named `<Your_Topic>_presentation.pptx` will be created in the project directory.

---

## Example

If you run the script with the topic `"Climate Change"`, the tool will:

- Create a title slide: **Climate Change**
- Generate up to 7 slides, each titled with a main section from the Wikipedia article (e.g., "Terminology", "Impacts").
- Add five bullet points (sentences) under each section, with font size set to 20.

---

## Project Structure

```
wikipedia-to-ppt/
├── your_script_name.py
├── README.md
└── <output_presentation>.pptx
```


---

## Notes

- The script uses the `wikipedia` Python library to fetch content and `python-pptx` to generate presentations.
- Only the first seven main sections with available content are included.
- Bullet points are extracted as the first five sentences from each section.
- For best results, use topics with well-structured Wikipedia pages.

---

## License

This project is open-source and free to use for educational and non-commercial purposes.

---

## Acknowledgments

- Built using [`wikipedia`](https://pypi.org/project/wikipedia/) and [`python-pptx`](https://python-pptx.readthedocs.io/).
- Inspired by the need to automate content creation for presentations from reliable sources[^1].

<div style="text-align: center">⁂</div>

[^1]: projects.ai_tools


<img src="https://r2cdn.perplexity.ai/pplx-full-logo-primary-dark%402x.png" class="logo" width="120"/>

# make the required changes in the readme and rewrite it

Certainly! Here’s a rewritten **README.md** reflecting the latest code and requirements:

---

# AI-Powered PowerPoint Generator

This project automatically generates a detailed, professional PowerPoint presentation on any technical topic using the GPT-4.1 API (via your GitHub Model API key).
The script creates slides with maximized, well-formatted bullet points, ensuring all content fits the canvas and is easy to read.

---

## Features

- **GPT-4.1 Content Generation:** Uses your GitHub Model API key to generate detailed, technical, and well-structured slide content.
- **No Slide Numbering:** Slide titles are clean and free from numbering or prefixes.
- **Professional Formatting:**
    - Slide titles: 28pt, bold
    - All content: 20pt, bulleted
    - Content box is shifted up by 7 pixels for optimal layout
- **Maximized Content:** Each slide fits as much detailed content as possible without overflowing the slide area.
- **No Images:** Text-only slides for clarity and focus.
- **Environment Variable Support:** Securely manage your API key with a `.env` file.

---

## Setup

### 1. Clone the Repository

```bash
git clone https://github.com/Sid-1524/Text-To-PPT.git
cd Tex-To-PPT
```


### 2. Install Dependencies

```bash
pip install python-pptx openai python-dotenv
```


### 3. Configure Your API Key

Create a `.env` file in your project directory with this line (no quotes):

```
GITHUB_TOKEN=your_github_model_api_key
```

- **Windows:**

```cmd
set GITHUB_TOKEN=your_github_model_api_key
```

- **Mac/Linux:**

```bash
export GITHUB_TOKEN=your_github_model_api_key
```


---

## Usage

Run the script:

```bash
python your_script.py
```

- Enter your desired presentation topic when prompted.
- The script will generate a PowerPoint file named `<Topic>_presentation.pptx` in the same folder.

---

## Output Example

- **Title Slide:** Topic name, 28pt, bold
- **Up to 7 Content Slides:**
    - Slide title, 28pt, bold
    - All content bulleted, 20pt
    - Content box shifted up by 7 pixels
    - No slide numbering, no images
    - Each slide contains as much detailed content as fits the canvas

---

## Notes

- The `.env` file should **not** be committed to version control. Add `.env` to your `.gitignore`.
- The script uses the [python-pptx](https://python-pptx.readthedocs.io/), [openai](https://pypi.org/project/openai/), and [python-dotenv](https://pypi.org/project/python-dotenv/) libraries.
- Make sure your GitHub Model API key has access to the GPT-4.1 model endpoint.

---

## License

This project is open-source and free to use for educational and non-commercial purposes.

---

## Acknowledgments

- Built with [python-pptx](https://python-pptx.readthedocs.io/), [OpenAI Python SDK](https://pypi.org/project/openai/), and [python-dotenv](https://pypi.org/project/python-dotenv/).
- GPT-4.1 content powered by your GitHub-hosted model endpoint.

---

**Enjoy creating high-quality presentations with just one prompt!**


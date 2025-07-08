# QuickSlide 2 – AI Presentation Generator

QuickSlide 2 is an AI-powered application that generates professional PowerPoint presentations from text input, voice recordings, or uploaded documents. It uses large language models to structure content and output clean, readable slide decks — fast.

---

## Features

- **Multi-modal Input**  
  Generate presentations from:
  - Text descriptions
  - Voice recordings
  - Documents (`.txt`, `.pdf`, `.docx`, `.csv`, `.xlsx`)

- **AI-Powered Content Generation**  
  Uses Mistral AI to structure ideas into sections and bullet points

- **Automated Slide Design**  
  Slides are generated using `python-pptx` with clean layouts and formatting

- **Customizable Output**  
  - Select themes
  - Control slide count
  - Toggle detail level

- **Voice-to-Text**  
  Record and transcribe presentation ideas using Google's speech recognition

- **Document Analysis**  
  Automatically extracts relevant content from uploaded files to integrate into the presentation

---

## Installation

### Prerequisites

- Python 3.8+
- pip (Python package manager)

### Setup

Clone the repository:
```bash
git clone https://github.com/CodexAbhi/QuickSlide2
cd QuickSlide2
````

Install the required packages:

```bash
pip install -r requirements.txt
```

Create a `.env` file in the root directory with your API keys:

```
MISTRAL_API_KEY=your_mistral_api_key
OPENAI_API_KEY=your_openai_api_key
```

---

## Usage

Start the Streamlit app:

```bash
streamlit run app.py
```

Access the app in your browser at: `http://localhost:8501`

### Creating a Presentation

1. Enter your topic or brief description
2. (Optional) Record a voice input or upload a reference document
3. Configure:

   * Detail level
   * Theme
   * Approximate slide count
4. Click "Generate Presentation"
5. Download the resulting `.pptx` file

---

## Project Structure

```
QuickSlide2/
├── app.py                 # Streamlit frontend
├── ppt_generator.py       # Slide creation logic
├── mistral_client.py      # Mistral API interface
├── requirements.txt       # Dependency list
└── .env                   # API keys (not included in repo)
```

---

## Dependencies

* streamlit
* python-pptx
* openai
* python-dotenv
* docx2txt
* PyPDF2
* pandas
* audio-recorder-streamlit
* SpeechRecognition

Install with:

```bash
pip install -r requirements.txt
```

---

## Future Improvements

* Image generation support (DALL·E, Stable Diffusion)
* Custom template uploads
* Real-time slide preview
* Google Slides or PDF export options
* Collaborative editing features

---

## Acknowledgments

* Mistral AI – content generation API
* OpenAI – language model integration
* Google Speech Recognition – voice input handling
* Streamlit – web app framework
* python-pptx – PowerPoint file creation

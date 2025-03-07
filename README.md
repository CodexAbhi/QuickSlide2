# AI Presentation Generator

An application that leverages AI to automatically generate professional PowerPoint presentations from text descriptions, voice input, and reference documents.

![AI Presentation Generator]([https://github.com/CodexAbhi/QuickSlide2/images/cover.jpeg](https://github.com/CodexAbhi/QuickSlide2/blob/main/images/cover.jpeg))

## Features

- **Multi-modal Input**: Create presentations from text descriptions, voice recordings, or uploaded documents
- **AI-powered Content Generation**: Uses Mistral AI to generate structured, coherent presentation content
- **Professional Slide Design**: Automated creation of visually appealing slides with proper formatting and layout
- **Multiple Themes**: Choose from different presentation styles
- **Voice-to-Text**: Record and transcribe your presentation ideas using Google's speech recognition
- **Document Analysis**: Extract content from TXT, PDF, DOCX, CSV, and XLSX files to incorporate into presentations
- **Customizable Slide Count**: Control the approximate number of slides in your presentation

## Installation

### Prerequisites

- Python 3.8+
- pip package manager

### Setup

1. Clone this repository:

   ```bash
   git clone https://github.com/yourusername/ai-presentation-generator.git
   cd ai-presentation-generator
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Create a `.env` file in the project root with your API keys:

   ```
   MISTRAL_API_KEY=your_mistral_api_key
   OPENAI_API_KEY=your_openai_api_key
   ```

   You can obtain these API keys from:

   - [Mistral AI](https://console.mistral.ai/)
   - [OpenAI](https://platform.openai.com/)

## Usage

### Starting the Application

Run the Streamlit app:

```bash
streamlit run app.py
```

The application will be available at `http://localhost:8501` in your web browser.

### Creating a Presentation

1. Enter your presentation topic or description in the text area
2. Optionally:
   - Record a voice description
   - Upload a reference document (TXT, PDF, DOCX, CSV, XLSX)
3. Configure presentation options:
   - Select whether to generate detailed content
   - Choose a presentation theme
   - Set the approximate slide count
4. Click "Generate Presentation"
5. Download the PPTX file when it's ready

## Project Structure

- `app.py`: Main Streamlit application
- `mistral_client.py`: Client for interacting with Mistral AI API
- `ppt_generator.py`: PowerPoint generation using python-pptx
- `requirements.txt`: Project dependencies

## Key Components

### PPT Generator

The PowerPoint generator takes structured content from the AI and creates visually appealing slides with proper formatting, including:

- Title slides
- Section headers
- Content slides with bullet points
- Closing slides

### Mistral AI Integration

The application uses Mistral AI's large language model to:

- Parse user input for presentation requirements
- Structure content into logical sections
- Generate detailed bullet points
- Create a cohesive presentation flow

### Voice Input & Document Analysis

- Voice recordings are transcribed using Google's speech recognition
- Document analysis extracts content from various file formats to enhance presentations

## Requirements

```
streamlit
python-pptx
openai
python-dotenv
docx2txt
PyPDF2
pandas
audio-recorder-streamlit
SpeechRecognition
```

## Future Improvements

- Add image generation capabilities using DALL-E or Stable Diffusion
- Implement custom template uploads
- Add collaborative editing features
- Support for more presentation formats (beyond PPTX)
- Enhance the UI with preview capabilities

## Acknowledgments

- Mistral AI for the content generation API
- Google for speech recognition capabilities
- Streamlit for the web application framework
- Python-PPTX for PowerPoint generation capabilities

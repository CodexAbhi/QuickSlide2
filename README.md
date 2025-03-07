# AI Presentation Generator

A Streamlit application that generates PowerPoint presentations from text prompts using Mistral AI.

## Features

- Generate detailed or concise presentation content from simple prompts
- Automatic organization of content into slides with appropriate sections
- Export to PowerPoint (.pptx) format
- Clean and user-friendly interface

## Setup

1. Clone this repository:
```
git clone https://github.com/yourusername/ai-presentation-generator.git
cd ai-presentation-generator
```

2. Install the required dependencies:
```
pip install -r requirements.txt
```

3. Create a `.env` file in the project root with your Mistral API key:
```
MISTRAL_API_KEY=your_mistral_api_key_here
```

4. Run the application:
```
streamlit run app.py
```

## Usage

1. Enter a topic or detailed prompt in the text area
2. Toggle the "Generate detailed content" checkbox based on your needs
3. Click "Generate Presentation" 
4. Review the generated content and download the PowerPoint file

## Project Structure

- `app.py`: Main Streamlit application
- `mistral_client.py`: Handles interactions with the Mistral AI API
- `ppt_generator.py`: PowerPoint generation logic
- `.env`: Stores your API key (not tracked in git)

## Future Enhancements

- Additional PowerPoint themes
- Custom slide templates
- Image generation for slides
- Saving and loading presentation drafts
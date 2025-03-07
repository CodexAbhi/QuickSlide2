import streamlit as st
import os
import tempfile
import base64
import json
from mistral_client import MistralClient
from ppt_generator import PPTGenerator
import openai
from dotenv import load_dotenv
import io
import docx2txt
import PyPDF2
import pandas as pd
from audio_recorder_streamlit import audio_recorder
import time
import speech_recognition as sr

# Load environment variables
load_dotenv()

# Set OpenAI API key and initialize client
openai_api_key = os.getenv("OPENAI_API_KEY")
if openai_api_key:
    # Initialize with new client API for OpenAI >= 1.0.0
    openai_client = openai.OpenAI(api_key=openai_api_key)

# Set page config
st.set_page_config(
    page_title="AI Presentation Generator",
    page_icon="📊",
    layout="wide"
)

# Initialize session state for storing generated content
if 'ppt_content' not in st.session_state:
    st.session_state.ppt_content = None
if 'download_ready' not in st.session_state:
    st.session_state.download_ready = False
if 'temp_file_path' not in st.session_state:
    st.session_state.temp_file_path = None
if 'speech_text' not in st.session_state:
    st.session_state.speech_text = ""
if 'file_text' not in st.session_state:
    st.session_state.file_text = ""
if 'is_recording' not in st.session_state:
    st.session_state.is_recording = False
    
# Function to transcribe speech using OpenAI Whisper - Updated for OpenAI >= 1.0.0
def transcribe_audio(audio_bytes):
    try:
        # Create a recognizer instance
        recognizer = sr.Recognizer()
        
        # Save audio bytes to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.wav') as tmp_file:
            tmp_file.write(audio_bytes)
            tmp_file_path = tmp_file.name
        
        # Load the audio file and transcribe it
        with sr.AudioFile(tmp_file_path) as source:
            audio_data = recognizer.record(source)
            
            # Use Google's free speech recognition
            # You can also try other engines like recognizer.recognize_sphinx() which is offline
            text = recognizer.recognize_google(audio_data)
            
        # Clean up the temporary file
        os.unlink(tmp_file_path)
        
        return text
        
    except sr.UnknownValueError:
        return "Speech recognition could not understand the audio"
    except sr.RequestError as e:
        return f"Error with speech recognition service: {e}"
    except Exception as e:
        st.error(f"Error transcribing audio: {str(e)}")
        return f"Error: {str(e)}"

# Function to extract text from uploaded files with improved error handling
def extract_text_from_file(uploaded_file):
    text = ""
    file_extension = os.path.splitext(uploaded_file.name)[1].lower()
    
    try:
        # Handle different file types
        if file_extension == '.txt':
            text = uploaded_file.getvalue().decode('utf-8')
        
        elif file_extension == '.docx':
            try:
                text = docx2txt.process(io.BytesIO(uploaded_file.getvalue()))
            except Exception as e:
                return f"Error processing DOCX file: {str(e)}. Make sure it's a valid Word document."
        
        elif file_extension == '.pdf':
            try:
                pdf_reader = PyPDF2.PdfReader(io.BytesIO(uploaded_file.getvalue()))
                for page_num in range(len(pdf_reader.pages)):
                    text += pdf_reader.pages[page_num].extract_text() + "\n"
                
                # Check if we got any text
                if not text.strip():
                    return "The PDF appears to contain scanned images rather than text. Cannot extract content."
            except Exception as e:
                return f"Error processing PDF file: {str(e)}. Make sure it's a valid PDF document."
        
        elif file_extension in ['.csv', '.xlsx', '.xls']:
            try:
                if file_extension == '.csv':
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                # Check if dataframe is empty
                if df.empty:
                    return "The uploaded file appears to be empty."
                
                # Convert the dataframe to a text summary
                text = f"File summary: {uploaded_file.name}\n\n"
                text += f"Columns: {', '.join(df.columns.tolist())}\n"
                text += f"Rows: {len(df)}\n\n"
                text += "Sample data (first 5 rows):\n"
                text += df.head().to_string() + "\n\n"
                text += "Statistical summary:\n"
                
                # Add basic statistics for numerical columns
                numeric_cols = df.select_dtypes(include=['number']).columns
                if len(numeric_cols) > 0:
                    text += df[numeric_cols].describe().to_string()
            except Exception as e:
                return f"Error processing spreadsheet: {str(e)}. Make sure it's a valid file."
        
        else:
            text = f"Unsupported file type: {file_extension}. Please upload a .txt, .docx, .pdf, .csv, or .xlsx file."
    
    except Exception as e:
        text = f"Error processing file: {str(e)}"
    
    # Truncate very large files to prevent issues
    if len(text) > 10000:
        text = text[:10000] + "\n\n... (content truncated for length)"
    
    return text

# Function to download the generated presentation
def get_download_link(file_path, file_name):
    with open(file_path, "rb") as file:
        contents = file.read()
    b64 = base64.b64encode(contents).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{file_name}" class="download-button">Download Presentation</a>'
    return href

# Add some custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem !important;
        color: #0072C6;
    }
    .sub-header {
        font-size: 1.5rem !important;
        margin-bottom: 1rem;
    }
    .download-button {
        display: inline-block;
        padding: 0.5em 1em;
        background-color: #0072C6;
        color: white !important;
        text-decoration: none;
        font-weight: bold;
        border-radius: 4px;
        text-align: center;
        transition: background-color 0.3s;
    }
    .download-button:hover {
        background-color: #005999;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 15px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        border-radius: 4px 4px 0 0;
    }
</style>
""", unsafe_allow_html=True)

# Title and description
st.markdown('<p class="main-header">AI Presentation Generator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Generate professional PowerPoint presentations from simple prompts using AI.</p>', unsafe_allow_html=True)

# Main prompt input - Always visible
with st.container():
    prompt = st.text_area(
        "Enter your presentation topic or prompt:",
        height=150,
        placeholder="Example: 'The impact of artificial intelligence on healthcare in the next decade'"
    )

# Additional input options
with st.container():
    st.subheader("Additional Input Methods (Optional)")
    
    # Input method selection
    input_tabs = st.tabs(["Speech Input", "File Upload"])
    
    # Speech input
    with input_tabs[0]:
        st.markdown("### 🎙️ Record your presentation idea")
        
        col1, col2 = st.columns([3, 1])
        with col1:
            # Improved audio recorder with better parameters
            audio_bytes = audio_recorder(
                text="Click the microphone to start/stop recording",
                recording_color="#e8585c", 
                neutral_color="#0072C6",
                energy_threshold=(-1.0, 1.0),
                pause_threshold=300.0,  # Increased to 5 minutes
                sample_rate=44100
            )
            
        with col2:
            if st.button("Clear Recording"):
                st.session_state.speech_text = ""
                # Need to reset the audio_bytes variable somehow
                # This is challenging because the audio_recorder component doesn't have a clear method
                # A workaround is to refresh the page
                st.rerun()
        
        if audio_bytes:
            st.audio(audio_bytes, format="audio/wav")
            
            # Transcribe button with improved error handling
            # Modify the transcribe button code to remove OpenAI API key check
            if st.button("Transcribe Audio"):
                with st.spinner("Transcribing..."):
                    transcribed_text = transcribe_audio(audio_bytes)
                    if transcribed_text and not transcribed_text.startswith("Error:") and not transcribed_text.startswith("Speech recognition could not understand"):
                        st.session_state.speech_text = transcribed_text
                        st.success("Transcription complete!")
                    else:
                        st.error(transcribed_text or "Failed to transcribe audio. Please try again.")
        
        # Display transcribed text with better formatting
        if st.session_state.speech_text:
            st.markdown("##### 📝 Transcribed Text:")
            st.markdown(f"""
            <div style='background-color:#f0f2f6;padding:10px;border-radius:5px;'>
                {st.session_state.speech_text}
            </div>
            """, unsafe_allow_html=True)
            st.info("This text will be combined with your main prompt.")
    
    # File upload
    with input_tabs[1]:
        st.markdown("### 📄 Upload a reference document")
        
        uploaded_file = st.file_uploader(
            "Upload a file to use as reference for your presentation",
            type=["txt", "pdf", "docx", "csv", "xlsx", "xls"],
            help="The content of this file will be analyzed and used to enhance your presentation."
        )
        
        if uploaded_file is not None:
            # Process file with progress indicator
            with st.spinner(f"Processing {uploaded_file.name}..."):
                extracted_text = extract_text_from_file(uploaded_file)
                
                # Check if we got an error message
                if extracted_text.startswith("Error"):
                    st.error(extracted_text)
                else:
                    st.session_state.file_text = extracted_text
                    
                    # Show success with file details
                    file_size = len(uploaded_file.getvalue()) / 1024  # Size in KB
                    st.success(f"File '{uploaded_file.name}' ({file_size:.1f} KB) successfully processed")
                    
                    # Show preview with expandable section
                    with st.expander("View extracted content", expanded=False):
                        if len(extracted_text) > 1000:
                            preview = extracted_text[:1000] + "... (content truncated for preview)"
                            st.text_area("File content preview", preview, height=200)
                        else:
                            st.text_area("File content", extracted_text, height=200)
                    
                    # Explain how it will be used
                    with st.expander("How this will enhance your presentation", expanded=True):
                        st.markdown("The AI will analyze this content and incorporate relevant information into your presentation.")
                        st.markdown("**Examples of what the AI might extract:**")
                        st.markdown("- Key statistics and figures")
                        st.markdown("- Important concepts and terminology")
                        st.markdown("- Structure and organization of ideas")
                        st.markdown("- Supporting evidence for your main points")
    
    # Presentation options
    st.subheader("Presentation Options")
    
    col1, col2 = st.columns(2)
    with col1:
        detailed = st.checkbox("Generate detailed content", value=True)
    with col2:
        theme = st.selectbox(
            "Select presentation theme:",
            ["modern_blue", "elegant_dark", "vibrant", "minimal"],
            index=0
        )
    
    # Generate button
    if st.button("Generate Presentation", type="primary"):
        # Validate that we have some prompt content
        if not prompt.strip():
            st.error("Please enter a presentation topic or prompt to continue.")
        else:
            # Combine all inputs to create an enhanced prompt
            full_prompt = prompt
            
            # Add speech input if available
            if st.session_state.speech_text:
                full_prompt += f"\n\nAdditional context from speech: {st.session_state.speech_text}"
            
            # Add file content if available - handle large content better
            if st.session_state.file_text:
                # Truncate very large file content to prevent token limits
                file_text = st.session_state.file_text
                if len(file_text) > 4000:  # Arbitrary limit to prevent token overflows
                    file_text = file_text[:4000] + "... (content truncated for length)"
                
                full_prompt += f"\n\nIncorporate the following reference material: {file_text}"
            
            # Generate the presentation
            with st.spinner("Creating your presentation..."):
                try:
                    # Initialize Mistral client
                    client = MistralClient()
                    
                    # Generate content
                    st.session_state.ppt_content = client.generate_content(full_prompt, detailed)
                    
                    if "error" in st.session_state.ppt_content:
                        st.error(f"Error: {st.session_state.ppt_content['error']}")
                    else:
                        # Generate PPT with selected theme
                        ppt_gen = PPTGenerator(theme=theme)
                        ppt = ppt_gen.generate_from_content(st.session_state.ppt_content)
                        
                        # Save to temporary file
                        temp_dir = tempfile.mkdtemp()
                        # Use a safe version of the prompt for the filename
                        safe_name = ''.join(c if c.isalnum() else '_' for c in prompt[:20]).strip('_')
                        file_name = f"presentation_{safe_name}.pptx"
                        file_path = os.path.join(temp_dir, file_name)
                        ppt_gen.save(file_path)
                        
                        st.session_state.temp_file_path = file_path
                        st.session_state.download_ready = True
                        
                        # Show presentation ready message
                        st.success("✅ Your presentation is ready to download!")
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    st.error("If this is an API error, please check that your Mistral API key is configured correctly in the .env file.")

# Display only the download link when ready
if st.session_state.download_ready and st.session_state.temp_file_path:
    # Display download button
    st.markdown(get_download_link(st.session_state.temp_file_path, 
                                  os.path.basename(st.session_state.temp_file_path)), 
                unsafe_allow_html=True)

# Add some information at the bottom
st.markdown("---")
st.markdown("This app uses Mistral AI to generate presentation content, OpenAI Whisper for speech transcription, and Python-PPTX to create PowerPoint files.")
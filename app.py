import streamlit as st
import os
import tempfile
import base64
from mistral_client import MistralClient
from ppt_generator import PPTGenerator

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

# Function to download the generated presentation
def get_download_link(file_path, file_name):
    with open(file_path, "rb") as file:
        contents = file.read()
    b64 = base64.b64encode(contents).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{file_name}">Download Presentation</a>'
    return href

# Title and description
st.title("AI Presentation Generator")
st.markdown("Generate professional PowerPoint presentations from simple prompts using AI.")

# Input area
with st.container():
    prompt = st.text_area(
        "Enter your presentation topic or prompt:",
        height=150,
        placeholder="Example: 'The impact of artificial intelligence on healthcare in the next decade'"
    )
    
    col1, col2 = st.columns(2)
    with col1:
        detailed = st.checkbox("Generate detailed content", value=True)
    with col2:
        theme = st.selectbox(
            "Select presentation theme:",
            ["modern_blue", "elegant_dark", "vibrant", "minimal"],
            index=0
        )
    
    if st.button("Generate Presentation"):
        if prompt:
            with st.spinner("Generating your presentation content..."):
                try:
                    # Initialize Mistral client
                    client = MistralClient()
                    
                    # Generate content
                    st.session_state.ppt_content = client.generate_content(prompt, detailed)
                    
                    if "error" in st.session_state.ppt_content:
                        st.error(f"Error: {st.session_state.ppt_content['error']}")
                    else:
                        # Generate PPT with selected theme
                        with st.spinner("Creating PowerPoint presentation..."):
                            ppt_gen = PPTGenerator(theme=theme)
                            ppt = ppt_gen.generate_from_content(st.session_state.ppt_content)
                            
                            # Save to temporary file
                            temp_dir = tempfile.mkdtemp()
                            file_name = f"presentation_{prompt[:20].replace(' ', '_')}.pptx"
                            file_path = os.path.join(temp_dir, file_name)
                            ppt_gen.save(file_path)
                            
                            st.session_state.temp_file_path = file_path
                            st.session_state.download_ready = True
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
        else:
            st.warning("Please enter a prompt to generate content.")

# Display the generated content and download link
if st.session_state.ppt_content and "error" not in st.session_state.ppt_content:
    st.success("Presentation generated successfully!")
    
    with st.expander("View Presentation Content", expanded=True):
        st.subheader(st.session_state.ppt_content.get("title", "Presentation"))
        if "subtitle" in st.session_state.ppt_content and st.session_state.ppt_content["subtitle"]:
            st.caption(st.session_state.ppt_content["subtitle"])
        
        for i, section in enumerate(st.session_state.ppt_content.get("sections", [])):
            st.markdown(f"### {section.get('title', f'Section {i+1}')}")
            for point in section.get("content", []):
                st.markdown(f"- {point}")
    
    if st.session_state.download_ready and st.session_state.temp_file_path:
        st.markdown(get_download_link(st.session_state.temp_file_path, 
                                      os.path.basename(st.session_state.temp_file_path)), 
                    unsafe_allow_html=True)
        
        st.info("You can also modify the content above, then regenerate the presentation.")

        # Show theme preview
        st.subheader("Theme Preview")
        theme_preview = {
            "modern_blue": "https://via.placeholder.com/800x200/0072C6/FFFFFF?text=Modern+Blue+Theme",
            "elegant_dark": "https://via.placeholder.com/800x200/282828/FFC300?text=Elegant+Dark+Theme",
            "vibrant": "https://via.placeholder.com/800x200/D50052/FFFFFF?text=Vibrant+Theme",
            "minimal": "https://via.placeholder.com/800x200/464646/FF674D?text=Minimal+Theme"
        }
        st.image(theme_preview.get(theme, theme_preview["modern_blue"]))

# Add some information at the bottom
st.markdown("---")
st.markdown("This app uses Mistral AI to generate presentation content and Python-PPTX to create PowerPoint files.")
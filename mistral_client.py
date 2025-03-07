import os
import requests
import json
from dotenv import load_dotenv
import re

# Load API key from .env file
load_dotenv()

class MistralClient:
    def __init__(self):
        # Get API key from environment variables
        self.api_key = os.getenv("MISTRAL_API_KEY")
        if not self.api_key:
            raise ValueError("MISTRAL_API_KEY not found in environment variables")
        
        self.base_url = "https://api.mistral.ai/v1"
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
    def extract_presentation_instructions(self,text):
        """
        Extract specific presentation instructions from text.
        
        Args:
            text (str): The text content from uploaded file
            
        Returns:
            dict: Instructions for presentation generation
        """
        instructions = {
            "slide_instructions": [],
            "general_instructions": []
        }
        
        # Look for specific slide instructions
        slide_pattern = r"(leave|make|create|keep)\s+(the\s+)?(\w+)\s+(slide)\s+(blank|empty)"
        slide_matches = re.finditer(slide_pattern, text, re.IGNORECASE)
        
        for match in slide_matches:
            slide_indicator = match.group(3)
            action = match.group(5)
            
            # Handle both numeric and text indicators
            if slide_indicator.isdigit():
                # Convert to 0-indexed for internal use (1 becomes 0)
                slide_num = int(slide_indicator) - 1
                instructions["slide_instructions"].append({
                    "slide_index": slide_num,  # 0-based index
                    "action": action
                })
            elif slide_indicator.lower() in ["first", "second", "third", "fourth", "fifth"]:
                # Convert word to 0-indexed number
                num_map = {"first": 0, "second": 1, "third": 2, "fourth": 3, "fifth": 4}
                instructions["slide_instructions"].append({
                    "slide_index": num_map.get(slide_indicator.lower(), 0),
                    "action": action
                })
        
        return instructions
    
    def generate_content(self, prompt, detailed=True):
        """
        Generate content using Mistral AI based on the prompt.

        Args:
            prompt (str): The user's input prompt
            detailed (bool): Whether to generate detailed content

        Returns:
            dict: Generated content in structured format
        """
        # Extract instructions from the prompt if it contains file content
        file_instructions = {}
        if "Incorporate the following reference material:" in prompt:
            # Extract file content
            file_content_start = prompt.find("Incorporate the following reference material:") + len("Incorporate the following reference material:")
            file_content = prompt[file_content_start:].strip()
            
            # Extract instructions from file content
            file_instructions = self.extract_presentation_instructions(file_content)
            
            # Enhance prompt with extracted instructions
            for instr in file_instructions.get("general_instructions", []):
                prompt += f"\n\nPlease follow this specific instruction: {instr}"
            
            # Add slide instructions in a structured format
            if file_instructions.get("slide_instructions"):
                prompt += "\n\nSpecific slide instructions:"
                for instr in file_instructions.get("slide_instructions", []):
                    prompt += f"\n- Make slide {instr['slide_number']} {instr['action']}"

        # Rest of your generate_content method remains the same...
        # Include the extracted instructions in your system prompt
        system_prompt = """
        You are an expert presentation content creator specializing in insightful and structured AI-driven presentations.
        
        **Instructions:**
        - Create a **comprehensive, well-structured** presentation based on the user's prompt.
        - If specific slide instructions are provided (like 'leave slide 3 blank'), you MUST follow them exactly.
        - Ensure the presentation has a **logical flow** from past to future impacts.
        - Include **real-world examples, case studies, and statistics** where relevant.
        - Balance **technical depth** while keeping it engaging for a general audience.
        - Use **rich text formatting** in your content points:
            - Use **double asterisks** for important terms or concepts that should be bold
            - Use *single asterisks* for terms that should be italic
        
        **Format Requirements:**
        Your response should be a **JSON object** with the following structure:

        {
            "title": "Presentation Title",
            "subtitle": "Optional Subtitle",
            "sections": [
                {
                    "title": "Section Title",
                    "content": ["Point 1 with **bold** and *italic* text", "Point 2", "Point 3"]
                }
            ],
            "call_to_action": "Key takeaways and next steps",
            "special_instructions": []
        }
        
        **Important**: The text formatting (bold, italic) in content will be preserved in the final presentation.
        If asked to leave certain slides blank, add a special instruction in the "special_instructions" array.
        """
        
        # Call Mistral API
        try:
            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers=self.headers,
                json={
                    "model": "mistral-large-latest",
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": prompt}
                    ],
                    "temperature": 0.7,
                    "response_format": {"type": "json_object"}
                }
            )
            
            response.raise_for_status()
            result = response.json()
            
            # Extract the JSON content from the response
            try:
                content = result["choices"][0]["message"]["content"]
                return json.loads(content)
            except (KeyError, json.JSONDecodeError) as e:
                return {"error": f"Failed to parse response: {str(e)}"}
                
        except requests.exceptions.RequestException as e:
            return {"error": f"API request failed: {str(e)}"}
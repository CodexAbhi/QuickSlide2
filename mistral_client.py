#mistral_client.py
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
    
    def extract_presentation_instructions(self, text):
        """
        Extract any presentation instructions from the input text.
        
        Args:
            text (str): Input text to analyze for presentation instructions
            
        Returns:
            dict: Dictionary with extracted instructions
        """
        instructions = {
            "general_instructions": [],
            "slide_instructions": []
        }
        
        # Look for general instructions
        general_patterns = [
            r"please (make|create|design) (a|the) presentation (that|which) (.*?)[\.!\?]",
            r"the presentation should (.*?)[\.!\?]",
            r"make sure (to|that) (.*?)[\.!\?]"
        ]
        
        for pattern in general_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                instruction = match.group(0)
                if instruction and len(instruction) > 10:  # Minimal length check
                    instructions["general_instructions"].append(instruction)
        
        # Look for specific slide instructions
        slide_patterns = [
            r"(slide|page) (\d+) should (.*?)[\.!\?]",
            r"(on|in|for) (slide|page) (\d+)[,]? (.*?)[\.!\?]",
            r"(leave|make) (slide|page) (\d+) (blank|empty)"
        ]
        
        for pattern in slide_patterns:
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                groups = match.groups()
                # Extract slide number
                if "leave" in groups or "make" in groups:
                    # Format: "leave slide 3 blank"
                    slide_num = int(groups[2])
                    action = groups[3]  # blank or empty
                else:
                    # Other formats
                    slide_num_idx = 1 if groups[0].lower() in ["slide", "page"] else 2
                    slide_num = int(groups[slide_num_idx])
                    action = groups[-1]
                
                instructions["slide_instructions"].append({
                    "slide_number": slide_num,
                    "action": action
                })
        
        return instructions
    
    def generate_content(self, prompt, detailed=True):
        """
        Generate content using Mistral AI based on the prompt.

        Args:
            prompt (str): The user's comprehensive input prompt
            detailed (bool): Whether to generate detailed content

        Returns:
            dict: Generated content in structured format
        """
        # Extract instructions from the entire prompt
        instructions = self.extract_presentation_instructions(prompt)
        
        # Extract slide count from the prompt
        slide_count_match = re.search(r'Target exactly (\d+) slides total', prompt)
        target_slides = int(slide_count_match.group(1)) if slide_count_match else 15
        
        # Enhance prompt with extracted instructions as needed
        enhanced_prompt = prompt
        for instr in instructions.get("general_instructions", []):
            if instr not in enhanced_prompt:
                enhanced_prompt += f"\n\nPlease follow this specific instruction: {instr}"
        
        # Add slide instructions in a structured format if they're not already in the prompt
        if instructions.get("slide_instructions"):
            if "Specific slide instructions:" not in enhanced_prompt:
                enhanced_prompt += "\n\nSpecific slide instructions:"
                for instr in instructions.get("slide_instructions", []):
                    enhanced_prompt += f"\n- Make slide {instr['slide_number']} {instr['action']}"

        # Set detail level based on user preference
        detail_level = "highly detailed and comprehensive" if detailed else "concise and focused"
        
        # System prompt with updated instructions
        system_prompt = f"""
        You are an expert presentation content creator specializing in insightful and {detail_level} AI-driven presentations.
        
        **Instructions:**
        - Create a **well-structured** presentation based on the user's input.
        - CRITICALLY IMPORTANT: Generate EXACTLY {target_slides} slides total, including title and closing slides.
        - You MUST output content that will result in EXACTLY {target_slides} slides when processed.
        - If specific slide instructions are provided (like 'leave slide 3 blank'), you MUST follow them exactly.
        - Ensure the presentation has a **logical flow** with clear progression between topics.
        - Include **real-world examples, case studies, and statistics** where relevant.
        - Balance **technical depth** while keeping it engaging for a general audience.
        - Use **rich text formatting** in your content points:
            - Use **double asterisks** for important terms or concepts that should be bold
            - Use *single asterisks* for terms that should be italic
        
        **Format Requirements:**
        Your response should be a **JSON object** with the following structure:

        {{
            "title": "Presentation Title",
            "subtitle": "Optional Subtitle",
            "target_slides": {target_slides},
            "sections": [
                {{
                    "title": "Section Title",
                    "content": ["Point 1 with **bold** and *italic* text", "Point 2", "Point 3"]
                }}
            ],
            "call_to_action": "Key takeaways and next steps",
            "special_instructions": []
        }}
        
        **Important Notes for Achieving {target_slides} Slides:** 
        - The presentation will ALWAYS include title and closing slides (2 slides total).
        - Each major section (before the colon in section titles) gets a section header slide.
        - Content slides have a maximum of 7 bullet points each.
        - If a section has more bullet points, it will be split across multiple slides.
        - To ensure you hit exactly {target_slides} slides:
        1. Carefully plan your major sections (each adds 1 slide)
        2. Adjust the number of content bullet points to achieve the right slide count
        3. Include the exact "target_slides" value of {target_slides} in your JSON response
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
                        {"role": "user", "content": enhanced_prompt}
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
                data = json.loads(content)
                
                # Ensure target_slides is included
                if "target_slides" not in data:
                    data["target_slides"] = target_slides
                    
                return data
            except (KeyError, json.JSONDecodeError) as e:
                return {"error": f"Failed to parse response: {str(e)}"}
                
        except requests.exceptions.RequestException as e:
            return {"error": f"API request failed: {str(e)}"}
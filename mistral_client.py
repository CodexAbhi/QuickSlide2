import os
import requests
from dotenv import load_dotenv

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
    
    def generate_content(self, prompt, detailed=True):
        """
        Generate content using Mistral AI based on the prompt.

        Args:
            prompt (str): The user's input prompt
            detailed (bool): Whether to generate detailed content

        Returns:
            dict: Generated content in structured format
        """

        # Enhance the prompt based on detail level
        if detailed:
            system_prompt = """
            You are an expert presentation content creator specializing in insightful and structured AI-driven presentations.
            
            **Instructions:**
            - Create a **comprehensive, well-structured** presentation based on the user's prompt.
            - Ensure the presentation has a **logical flow** from past to future impacts.
            - Include **real-world examples, case studies, and statistics** where relevant.
            - Balance **technical depth** while keeping it engaging for a general audience.

            **Format Requirements:**
            Your response should be a **JSON object** with the following structure:

            {
                "title": "Presentation Title",
                "subtitle": "Optional Subtitle",
                "sections": [
                    {
                        "title": "Section Title",
                        "content": ["Point 1", "Point 2", "Point 3"]
                    }
                ],
                "call_to_action": "Key takeaways and next steps"
            }

            **Presentation Structure:**
            1. **Title Slide**: A compelling title.
            2. **Introduction**: A strong hook, why the topic matters, key objectives.
            3. **Historical Context & Current State**: How the topic evolved, key milestones.
            4. **Future Predictions**: How AI will impact industries, economy, society.
            5. **Case Studies & Examples**: Real-world applications, companies, research.
            6. **Challenges & Ethical Considerations**: Bias, job impact, privacy, AI governance.
            7. **Solutions & Call to Action**: Steps individuals, companies, and governments should take.
            8. **Conclusion**: Summary of key points, an inspiring closing statement.

            Expand the user's prompt with relevant insights, examples, and future trends.
            """

        else:
            system_prompt = """
            Create a concise **but impactful** presentation outline based on the user's prompt.

            **Format Requirements:**
            Your response should be a **JSON object** with the following structure:

            {
                "title": "Presentation Title",
                "subtitle": "Optional Subtitle",
                "sections": [
                    {
                        "title": "Section Title",
                        "content": ["Point 1", "Point 2"]
                    }
                ],
                "call_to_action": "Key takeaways and next steps"
            }

            **Presentation Structure:**
            1. **Title Slide**
            2. **Introduction** (Short, with a strong hook)
            3. **3-5 Key Sections** (Each with 2-3 bullet points)
            4. **Conclusion** (Summarize and provide a next step)

            Ensure the content remains **insightful, engaging, and logically structured.**
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
                import json
                return json.loads(content)
            except (KeyError, json.JSONDecodeError) as e:
                return {"error": f"Failed to parse response: {str(e)}"}
                
        except requests.exceptions.RequestException as e:
            return {"error": f"API request failed: {str(e)}"}
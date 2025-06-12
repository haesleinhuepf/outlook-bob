import os
from openai import OpenAI
from pathlib import Path
from dotenv import load_dotenv
import argparse

class EmailAssistant:
    def __init__(self):
        # Load environment variables
        load_dotenv()
        
        # Initialize OpenAI client for Ollama
        self.openai_client = OpenAI(
            base_url=os.getenv('OLLAMA_BASE_URL', 'http://localhost:11434/v1'),
            api_key=os.getenv('OLLAMA_API_KEY', 'not-needed')
        )
        
        # Get the system prompt from file
        script_dir = Path(__file__).parent.absolute()
        prompt_file = script_dir / "system_prompt.txt"
        try:
            with open(prompt_file, 'r') as file:
                self.system_prompt = file.read().strip()
        except FileNotFoundError:
            print(f"Warning: system_prompt.txt not found. Using default prompt.")
            self.system_prompt = 'You are a professional email assistant.'

    def read_email(self, input_file_path):
        """Read email content from a text file"""
        try:
            with open(input_file_path, 'r') as file:
                return file.read()
        except Exception as e:
            print(f"Error reading email file: {str(e)}")
            return None

    def generate_response(self, email_content, response_type):
        """Generate email response using Ollama"""
        if not email_content:
            return "Error: No email content provided."

        # Create prompt for AI analysis
        prompt = f"""Please draft a professional email response to the following email. The response should be {response_type}:

{email_content}

Please provide a response that is:
1. Professional and courteous
2. Addresses all points in the original email
3. Maintains appropriate tone and formality
4. Includes a proper greeting and sign-off
5. Specifically crafted to be {response_type}
"""

        try:
            response = self.openai_client.chat.completions.create(
                model=os.getenv('OLLAMA_MODEL', 'gemma3:4b'),
                messages=[
                    {"role": "system", "content": self.system_prompt},
                    {"role": "user", "content": prompt}
                ]
            )
            return handle_response(response.choices[0].message.content)
        except Exception as e:
            return f"Error generating response: {str(e)}"

    def write_response(self, response, output_file_path):
        """Write the generated response to a text file"""
        try:
            with open(output_file_path, 'w') as file:
                file.write(response)
            return True
        except Exception as e:
            print(f"Error writing response file: {str(e)}")
            return False
        
def handle_response(response):
    """Handle the response from the AI"""
    return response.replace("```html", "").replace("```", "")

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Generate email responses with specific tones.')
    parser.add_argument('response_type',
                      help='Type of response to generate')
    args = parser.parse_args()

    # Get the directory where the script is located
    script_dir = Path(__file__).parent.absolute()
    
    # Initialize the email assistant
    assistant = EmailAssistant()
    
    # Define input and output file paths relative to the script location
    input_file = script_dir / "temp" / "email_body.txt"
    output_file = script_dir / "temp" / "draft_reply.txt"
    
    # Create temp directory if it doesn't exist
    input_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Read the email
    email_content = assistant.read_email(input_file)
    if email_content:
        # Generate response with the specified type
        response = assistant.generate_response(email_content, args.response_type)
        
        # Write response to file
        if assistant.write_response(response, output_file):
            print(f"\nResponse has been written to {output_file}")
            print("\nGenerated Response:")
            print(response)
        else:
            print("Failed to write response to file.")

if __name__ == "__main__":
    main() 
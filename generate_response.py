import sys
import json
import subprocess
import re

def generate_response_with_ollama(prompt):
    try:
        # Extract sender's first name from the prompt
        sender_name_match = re.search(r'Dear\s+(\w+)', prompt)
        sender_name = sender_name_match.group(1) if sender_name_match else ""

        # Escape special characters for PowerShell
        instruction = "Create an email response based on the following email body. Start with 'Dear' followed by the sender's first name, then include a line break. Do not insert line breaks after a fixed length. Conclude with a line break and the phrase 'Thanks,' followed by another line break and my name."

        escaped_prompt = instruction + " " + prompt.replace("'", "''").replace("\n", " ")
        
        # Define the PowerShell command
        powershell_command = f"""
        $Body = @{{
            "model" = "llama3"
            "prompt" = "{escaped_prompt}"
            "stream" = $false
        }} | ConvertTo-Json
        Invoke-RestMethod -Uri "http://localhost:11434/api/generate" -Method POST -Headers @{{"Content-Type"="application/json"}} -Body $Body
        """

        # Execute the PowerShell command
        result = subprocess.run(["powershell", "-Command", powershell_command], capture_output=True, text=True)

        # Check if the command was successful
        if result.returncode == 0:
            # Attempt to extract the response text
            lines = result.stdout.splitlines()
            response_text = []
            capture = False
            for line in lines:
                if line.strip().startswith("response"):
                    capture = True
                    response_text.append(line.split("response", 1)[-1].strip().lstrip(":").strip())
                    continue
                if capture:
                    if line.strip() == "done" or line.strip().startswith("done_reason"):
                        break
                    response_text.append(line.strip())

            # Remove empty lines and metadata from response
            cleaned_response = [line for line in response_text if line and not line.startswith("done") and not line.startswith("context")]

            if cleaned_response:
                # Remove unnecessary prefixes and anything before 'Dear'
                response = " ".join(cleaned_response).strip()
                response = response.replace("Generating response, please wait...", "").replace("Here is an appropriate email response:", "").strip()
                if "Dear" in response:
                    response = response.split("Dear", 1)[-1].strip()
                    response = "Dear " + response
                
                # Swap names in 'Dear' and 'Thanks'
                if sender_name:
                    response = response.replace(f"Dear {sender_name}", "Dear TEMP_NAME_PLACEHOLDER")
                    response = response.replace("Thanks,", f"\nThanks,\n{sender_name}")
                    response = response.replace("TEMP_NAME_PLACEHOLDER", sender_name)
                return response
            else:
                print("Debug: The response key is either empty or contains no meaningful data.")
                return "No valid response generated by Ollama."
        else:
            print(f"Error executing PowerShell command: {result.stderr}")
            return "Failed to generate response with Ollama."
    except Exception as e:
        print(f"Error calling Ollama API: {e}")
        return "Failed to generate response with Ollama."

if __name__ == "__main__":
    prompt = sys.argv[1]
    response = generate_response_with_ollama(prompt)

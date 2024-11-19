import pandas as pd
from typing import Dict, List, Optional
import os
from datetime import datetime
import requests
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

class PartFinderChatbot:
    def __init__(self, excel_path: Optional[str] = None, sharepoint_url: Optional[str] = None):
        self.parts_data = {}
        self.models_data = {}
        self.current_context = {}
        
        # Initialize data sources
        if excel_path:
            self.load_excel_data(excel_path)
        if sharepoint_url:
            self.setup_sharepoint(sharepoint_url)

    def load_excel_data(self, file_path: str) -> None:
        """Load parts and models data from Excel file"""
        try:
            # Load different sheets for different data types
            parts_df = pd.read_excel(file_path, sheet_name='Parts')
            models_df = pd.read_excel(file_path, sheet_name='Models')
            
            # Convert DataFrames to dictionaries for easier access
            self.parts_data = parts_df.to_dict('index')
            self.models_data = models_df.to_dict('index')
        except Exception as e:
            print(f"Error loading Excel data: {str(e)}")

    def setup_sharepoint(self, sharepoint_url: str) -> None:
        """Setup SharePoint connection for OneDrive integration"""
        try:
            # You would need to implement SharePoint authentication here
            # This is a placeholder for the setup
            pass
        except Exception as e:
            print(f"Error setting up SharePoint: {str(e)}")

    def get_part_info(self, model_number: str, part_description: str) -> Dict:
        """Get part information based on model number and description"""
        try:
            # Search through parts data
            for part_id, part_info in self.parts_data.items():
                if (part_info['model_number'] == model_number and 
                    part_info['description'].lower() == part_description.lower()):
                    return {
                        'part_number': part_info['part_number'],
                        'type': part_info['type'],
                        'year_sold': part_info['year_sold'],
                        'price': part_info['price']
                    }
            return None
        except Exception as e:
            print(f"Error retrieving part info: {str(e)}")
            return None

    def process_message(self, message: str) -> str:
        """Process user message and return appropriate response"""
        message = message.lower()
        
        # Extract model number if provided
        if "model" in message:
            model_number = self._extract_model_number(message)
            self.current_context['model_number'] = model_number
            return f"Model number {model_number} selected. What part are you looking for?"

        # Handle part description query
        if "part" in message or "looking for" in message:
            if 'model_number' not in self.current_context:
                return "Please provide a model number first."
            
            part_description = self._extract_part_description(message)
            part_info = self.get_part_info(self.current_context['model_number'], part_description)
            
            if part_info:
                return self._format_part_response(part_info)
            else:
                return "I couldn't find that part. Please check the model number and part description."

        # Handle price queries
        if "price" in message and 'model_number' in self.current_context:
            part_info = self.get_part_info(self.current_context['model_number'], 
                                         self.current_context.get('part_description', ''))
            if part_info:
                return f"The price for this part is ${part_info['price']:.2f}"

        # Handle diagram requests
        if "diagram" in message and 'model_number' in self.current_context:
            return "I can help you locate the diagram. Would you like to see it in the browser or download it?"

        # Default response
        return "I can help you find parts by model number. Please provide a model number to start."

    def _extract_model_number(self, message: str) -> str:
        """Extract model number from message"""
        # This is a simple implementation - you might want to use regex or more sophisticated parsing
        words = message.split()
        for word in words:
            if any(c.isdigit() for c in word):
                return word
        return ""

    def _extract_part_description(self, message: str) -> str:
        """Extract part description from message"""
        # This is a simple implementation - you might want to use more sophisticated NLP
        common_parts = ["power cord", "filter", "fan", "control board", "display"]
        for part in common_parts:
            if part in message.lower():
                return part
        return ""

    def _format_part_response(self, part_info: Dict) -> str:
        """Format part information response"""
        return f"""Here are the details for your part:
Part Number: {part_info['part_number']}
Type: {part_info['type']}
Year: {part_info['year_sold']}
Price: ${part_info['price']:.2f}

Would you like to see the diagram for this part?"""

# Example usage
def main():
    # Initialize chatbot
    chatbot = PartFinderChatbot(excel_path="parts_database.xlsx")
    
    print("Parts Finder Chatbot")
    print("Type 'quit' to exit")
    
    while True:
        user_input = input("You: ")
        if user_input.lower() == 'quit':
            break
            
        response = chatbot.process_message(user_input)
        print(f"Bot: {response}")

if __name__ == "__main__":
    main()

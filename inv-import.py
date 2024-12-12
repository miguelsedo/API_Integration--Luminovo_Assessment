import pandas as pd
import json
import requests
import time
from typing import Optional
from datetime import datetime

"""Client config"""
path = "API_Integration-Luminovo_Assessment\Availability.xlsx"
bearer_token = "GET TOKEN FROM LUMINOVO"

LUMINOVO_API_URL = "https://api.luminovo.com"


def get_from_excel(excel_file_path: str, sheet_name: int):

    try:
        # Read the Excel file
        content = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        
        inventory = []
        
        # Iterate over each row in the DataFrame
        for _, row in content.iterrows():
            # Extract values from the row
            available = int(row.iloc[0])  
            total = int(row.iloc[1])      
            price = float(row.iloc[2])     
            number = str(row.iloc[3])     
            
            # Create the offer dictionary for this row
            item = {
                "availability": {
                    "available_stock": available,
                    "total_stock": total
                },
                "part": {"internal_part_number": number},
                "prices": [{
                    "unit_price": {"amount": str(price), "currency": "EUR"}
                }],
                "supplier": {
                    "type": "Internal",
                    "supplier": "Muenchen"
                }
            }
            
            # Add the offer to the list
            inventory.append(item)
        return inventory
        
    except FileNotFoundError:
        print(f"Error: File '{excel_file_path}' not found")
        return None
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None
    
def get_access_token(bearer_token: str):

    """Get access token from Luminovo API using authentication token"""
    try:
        response = requests.post(
            f"{LUMINOVO_API_URL}/token",
            headers={"content-type": "application/json"},
            json={"token": bearer_token}
        )
        if response.status_code == 200:
            return response.text
        print(f"Failed to get access token. Status code: {response.status_code}")
        return None
    except Exception as e:
        print(f"Error getting access token: {str(e)}")
        return None

def send_to_luminovo(inventory: list, access_token: str) -> Optional[requests.Response]:

    """Send data to Luminovo API"""
    try:
        headers = {"Authorization": f"Bearer {access_token}"}
        response = requests.post(
            f"{LUMINOVO_API_URL}/offers/import", 
            json=inventory,
            headers=headers
        )
        print(f"API Response: {response.status_code}")
        return response
    except Exception as e:
        print(f"Error sending to API: {str(e)}")
        return None


# Main execution
if __name__ == "__main__":
    # Get access token
    access_token = get_access_token(bearer_token)

    # If successful, process excel and import data
    if access_token:
        inventory = get_from_excel(path, 0)
        if inventory:
            send_to_luminovo(inventory, access_token)

import pandas as pd
import json
import requests
import time
from typing import Optional
from datetime import datetime

path = "luminovo-integration\Availability.xlsx"
tenant_name = "ElektronikAG"
auth_token = "GET TOKEN FROM LUMINOVO"

LUMINOVO_API_URL = f"https://{tenant_name}.luminovo.com/api/offers/import"


def get_from_excel(excel_file_path: str, sheet_name: int) -> Optional[dict]:

    """Read Excel file and convert to Luminovo's Endpoint format"""
    try:
        content = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        
        # Extract values from the columns in the excel file.
        available = int(content.iloc[0, 0])  
        total = int(content.iloc[0, 1])      
        price = float(content.iloc[0, 2])     
        number = str(content.iloc[0, 3])     
        
        # Data is restructured according to the Luminovo API documentation
        offer = {
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
        
        print("Generated offer:")
        print(json.dumps(offer, indent=4))
        return offer
        
    except FileNotFoundError:
        print(f"Error: File '{excel_file_path}' not found")
        return None
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None
    
def get_access_token(auth_token: str):

    """Get access token from Luminovo API using authentication token"""
    try:
        response = requests.post(
            f"{LUMINOVO_API_URL}/token",
            headers={"content-type": "application/json"},
            json={"token": auth_token}
        )
        if response.status_code == 200:
            return response.text
        print(f"Failed to get access token. Status code: {response.status_code}")
        return None
    except Exception as e:
        print(f"Error getting access token: {str(e)}")
        return None

def send_to_luminovo(offer: dict, access_token: str) -> Optional[requests.Response]:

    """Send data to Luminovo API"""
    try:
        headers = {"Authorization": f"Bearer {access_token}"}
        response = requests.post(
            LUMINOVO_API_URL, 
            json=offer,
            headers=headers
        )
        print(f"API Response: {response.status_code}")
        return response
    except Exception as e:
        print(f"Error sending to API: {str(e)}")
        return None


def sync_data():
    """Function to handle one sync cycle"""
    print(f"\nStarting sync at {datetime.now()}")
    
    # Get access token
    access_token = get_access_token(auth_token)

    # If successful, process excel and import data
    if access_token:
        offer = get_from_excel(path, 0)
        if offer:
            send_to_luminovo(offer, access_token)


# Main execution
if __name__ == "__main__":
    sync_schedule = 7200  # 2 hours
    
    print("Starting periodic sync every 2 hours...")
    
    while True:
        sync_data()

        # Pauses Sync for 7200 seconds, 2 hours
        time.sleep(sync_schedule)

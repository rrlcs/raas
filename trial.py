from googleapiclient.discovery import build
from google.oauth2 import service_account
from dotenv import load_dotenv


from PIL import Image, ImageDraw, ImageFont

def create_image_from_data(values):
    # Define image size and other parameters
    cell_width = 100
    cell_height = 30
    font_size = 14

    # Determine image dimensions
    num_rows = len(values)
    num_cols = len(values[0]) if num_rows > 0 else 0
    image_width = cell_width * num_cols
    image_height = cell_height * num_rows

    # Create a new image with a white background
    image = Image.new('RGB', (image_width, image_height), 'white')
    draw = ImageDraw.Draw(image)
    
    # Load a font
    font = ImageFont.truetype("Arial.ttf", font_size)

    # Draw cells and text
    # for row_idx, row in enumerate(values):
    #     for col_idx, cell in enumerate(row):
    #         x1 = col_idx * cell_width
    #         y1 = row_idx * cell_height
    #         x2 = x1 + cell_width
    #         y2 = y1 + cell_height

    #         draw.rectangle([x1, y1, x2, y2], outline='black')
    #         draw.text((x1 + 10, y1 + 5), str(cell), fill='black', font=font)
    # Modify the $SELECTION_PLACEHOLDER$ code
    for row_idx, row in enumerate(values):
        for col_idx, cell in enumerate(row):
            x1 = col_idx * cell_width
            y1 = row_idx * cell_height
            x2 = x1 + cell_width
            y2 = y1 + cell_height
            # Calculate the color based on cell value
            color_value = int(cell)
            red = min(255, max(0, 255 - (color_value * 255 // 100)))
            green = min(255, max(0, color_value * 255 // 100))
            blue = 0
            color = (red, green, blue)
            draw.rectangle([x1, y1, x2, y2], outline='black', fill=color)
            draw.text((x1 + 10, y1 + 5), str(cell), fill='black', font=font)

    # Save the image
    image_path = 'sheet_data_image.png'
    image.save(image_path)
    # return image_path
    

# Load environment variables from .env file
load_dotenv('.env')

# Service account credentials file
SERVICE_ACCOUNT_FILE = 'sadhana-428908-168775d3d385.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Spreadsheet ID and range
SPREADSHEET_ID = '1-i3ZEiFQlnQJtDaWOp7zVf-iTFe_NPlAJqPT34bC0m4'
RANGE_NAME = 'Reports!A1:D7'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

service = build('sheets', 'v4', credentials=credentials)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
values = result.get('values', [])

if not values:
    print('No data found.')
else:
    for row in values:
        print(row)

create_image_from_data(values)
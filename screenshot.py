from PIL import Image
import pytesseract
import pandas as pd

# Load the image from path
name = "Screenshot from 2024-04-30 15-56-18"
img = Image.open("input/" + name + ".png")

# Use Tesseract to do OCR on the image
text = pytesseract.image_to_string(img)

# You might need custom processing here to split text into rows and columns
# This is a basic example:
rows = text.split("\n")
data = [row.split() for row in rows if row.strip()]

# Convert list to DataFrame
df = pd.DataFrame(data)

# Export to CSV
df.to_csv("output/output.csv", index=False)

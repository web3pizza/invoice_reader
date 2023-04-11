import os
import io
import glob
from pdf2image import convert_from_path
import pandas as pd
from google.cloud import vision
from google.cloud.vision import types
from openpyxl import Workbook

# Load credentials from JSON key file
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "/Users/aaron/keys/new-key1.json"

# Initialize the client
client = vision.ImageAnnotatorClient()

# Set up a list of keywords to search for in the invoice text
keywords = ["Product No.", "Description", "Weight", "Price", "Ext."]

# Set up the directory for storing Excel files
if not os.path.exists("invoices"):
    os.mkdir("invoices")

# Set up the directory for storing processed PDF files
if not os.path.exists("processed"):
    os.mkdir("processed")

# Prompt user for path to scanned image
image_path = input("Enter the path to the scanned image: ")

# Convert the PDF file to images
print("Converting PDF to images...")
images = convert_from_path(image_path)

# Loop through the images and detect text in each one
text = ""
for i, image in enumerate(images):
    print("Detecting text in image {}/{}...".format(i + 1, len(images)))
    with io.BytesIO() as output:
        image.save(output, format="JPEG")
        content = output.getvalue()

    # Detect text in the image
    image = types.Image(content=content)
    response = client.text_detection(image=image)
    texts = response.text_annotations

    # Concatenate the text from all images
    for text_annotation in texts:
        text += text_annotation.description + "\n"

# Check if any of the keywords are present in the text
if any(keyword.lower() in text.lower() for keyword in keywords):
    # Split the text into lines
    lines = text.split("\n")

    # Initialize lists to store the data
    product_nos = []
    descriptions = []
    weights = []
    prices = []
    exts = []

    # Loop through the lines and extract the relevant data
    for line in lines:
        for keyword in keywords:
            if keyword.lower() in line.lower():
                index = line.lower().index(keyword.lower())
                value = line[index + len(keyword) :].strip()
                if keyword == "Product No.":
                    product_nos.append(value)
                elif keyword == "Description":
                    descriptions.append(value)
                elif keyword == "Weight":
                    weights.append(value)
                elif keyword == "Price":
                    prices.append(value)
                elif keyword == "Ext.":
                    exts.append(value)

    # Check that all arrays are the same length
    array_lengths = [
        len(product_nos),
        len(descriptions),
        len(weights),
        len(prices),
        len(exts),
    ]
    print("Array lengths:", array_lengths)
    if len(set(array_lengths)) != 1:
        print("Error: All arrays must be of the same length")
    else:
        # Create a Pandas DataFrame from the data
        data = {
            "Product No.": product_nos,
            "Description": descriptions,
            "Weight": weights,
            "Price": prices,
            "Ext.": exts,
        }
        print("Extracted data:")
        print(data)
        df = pd.DataFrame(data)

        # Print the DataFrame
        print("DataFrame:")
        print(df)

        # Save the DataFrame to an Excel file
        output_path = os.path.join("invoices", os.path.basename(image_path) + ".xlsx")
        print("Saving output to {}...".format(output_path))

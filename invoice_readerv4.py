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
        product_no = ""
        description = ""
        weight = ""
        price = ""
        ext = ""

        for keyword in keywords:
            if keyword.lower() in line.lower():
                index = line.lower().index(keyword.lower())
                value = line[index + len(keyword) :].strip()

                if keyword == "Product No.":
                    product_no = value
                elif keyword == "Description":
                    description = value
                elif keyword == "Weight":
                    weight = value
                elif keyword == "Price":
                    price = value
                elif keyword == "Ext.":
                    ext = value

        product_nos.append(product_no)
        descriptions.append(description)
        weights.append(weight)
        prices.append(price)
        exts.append(ext)

    # Determine the maximum length of the lists
    max_length = max(
        len(product_nos), len(descriptions), len(weights), len(prices), len(exts)
    )

    # Add empty strings to the lists where there is no data
    product_nos += [""] * (max_length - len(product_nos))
    descriptions += [""] * (max_length - len(descriptions))
    weights += [""] * (max_length - len(weights))
    prices += [""] * (max_length - len(prices))
    exts += [""] * (max_length - len(exts))

    # Create a Pandas DataFrame from the data
    data = {
        "Product No.": product_nos,
        "Description": descriptions,
        "Weight": weights,
        "Price": prices,
        "Ext.": exts,
    }
    df = pd.DataFrame(data)

    # Print the DataFrame
    print("DataFrame:")
    print(df)

    # Save the DataFrame to an Excel file
    output_path = os.path.join("invoices", os.path.basename(image_path) + ".xlsx")
    print("Saving output to {}...".format(output_path))

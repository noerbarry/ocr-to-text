import streamlit as st
from PIL import Image
import pytesseract
import pandas as pd
import io
import base64

# Path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

def perform_ocr(image):
    ocr_result = pytesseract.image_to_string(image, lang='ind')
    return ocr_result

def main():
    st.title("OCR KTP Indonesia to Excel")

    method = st.radio("Select Input Method:", ("Upload Image", "Use Webcam"))

    if method == "Upload Image":
        uploaded_file = st.file_uploader("Upload an image", type=["jpg", "jpeg", "png"])

        if uploaded_file is not None:
            image = Image.open(uploaded_file)
            st.image(image, caption="Uploaded Image", use_column_width=True)

            if st.button("Perform OCR"):
                ocr_result = perform_ocr(image)
                add_to_excel(ocr_result)

    elif method == "Use Webcam":
        st.warning("Please grant camera permissions to your browser for webcam access.")
        camera_image = st.camera()

        if st.button("Capture Image"):
            if camera_image is not None:
                # Convert camera image to PIL Image
                image = Image.open(io.BytesIO(base64.b64decode(camera_image)))

                # Perform OCR on the captured image
                ocr_result = perform_ocr(image)
                add_to_excel(ocr_result)

def add_to_excel(ocr_result):
    # Load existing Excel file or create a new one if it doesn't exist
    try:
        existing_df = pd.read_excel("ocr_result.xlsx")
    except FileNotFoundError:
        existing_df = pd.DataFrame()

    # Process OCR result and add to DataFrame
    lines = ocr_result.split('\n')
    data = {}
    for line in lines:
        key_value = line.split(':')
        if len(key_value) == 2:
            key = key_value[0].strip()
            value = key_value[1].strip()
            data[key] = [value]

    new_row_df = pd.DataFrame(data)

    # Concatenate existing DataFrame with new row
    updated_df = pd.concat([existing_df, new_row_df], ignore_index=True)

    # Save updated DataFrame to Excel
    updated_df.to_excel("ocr_result.xlsx", index=False)

    st.success("OCR result added to Excel")

if __name__ == "__main__":
    main()

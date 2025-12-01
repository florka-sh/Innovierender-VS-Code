import os
import urllib.request
import sys

# URL for German language data (deu.traineddata)
DEU_URL = "https://github.com/tesseract-ocr/tessdata/raw/main/deu.traineddata"
TESSDATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tesseract-OCR", "tessdata")

def download_progress(count, block_size, total_size):
    percent = int(count * block_size * 100 / total_size)
    sys.stdout.write(f"\rDownloading deu.traineddata... {percent}%")
    sys.stdout.flush()

def download_german_language():
    if not os.path.exists(TESSDATA_DIR):
        print(f"Error: tessdata directory not found at {TESSDATA_DIR}")
        print("Please ensure Tesseract-OCR is installed/present in the project folder.")
        return

    target_path = os.path.join(TESSDATA_DIR, "deu.traineddata")
    
    if os.path.exists(target_path):
        print(f"German language data already exists at {target_path}")
        return

    print(f"Downloading German language data to {TESSDATA_DIR}...")
    try:
        urllib.request.urlretrieve(DEU_URL, target_path, reporthook=download_progress)
        print("\nDownload complete!")
    except Exception as e:
        print(f"\nError downloading file: {e}")

if __name__ == "__main__":
    download_german_language()

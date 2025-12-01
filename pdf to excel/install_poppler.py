import os
import urllib.request
import zipfile
import shutil
import sys

# URL for Poppler Windows binary (Release 24.02.0-0)
POPPLER_URL = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.02.0-0/Release-24.02.0-0.zip"
INSTALL_DIR = "poppler"

def download_progress(count, block_size, total_size):
    percent = int(count * block_size * 100 / total_size)
    sys.stdout.write(f"\rDownloading Poppler... {percent}%")
    sys.stdout.flush()

def install_poppler():
    print(f"Installing Poppler to: {os.path.abspath(INSTALL_DIR)}")
    
    if os.path.exists(INSTALL_DIR):
        print("Poppler directory already exists. Skipping download.")
        return

    zip_path = "poppler.zip"
    
    try:
        print(f"Downloading from {POPPLER_URL}")
        urllib.request.urlretrieve(POPPLER_URL, zip_path, reporthook=download_progress)
        print("\nDownload complete.")

        print("Extracting...")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall("temp_poppler")
        
        # Move the inner folder to the final location
        # The zip usually contains a folder like "poppler-24.02.0"
        extracted_root = "temp_poppler"
        inner_folder = os.listdir(extracted_root)[0]
        shutil.move(os.path.join(extracted_root, inner_folder), INSTALL_DIR)
        
        # Cleanup
        os.remove(zip_path)
        shutil.rmtree(extracted_root)
        
        print("Installation complete!")
        print(f"Poppler bin path: {os.path.abspath(os.path.join(INSTALL_DIR, 'Library', 'bin'))}")

    except Exception as e:
        print(f"\nError installing Poppler: {e}")
        if os.path.exists(zip_path):
            os.remove(zip_path)
        if os.path.exists("temp_poppler"):
            shutil.rmtree("temp_poppler")

if __name__ == "__main__":
    install_poppler()

import os
import shutil
import zipfile
import subprocess

APP_NAME = "convert_all"
DIST_DIR = "dist"
BUILD_DIR = "build"
ZIP_NAME = f"{APP_NAME}_portable.zip"

# Step 1: Clean old build
for path in [DIST_DIR, BUILD_DIR, f"{APP_NAME}.spec", ZIP_NAME]:
    if os.path.isdir(path):
        shutil.rmtree(path)
    elif os.path.isfile(path):
        os.remove(path)

# Step 2: Run PyInstaller
print("📦 Building exe with PyInstaller...")
subprocess.run([
    "pyinstaller",
    "--onefile",
    f"--add-data=config.json;.",
    f"--add-data=invoices;invoices",
    f"{APP_NAME}.py"
], check=True)

# Step 3: Prepare package folder
package_dir = f"{APP_NAME}_package"
if os.path.exists(package_dir):
    shutil.rmtree(package_dir)
os.makedirs(package_dir, exist_ok=True)

# Copy exe
shutil.copy(os.path.join(DIST_DIR, f"{APP_NAME}.exe"), package_dir)

# Copy config.json
shutil.copy("config.json", package_dir)

# Create empty invoices folder
invoices_dir = os.path.join(package_dir, "invoices")
os.makedirs(invoices_dir, exist_ok=True)

# Step 4: Zip it
print("📦 Creating zip file...")
with zipfile.ZipFile(ZIP_NAME, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for root, _, files in os.walk(package_dir):
        for file in files:
            full_path = os.path.join(root, file)
            arcname = os.path.relpath(full_path, package_dir)
            zipf.write(full_path, arcname)

print(f"✅ Done! Portable zip created: {ZIP_NAME}")

from PIL import Image
import sys
import os

try:
    img = Image.open("icon.icns")
    # Save as ICO (containing multiple sizes)
    # Windows icons standard sizes: 16, 32, 48, 64, 128, 256
    img.save("icon.ico", format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    print("Successfully converted icon.icns to icon.ico")
except Exception as e:
    print(f"Error converting icon: {e}")
    sys.exit(1)

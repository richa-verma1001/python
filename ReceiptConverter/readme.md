This file has instruction on how to take any .py file and package that into an .exe

Step 1. Activate your virtual environment
If you already created one earlier:

cd C:\pdf_to_word_app
.\venv\Scripts\activate
If not, create one:

python -m venv venv
.\venv\Scripts\activate
pip install --upgrade pip
pip install pyinstaller pdf2image python-docx Pillow
Note to self (@RV): stay on pip version 24.3.1 and not upgrade to 25.3. The new version only works with newer openai API.  Existing files will need to be convered for openai. 

📁 Step 2. (Optional) Bundle Poppler
If you plan to ship Poppler with your .exe so users don’t install it:
	1. Keep Poppler extracted here:

C:\pdf_to_word_app\poppler\Library\bin\pdftoppm.exe
	2. Update your script (top of file) with this logic:

import sys, os
from pdf2image import convert_from_path

if getattr(sys, '_MEIPASS', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

poppler_path = os.path.join(base_path, "poppler", "Library", "bin")

And change your convert call to:

images = convert_from_path(file_path, dpi=200, poppler_path=poppler_path)

🧱 Step 3. Run PyInstaller
If bundling Poppler:

pyinstaller --onefile --noconsole --add-data "C:\pdf_to_word_app\poppler;poppler" pdf_to_word_app.py
If Poppler is already installed system-wide (PATH added):

pyinstaller --onefile --noconsole pdf_to_word_app.py
or 
pyinstaller --noconsole --onefile receipts_to_word.py


Explanation
	• --onefile → single .exe
	• --noconsole → no command window (since you use Tkinter dialogs)
	• --add-data → include the Poppler folder (Windows uses ; separator)
	• Output goes into:

C:\pdf_to_word_app\dist\pdf_to_word_app.exe

🔍 Step 4. Test your EXE
	1. Go to:

cd dist
	2. Double-click pdf_to_word_app.exe
You should see the folder selection dialog pop up.
Test it with a folder of PDFs + JPGs.
✅ Confirm it creates Output/Combined_PDFs_and_Images.docx.

🧪 Step 5. Test on a clean PC (no Python installed)
Copy just the .exe (and Poppler folder if you didn’t bundle it) to another Windows PC.
Try running it there.
👉 If you bundled Poppler, it should “just work.”
👉 If you didn’t, and users haven’t added Poppler to PATH, you’ll get an error like “Poppler not found”.

🚀 Step 6. Distribute to users
Option 1: Simple portable version
Distribute a .zip like this:

pdf_to_word_app.zip
 ├── pdf_to_word_app.exe
 └── README.txt

README example:

1. Extract the zip.
2. Double-click pdf_to_word_app.exe.
3. Select the folder containing PDFs/JPGs.
4. The combined Word file will appear in an "Output" subfolder.
If you bundled Poppler, that’s all users need.
If not, your README should tell them to install Poppler or unzip it in C:\poppler and add it to PATH.

Option 2: Installer (for non-technical users)
Later, you can create an installer with Inno Setup or NSIS that:
	• Places the exe in Program Files
	• Adds Start Menu shortcut
	• Optionally adds Poppler silently
But that’s optional; .exe distribution works fine standalone.

🧰 Step 7. What users need installed
If you packaged correctly (--onefile with dependencies):
	✅ Users do not need:
		○ Python
		○ pip
		○ any libraries
		○ administrator access (unless they install in Program Files)
The only external dependency is Poppler:
	• If bundled inside → no install needed.
	• If not bundled → users must have Poppler installed or in PATH.

⚠️ Notes and Tips
Issue	Solution
EXE flagged by antivirus	Happens occasionally with PyInstaller — safe to ignore or sign the exe if distributing widely
Big file size (~80–100 MB)	Normal — Python + dependencies are inside
Slow first start	.exe self-extracts to temp once, then runs fast
Need debug	Remove --noconsole so you can see error output

✅ Quick Recap
Step	Command / Action
Create venv	python -m venv venv && .\venv\Scripts\activate
Install deps	pip install pyinstaller pdf2image python-docx Pillow
Package	pyinstaller --onefile --noconsole --add-data "C:\pdf_to_word_app\poppler;poppler" pdf_to_word_app.py
Find exe	dist\pdf_to_word_app.exe
Test & share	Copy or zip and distribute

Would you like me to give you a ready-to-copy PyInstaller command that includes your current folder structure (once you confirm where Poppler is extracted)?
That’ll make it literally one command to run.




ChatGPT can make mistakes. Check important info.

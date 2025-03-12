# toPDF_converter


This PowerShell script automates the process of converting PowerPoint (.pptx) and Word (.docx) files to PDF format. It scans the directory where the script is located, converts all .pptx and .docx files it finds, and saves them as .pdf in the same directory.

Requirements
Microsoft PowerPoint and Microsoft Word installed on your machine (the script uses COM objects to interact with these applications).
PowerShell (the script is designed to run on Windows with PowerShell).

Files Included
to_pdf.ps1: The PowerShell script that performs the conversion.

Instructions
Step 1: Prepare Your Files
Place all the .pptx and .docx files you want to convert into the same folder as the ConvertToPDF.ps1 script.

Step 2: Run the Script
1. Open PowerShell: Open PowerShell on your Windows machine.
2. Navigate to the Script Directory: Use the cd command to change the directory to where the script is located. For example:
            cd "C:\Path\To\Your\Script"

3. Run the Script: Type the following command to execute the script:
            .\ConvertToPDF.ps1
   
Step 3: Conversion Process
1. The script will automatically detect and convert all .pptx (PowerPoint) and .docx (Word) files in the same folder.
2. The converted PDF files will be saved in the same directory with the same name but with a .pdf extension.
For example:
            a. Presentation.pptx will be converted to Presentation.pdf.
            b. Document.docx will be converted to Document.pdf.
   
Notes
1. The script is designed to run in the same folder as the .pptx and .docx files you want to convert.
2. The script uses the COM automation interface to control PowerPoint and Word, so these applications must be installed on your machine and accessible.

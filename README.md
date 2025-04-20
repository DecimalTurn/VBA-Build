# VBA-Build

For a demo on how to use this GitHub Action: [VBA-Build-Demo](https://github.com/DecimalTurn/VBA-Build-Demo)

![Banner](https://github.com/DecimalTurn/VBA-Build/blob/main/images/Banner.png?raw=true)

## How does it work?

This GitHub Action automates the process of building VBA-Enabled Office documents from XML and VBA source code:

The main script is contained in `Main.ps1` and will perform the following actions:

- Installs 7-Zip for handling file compression
- Installs Office 365 (via Chocolatey) to provide the Office applications needed
- Zip File Creation:
    - Takes XML source files from your Office document structure
    - Compresses them into a ZIP file using 7-Zip
    - Renames the ZIP file with the appropriate Office extension (e.g., .xlsm)
- VBA Integration:
    - Ensures no Office applications are running that could interfere
    - Enables the Visual Basic Object Model (VBOM) and general macro permissions in the registry of the Windows GitHub Worker
    - Opens the Office file and imports all modules (.bas), Forms (.frm) and Class Modules (.cls) from your source directory
- Output:
    - Saves the final document with embedded VBA code
    - Uploads the resulting documents as build artifacts

## Supported File Formats

* Excel (.xlsm, .xlam, .xlsb)
* Word (.docm)
* PowerPoint (.pptm)

## Why? 

I mean, why not? Every other programming language has something like this to build and package your code.

This could be used to:

- Keep your VBA code in plain text format for version control
- Automate builds as part of your CI/CD pipeline
- Generate ready-to-use Office documents without manual intervention

## What's next?

Depending on the reaction of the community, I might add support for:
- Workbook, Worksheet and Document (Word) Objects
- Signature of the VBA Project (to facilitate distribution)
- More complex file structure using [vba-block](https://www.vba-blocks.com/manifest/) configuration file (manifest file)
- Microsoft Access file formats

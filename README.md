# VBA-Build

For a demo on how to use this GitHub Action: [VBA-Build-Demo](https://github.com/DecimalTurn/VBA-Build-Demo)

![Banner](https://github.com/DecimalTurn/VBA-Build/blob/main/images/Banner.png?raw=true)

## How does it work?

This GitHub Action automates the process of building VBA-Enabled Office documents from XML and VBA source code:

The main script is contained in `Main.ps1` and will perform the following actions:

- Install 7-Zip for handling file compression
- Install Office 365 (via Chocolatey) to provide the Office applications needed
- Create Office file from XML source:
    - Find the XML source files representing your Office document structure
    - Compress them into a zip file using 7-Zip
    - Rename the zip file with the appropriate Office extension (e.g., .xlsm)
- Import the VBA components:
    - Enable the Visual Basic Object Model (VBOM) and general macro permissions in the registry of the Windows GitHub Worker
    - Opens the Office file and imports all modules (.bas), Forms (.frm) and Class Modules (.cls) from your source directory
- Run tests:
    - If a testing framework was specified, install the required dependencies
    - Run the tests and output the results to the console
- Generate final output:
    - Save the final document with embedded VBA code
    - Upload the resulting documents as build artifacts

## Supported File Formats

* Excel (.xlsm, .xlam)
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
- Template formats .xltm, .dotm and .potm
- Signature of the VBA Project (to facilitate distribution)
- More complex file structure using [vba-block](https://www.vba-blocks.com/manifest/) configuration file (manifest file)
- Microsoft Access file formats

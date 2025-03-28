'This script will zip the content of src/XMLSource then rename the zip file to have .xlsm extension
' and move it to the destination folder.
' Then it will open the Excel file and import the VBA code modules in the folder src/Modules
' The script will then save the file and close Excel.

dim fso, zipFile, srcFolder, destFolder, excelApp, wb, newFileName
' Set the source and destination folders
srcFolder = "src\XMLSource"
destFolder = "\"
' Create a FileSystemObject
set fso = CreateObject("Scripting.FileSystemObject")

'Zip the source folder by running node zip.js
' This assumes you have Node.js installed and the zip.js script is in the same directory
Shell "node zip.js "

'Rename the file from Excel_Skeleton.zip in the current dir to Excel_Skeleton.xlsm
zipFile = "Excel_Skeleton.zip"
newFileName = destFolder & "Excel_Skeleton.xlsm"

' Check if the zip file exists
if fso.FileExists(zipFile) then
    ' Rename the zip file to .xlsm
    fso.MoveFile zipFile, "Excel_Skeleton.xlsm"
else
    WScript.Echo "Zip file not found: " & zipFile
    WScript.Quit 1
end if
' Check if the destination folder exists
if not fso.FolderExists(destFolder) then
    ' Create the destination folder
    fso.CreateFolder destFolder
end if
' Move the renamed file to the destination folder
if fso.FileExists("Excel_Skeleton.xlsm") then
    fso.MoveFile "Excel_Skeleton.xlsm", destFolder & "Excel_Skeleton.xlsm"
else
    WScript.Echo "File not found: Excel_Skeleton.xlsm"
    WScript.Quit 1
end if
' Check if the destination folder exists
if not fso.FolderExists(destFolder) then
    ' Create the destination folder
    fso.CreateFolder destFolder
end if
' Move the renamed file to the destination folder
if fso.FileExists("Excel_Skeleton.xlsm") then
    fso.MoveFile "Excel_Skeleton.xlsm", newFileName
else
    WScript.Echo "File not found: Excel_Skeleton.xlsm"
    WScript.Quit 1
end if



' Create a new Excel application
set excelApp = CreateObject("Excel.Application")
' Create a new workbook
set wb = excelApp.Workbooks.Open(newFileName)
' Make Excel visible
'excelApp.Visible = true

' Import the VBA code modules from the src/Modules folder
dim moduleFolder, moduleFile
moduleFolder = "src\Modules"
' Check if the module folder exists
if fso.FolderExists(moduleFolder) then
    ' Loop through each file in the module folder
    for each moduleFile in fso.GetFolder(moduleFolder).Files
        ' Check if the file is a .bas file
        if LCase(fso.GetExtensionName(moduleFile.Name)) = "bas" then
            ' Import the module into the workbook
            wb.VBProject.VBComponents.Import moduleFile.Path
        end if
    next
else
    WScript.Echo "Module folder not found: " & moduleFolder
    WScript.Quit 1
end if
' Save the workbook
wb.Save
' Close the workbook
wb.Close
' Quit Excel
excelApp.Quit

' Clean up
set wb = nothing
set excelApp = nothing
set fso = nothing
' End of script

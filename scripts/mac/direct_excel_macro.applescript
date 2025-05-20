on run argv
    set workbookPath to item 1 of argv
    set macroName to item 2 of argv
    
    tell application "Microsoft Excel"
        activate
        delay 2
        
        -- Open workbook
        set workbook_obj to open workbook workbook file name workbookPath
        delay 2
        
        -- Run the macro
        run VB macro macroName
        delay 2
        
        -- Save and close
        save workbook workbook_obj
        close workbook workbook_obj
        
        return "Macro executed successfully"
    end tell
end run
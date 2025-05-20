on run
  try
    tell application "System Events"
      tell process "Microsoft Excel"
        delay 2
        
        -- Take an approach where we just click any button that seems like it would bypass setup screens
        set foundButton to false
        
        -- First check for the welcome screen "Skip" button
        try
          -- Dump UI elements to debug in the log
          log "Searching for buttons in Excel windows..."
          set allWindows to every window
          repeat with w in allWindows
            log "Window: " & name of w
            try
              set allButtons to every button of w
              repeat with b in allButtons
                try
                  log "Button: " & name of b
                end try
              end repeat
            on error
              log "Error getting buttons"
            end try
          end repeat
          
          -- Look for any button that might skip setup
          repeat with w in allWindows
            try
              -- Common variations of the Skip to read-only button text
              if exists button "Skip to read-only mode" of w then
                click button "Skip to read-only mode" of w
                set foundButton to true
                log "Clicked 'Skip to read-only mode' button"
                exit repeat
              end if
              
              if exists button "Skip" of w then
                click button "Skip" of w
                set foundButton to true
                log "Clicked 'Skip' button"
                exit repeat
              end if
              
              if exists button "Continue" of w then
                click button "Continue" of w
                set foundButton to true
                log "Clicked 'Continue' button"
                exit repeat
              end if
              
              if exists button "Next" of w then
                click button "Next" of w
                set foundButton to true
                log "Clicked 'Next' button"
                exit repeat
              end if
              
              if exists button "Cancel" of w then
                click button "Cancel" of w
                set foundButton to true
                log "Clicked 'Cancel' button"
                exit repeat
              end if
              
              if exists button "OK" of w then
                click button "OK" of w
                set foundButton to true
                log "Clicked 'OK' button"
                exit repeat
              end if
              
              -- Try to detect buttons by their AXDescription if they don't have a proper name
              set allButtons to every button of w
              repeat with b in allButtons
                try
                  set buttonDesc to description of b
                  log "Button description: " & buttonDesc
                  if buttonDesc contains "skip" or buttonDesc contains "read-only" then
                    click b
                    set foundButton to true
                    log "Clicked button with description containing 'skip' or 'read-only'"
                    exit repeat
                  end if
                on error
                  -- Just continue if we can't get the description
                end try
              end repeat
            on error errMsg
              log "Error while checking window: " & errMsg
            end try
            
            if foundButton then exit repeat
          end repeat
          
          -- If no button found, try clicking in specific screen areas where Skip buttons might be
          if not foundButton then
            log "No button found by name, trying to click in common Skip button locations"
            
            -- Try to click in the bottom right where Skip buttons often are
            set screenSize to get size of window 1
            set screenWidth to item 1 of screenSize
            set screenHeight to item 2 of screenSize
            
            -- Bottom right area
            click at {screenWidth - 100, screenHeight - 50}
            log "Clicked at bottom right position where Skip buttons often are"
          end if
        on error errMsg
          log "Error in UI interaction: " & errMsg
        end try
        
        delay 2
        
        -- Press Escape key in case any dialog can be dismissed that way
        key code 53 -- Escape key
        delay 1
        
        return "Attempted to click through Excel welcome screens"
      end tell
    end tell
  on error errMsg
    return "Error in Excel UI handling: " & errMsg
  end try
end run
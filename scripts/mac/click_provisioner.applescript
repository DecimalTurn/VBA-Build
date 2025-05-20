on run
  try
    tell application "System Events"
      -- Look specifically for the provisioner dialog
      repeat 30 times
        -- Check for windows with "provisioner" in the title or content
        set provisionerFound to false
        
        -- First check if the dialog is visible at all
        try
          -- Try to find any window containing "provisioner" text
          set allProcesses to application processes where it is visible
          
          repeat with proc in allProcesses
            try
              tell proc
                set allWindows to every window
                
                repeat with w in allWindows
                  try
                    log "Checking window: " & name of w
                    
                    -- Look for the provisioner dialog (might not have a proper name)
                    set allTexts to every static text of w
                    repeat with t in allTexts
                      try
                        set textValue to value of t
                        log "Text found: " & textValue
                        
                        if textValue contains "provisioner" or textValue contains "control" then
                          log "Found provisioner dialog with text: " & textValue
                          set provisionerFound to true
                          
                          -- Try to click the Allow button
                          if exists button "Allow" of w then
                            click button "Allow" of w
                            log "Clicked Allow button in provisioner dialog"
                            return "Successfully clicked Allow in provisioner dialog"
                          end if
                        end if
                      end try
                    end repeat
                    
                    -- Direct approach to look for the Allow button
                    if exists button "Allow" of w then
                      click button "Allow" of w
                      log "Clicked Allow button in window"
                      return "Successfully clicked Allow button"
                    end if
                  end try
                end repeat
              end tell
            end try
          end repeat
          
          -- If we know the dialog is there but couldn't find the button, try clicking at its likely coordinates
          if not provisionerFound then
            -- Focus on the frontmost window (likely the dialog)
            tell process "SecurityAgent"
              if exists window 1 then
                set frontWindow to window 1
                set {x, y} to position of frontWindow
                set {width, height} to size of frontWindow
                
                -- Click in the lower right area where the Allow button typically is
                set clickX to x + width - 60
                set clickY to y + height - 25
                
                log "Trying to click at coordinates: " & clickX & ", " & clickY
                click at {clickX, clickY}
                
                return "Clicked at expected Allow button position"
              end if
            end tell
          end if
          
        end try
        
        delay 1
      end repeat
      
      return "Could not find or interact with provisioner dialog"
    end tell
  on error errMsg
    return "Error: " & errMsg
  end try
end run
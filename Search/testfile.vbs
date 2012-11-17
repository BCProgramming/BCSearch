
    Dim shellobj

	on error resume next
	 Set ShellObj = CreateObject("WScript.Shell")
	 if err <> 0 then
		msgbox "error:" & err
	else
		ShellObj.Exec ".\donate.html"
		if err <> 0 then
			msgbox error & " " & err.number 
		end if 
	 end if 
	
	
	
	

 set wsh1=wscript.createobject("WScript.Shell")
timeout = Now + TimeValue("00:00:05")
 do 
 loop until wsh1.appactivate("Message from Webpage") Or Now > timeout
 if wsh1.appactivate("Message from Webpage") then 

   wscript.sleep 1000 

   wsh1.sendkeys "~"
 end if


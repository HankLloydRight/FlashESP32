' hanklloydright FlashESP32 -- Erase Flash
' VBScript files to help deploy firmware updates in the field over USB when OTA isn't available
' @2023 HankLloydRight

baud="921600"
Set portList = GetComPorts()

Set oShell = CreateObject("WScript.Shell")
portnames = portList.Keys
found=false
for each pname in portnames
    Set portinfo = portList.item(pname)
    'wscript.echo "ESP32 module " & portinfo.Name &"found on COM port: " & pname'
    if (InStr(portinfo.Name, "CP210")>0 Or InStr(portinfo.Name, "CH340")>0 Or InStr(portinfo.Name, "CH9102")>0) Then
	    wscript.echo "ESP32 module found on COM port: " & pname & " The NVRAM will now be erased."
	    cmd="%COMSPEC% /c esptool --chip auto --baud "&baud&" --port "& pname & " erase_flash "
		wscript.echo cmd
	  	oShell.run   cmd &"& echo. & echo. & echo. & echo NVRAM Erase Complete & pause"
	  	found=true
	end if
Next

if found=false Then wscript.echo "No ESP32 com port found"
'
' For all the keys in an entity, see
'http://msdn.microsoft.com/en-us/library/windows/desktop/aa394353(v=vs.85).aspx
'

'
' listComPorts -- List all COM, even USB-to-serial-based onesports,
'                 along with other info about them
'
' Execute on the command line with:
'   cscript.exe //nologo listComPorts.vbs
'
' http://github.com/todbot/usbSearch
'
' 2012, Tod E. Kurt, https://todbot.com/blog/
' 2017  N. Teering: Fixed devicename with Null value
'
' core idea stolen from
' http://collectns.blogspot.com/2011/11/vbscript-for-detecting-usb-serial-com.html
' And this is fun, if not particularly useful:
' http://henryranch.net/software/jwmi-query-windows-wmi-from-java/
Function GetComPorts()
    set portList = CreateObject("Scripting.Dictionary")

    strComputer = "."
    set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    set colItems = objWMIService.ExecQuery ("Select * from Win32_PnPEntity")
    for each objItem in colItems
        If Not IsNull(objItem.Name) Then
        	set objRgx = CreateObject("vbScript.RegExp")
        	objRgx.Pattern = "COM[0-9]+"
	        Set objRegMatches = objRgx.Execute(objItem.Name)
	        if objRegMatches.Count = 1 Then  portList.Add objRegMatches.Item(0).Value, objItem
	    End if
    Next
    set GetComPorts = portList
End Function




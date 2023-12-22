' hanklloydright FlashESP32 -- Flash firmware and SPIFFS/LittleFS partitions
' VBScript files to help deploy firmware updates in the field over USB when OTA isn't available
' @2023 HankLloydRight

'these are the settings for my Platform.io projects -- the memory locations for your BIN files will likely vary
'try a "verbose" build in Platform.io to figure out the proper settings for your project.
'These values are also available in the Arduino IDE right before uploading
'The bootloader.bin, partitions.bin, and spiffs/littlefs.bin files should be in your project build folder
'The boot_app0.bin will be in the platform.io folder for your particular build environment.

flashfreq="40m"
baud="921600"
bootloc="0x1000"
bootloaderbin="bootloader.bin"
partloc="0x8000"
partbin="partitions.bin"
bootapploc="0xe000"
bootapp="boot_app0.bin"
firmwareloc="0x10000"
spiffloc="0x3B0000"
spiffbin="littlefs.bin"

Set portList = GetComPorts()
Set oShell = CreateObject("WScript.Shell")
portnames = portList.Keys
for each pname in portnames
    Set portinfo = portList.item(pname)
    if (InStr(portinfo.Name, "CP210")>0 Or InStr(portinfo.Name, "CH340")>0 Or InStr(portinfo.Name, "CH9102")>0) Then
	    'wscript.echo pname & " - " & portinfo.name'
	    wscript.echo "ESP32 module found on COM port: " & pname & " In the next dialog box, please select the firmware *.bin file you wish to upload to the ESP32 module."
	    filePath =Chr(34)& BrowseForFile()&Chr(34)
		if filePath = "" Then
		  wscript.echo "Operation canceled, quitting firmware update"
		  WScript.Quit
		Else
		    wscript.echo "BIN file selected  " & filePath & ". Please do not close the DOS window that will popup next. When the update is complete, press any key to close the update window."
		    cmd="%COMSPEC% /c esptool.exe --chip esp32 --port "& pname & " --baud "& baud & " --before default_reset --after hard_reset write_flash -z --flash_mode dio --flash_freq "&flashfreq&" --flash_size detect "&bootloc&" "&bootloaderbin&" "&partloc&" "&partbin&" "&bootapploc&" "&bootapp&" "&firmwareloc&" "& filePath & " & echo. & echo. & echo. & echo 'Firmware Update Complete. Uploading Test Patterns...' & %COMSPEC% /c esptool.exe --baud "& baud & " --port "& pname & " write_flash "& spiffloc & " "&spiffbin&" -u & echo Done. & pause "
   		    wscript.echo "Command to be executed: "& cmd
	  	    oShell.run cmd
		End if
	end if
Next
if filePath="" Then wscript.echo "No ESP32 com port found"


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



'--------------------------------------------------------------------------------------
Function BrowseForFile()
'@description: Browse for file dialog.
'@author: Jeremy England (SimplyCoded)
  BrowseForFile = CreateObject("WScript.Shell").Exec( _
    "mshta.exe ""about:<input type=file id=f>" & _
    "<script>resizeTo(0,0);f.click();new ActiveXObject('Scripting.FileSystemObject')" & _
    ".GetStandardStream(1).WriteLine(f.value);close();</script>""" _
  ).StdOut.ReadLine()
End Function
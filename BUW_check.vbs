vbsInterpreter = "cscript.exe"

'Function to check status of USB
function readFromRegistry (strRegistryKey, strDefault )
    Dim WSHShell, value

    On Error Resume Next
    Set WSHShell = CreateObject("WScript.Shell")
    value = WSHShell.RegRead( strRegistryKey )

    if value = "3"  then
        readFromRegistry= "Enabled"
    elseif value = "4"  then
        readFromRegistry="Disabled"
    else
		readFromRegistry="Unknown"
	end if

    set WSHShell = nothing
end function

Dim oWMI, WQL, Instances, Instance

'Get base WMI object, "." means computer name (local)
Set oWMI = GetObject("WINMGMTS:\\.\ROOT\StandardCimv2")

'Create a WMI query text 
WQL = "Select * from MSFT_NetAdapter"

'Get instances of MSFT_NetAdapter 
Set Instances = oWMI.ExecQuery(WQL)

bluetooth_status="Disabled"
wifi_status="Disabled"

For Each Instance In Instances 
  If InStr(Instance.Name, "Bluetooth") And Instance.State<>3 Then
	bluetooth_status="Enabled"
  End If
  
  If InStr(Instance.Name, "Wi-Fi")  And Instance.State<>3 Then
	wifi_status="Enabled"
  End If
  
Next

usb_status= readfromRegistry("HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR\Start", "4")
wscript.echo "USB is: " & usb_status & vbCrLf & "Wifi is: " & wifi_status & vbCrLf & "Bluetooth is: " & bluetooth_status
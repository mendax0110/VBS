' List Classic COM Class Settings


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_ClassicCOMClassSetting")

For Each objItem in colItems
    Wscript.Echo "Application ID: " & objItem.AppID
    Wscript.Echo "Component ID: " & objItem.ComponentId
    Wscript.Echo "Control: " & objItem.Control
    Wscript.Echo "Default Icon: " & objItem.DefaultIcon
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "In-process Server 32: " & objItem.InprocServer32
    Wscript.Echo "Insertable: " & objItem.Insertable
    Wscript.Echo "Java Class: " & objItem.JavaClass
    Wscript.Echo "ProgId: " & objItem.ProgId
    Wscript.Echo "Version Independent ProgId: " & _
        objItem.VersionIndependentProgId
    Wscript.Echo
Next

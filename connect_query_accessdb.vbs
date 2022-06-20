'Connect to Access Database and Query'

Set objConnection = CreateObject("ADODB.Connection")

Set shell = CreateObject("WScript.Shell")
folder = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\ACCESSDB\"

objConnection.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & "Database01.accdb"

Set ors = objConnection.Execute("SELECT * FROM table01")

Do While Not(ors.EOF)
    WScript.Echo ors("field01").Value
    ors.MoveNext
Loop

ors.Close


if WScript.Arguments.Item(0)="/QUERY_BASIC" Then

	on error resume next
    strQuery = "Select " + WScript.Arguments(1) + " from " + WScript.Arguments(2)
	Set objArray= GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(strQuery,,48)
    For each obj in objArray
		result = ","
		For each Prop in obj.Properties_
			result = result & Prop.Value & ","
        Next
		if NOT result = "," Then
			WScript.Echo result
		end if
    Next
	
ElseIf WScript.Arguments.Item(0)="/QUERY_ADVENCED" Then

	on error resume next
    strQuery = "Select " + WScript.Arguments(1) + " from " + WScript.Arguments(2) + " where " + WScript.Arguments(3)
	Set objArray= GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(strQuery,,48)
    For each obj in objArray
		result = ","
		For each Prop in obj.Properties_
			result = result & Prop.Value & ","
        Next
		if NOT result = "," Then
			WScript.Echo result
		end if
    Next
	
ElseIf WScript.Arguments.Item(0)="/ACTIVATE" Then
	
	' New methood Provided by abbodi1406
	on error resume next
	INSTANCE_ID="winmgmts:\\.\root\CIMV2:" + WScript.Arguments.Item(1) + ".ID='" + WScript.Arguments(2) + "'"
	GetObject(INSTANCE_ID).Activate()
	
	' To Err Is VBScript – Part 1
	' https://docs.microsoft.com/en-us/previous-versions/tn-archive/ee692852(v=technet.10)?redirectedfrom=MSDN
	
	WScript.Echo "Error:" & Err.Number
	Err.Clear
	
ElseIf WScript.Arguments.Item(0)="/DATA_FILE" Then
	
	' New methood Provided by abbodi1406
	on error resume next
	INSTANCE_ID="winmgmts:\\.\root\CIMV2:CIM_DataFile" + ".name='" + WScript.Arguments(1) + "'"
	WScript.Echo "," + GetObject(INSTANCE_ID).version

ElseIf WScript.Arguments.Item(0)="/UninstallProductKey" Then

	on error resume next
	strQuery = "Select * from " + WScript.Arguments(1) + " Where " + WScript.Arguments(2)
	Set objArray= GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(strQuery,,48)
	For each obj in objArray
		obj.UninstallProductKey()
	Next
	
ElseIf WScript.Arguments.Item(0)="/QUERY_INVOKE" Then
	
	' this is test methood
	' need to check.
	' how it work
	
	on error resume next
	strQuery = "Select * from " + WScript.Arguments(1) + " Where " + WScript.Arguments(2)
	Set objArray= GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(strQuery,,48)
	For each obj in objArray
		obj.ExecMethod(WScript.Arguments(3))
	Next
	
End If
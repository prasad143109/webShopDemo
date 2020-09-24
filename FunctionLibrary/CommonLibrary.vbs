'Option Explicit
Public strResultPath, strScriptPath, strLibPath, strAppLibPath, strReportLibPath, strTestDataPath, strObjRepPath, strCommonLibPath, strContPath, strBrowser , strURL
Public objDictBrowser,strBrowserType
Const objSyncwaitTime=10000,pageSyncwaitTime=60000
on error resume next
' ***********************************************************************************************
'
' 			C O M M O N   L I B R A R Y   F U N C T I O N S 
'
' ***********************************************************************************************

'1.killProcess
'2.Web_CloseSession
'3.booleanDBConnection
'4.booleanDBQuery
'5.booleanDBDisconnect

Function killProcess(byVal ProcessName)
	Dim objWMIService ,colProcess,objProcess
	on error resume next
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
	Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process Where Name = '" & ProcessName & "'")
	For Each objProcess in colProcess
		objProcess.Terminate()
		Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process Where Name = '" & ProcessName & "'")
		If  colProcess.count =0 Then
			Exit for
		End If
	Next
	set colProcess=nothing
    Set objWMIService=nothing
End Function

Function Web_CloseSession( )
	On Error Resume Next
	Set WshShell1 = CreateObject("WScript.Shell")
	WshShell1.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255", 1, True
	Set WshShell1=nothing
	KillProcess objDictBrowser.Item(strBrowserType)
End Function

Function launchBrowser(byVal strBrowser,byVal strURL)
	Dim mode_Maximized, mode_Minimized
	mode_Maximized = 3 'Open in maximized mode
	mode_Minimized = 2 'Open in minimized mode
	Web_CloseSession()
	SystemUtil.Run objDictBrowser.Item(strBrowser) ,strURL , , ,mode_Maximized
	
End Function

Public Function waittillbusy(oBrowser)
	Dim waittime
	waittime=now
	On Error resume next
	Do
		wait 1
	Loop while oBrowser.object.busy and datediff("s",now,waittime)<60
	If Err.number<>0 Then
		err.clear
		wait 8	
	End If	
End Function
Function login(byVal username,byVal password)
	On Error resume next
	IRIS_Login=False
	Browser("Demo Web Shop").Page("Demo Web Shop").Sync
	If trim(username)="" or trim(password)="" Then
		LogResult micFail,sTestCaseID, "Login DemoWebShop", "Please verify login credetials once, it should not be emply"
		Exit Function	
	ElseIf  verifyPatternMatch(password , pattern, true,false) Then
	End If
	screenName="Demo Web Shop Home"
	fieldName="Login Link"
	Set obj= Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("link_Login")
	If clickObject(sTestCaseID,screenname, obj,fieldName)<> true Then Exit function
	screenName="Demo Web shop Login"
	fieldName="user name"
	Set obj= Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebEdit("txt_Email")
	If setText(sTestCaseID,screenname, obj,username, fieldName)<> true Then Exit function
	fieldName="password"
	Set obj= Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebEdit("txt_Password")
	If setText(sTestCaseID,screenname, obj,crypt.Encrypt(password), fieldName)<> true Then Exit function
	fieldName="Login Button"
	Set obj= Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebButton("btn_Login")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	LogResult micPass,sTestCaseID, "Login DemoWebShop", "DemoWebShop Login is successful" 
	
End Function
Function logout()

 	fieldName="WebShop Home Image"
	Set obj= Browser("DemoWebShopHome").Page("DemoWebShopHome").Image("TricentisDemoWebShop")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	screenname="Demo WebShop Home"
	fieldName="WebShop Logout"
	Set obj= Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("link_Logout")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function

End Function
Function verifyPatternMatch(byVal strvalue , byVal pattern, byVal ignorecaseFlag,byVal globalFlag)
	on error resume next
	verifyPatternMatch=false
	set reg= new RegExp
	With reg
		.pattern=pattern
		.ignorecase=ignorecaseFlag
		.global=globalFlag
	End With
	If  reg.Test(strvalue) Then
		verifyPatternMatch=true
	End If
	
End Function
Function setText(TestCaseID,screenName,Obj,strValue,sFieldName)
	Dim ClassType, fieldType
	On Error resume next
	setText=False
	If Obj.exist(objSyncwaitTime) Then            
		ClassType= Obj.GetTOProperty("Class Name")
		FieldType=Trim(Obj.GetROProperty("type"))
		Select Case ClassType
			Case "WebEdit" 
			If lcase(FieldType)="password" Then
				Obj.setsecure strValue
			else
				Obj.set strValue
			End If
			
			Case "WebList"
			Obj.select strValue 
		End Select
	End If
	If err.number<>0 Then
        LogResult micFail, TestCaseID,"Enter '"&sFieldName&"' value", "error '"&err.description&"' occured for entering data for field '"&sFieldName&"'."
        err.clear                
    else
    	LogResult micPass, TestCaseID,"Enter '"&sFieldName&"' value", "user successfully enteredselected '"&sFieldName&"' value as "&strValue
        setText=True   
    End If
End Function
Function clickObject(TestCaseID,screenName,obj,sFieldName)
	Dim ClassType
	On Error resume next
	clickObject=False
	If obj.exist(objSyncwaitTime) Then            
		obj.click
	End If
	If err.number<>0 Then
        LogResult micFail, TestCaseID,"Click on object: '"&sFieldName&"'", "error '"&err.description&"' occured while clicking object '"&sFieldName&"'."
        err.clear               
    else
    	LogResult micPass, TestCaseID,"Click on object: '"&sFieldName&"'", "user successfully clicked on  '"&sFieldName&"' object"
        clickObject=True    
    End If
End Function
' ================================================================================================
'  NAME			: strFormatNowDate
'  DESCRIPTION 	  	: This function returns a string with date-time stamp
'  PARAMETERS		: nil
' ================================================================================================
Function strFormatNowDate()
	d=now()
	Dim arr(6)
	arr(0)="TestResults"
	arr(1)=DatePart("d",d)
	arr(2)=DatePart("m",d)
	arr(3)=DatePart("yyyy",d)
	arr(4)=DatePart("h",d) & "h"
	arr(5)=DatePart("n",d) & "m"
	arr(6)=DatePart("s",d)	& "s"
	strFormatNowDate= Join(arr,"_")
End Function


' ================================================================================================
'  NAME			: Initialization
'  DESCRIPTION 	  	: This function is used to create global variables which stores location 
'			   path of TestResult, TestData, Scripts, AppLib, Browser CommonLib & ObjectRepo 
'			  Loads common repository
'  PARAMETERS		: nil
' ================================================================================================

Public Function Initialization ()
'	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
'	Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")
'	For Each objProcess in colProcess
'		If LCase(objProcess.Name) = LCase("EXCEL.EXE") OR LCase(objProcess.Name) = LCase("EXCEL.EXE *32") Then
'        objProcess.Terminate()
'        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
'		End If
'	Next
'	For Each objProcess in colProcess
'		If LCase(objProcess.Name) = LCase("WerFault.exe") OR LCase(objProcess.Name) = LCase("WerFault.exe *32") Then
'        objProcess.Terminate()
'        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
'		End If
'	Next
	
	Set objDictBrowser= CreateObject("Scripting.dictionary")
	objDictBrowser.Add "IE","iexplore.exe"
	objDictBrowser.Add "CHROME","chrome.exe"
	objDictBrowser.Add "FIREFOX","firefox.exe"
	
	strBrowserType=Trim(ucase(strBrowserType))
	If objDictBrowser.Exists(Trim(ucase(strBrowserType))) Then
		strBrowserType=strBrowserType
	Else
		strBrowserType="CHROME"
	End If
	
	'Get the Directory of the framework
	SystemUtil.CloseProcessByName "EXCEL.EXE"
	sTestDir= Environment.Value ("TestDir")
	arrPath = Split (sTestDir, "\")

	'Save Result Path to variable strResultPath
	arrPath(UBound(arrPath)-1) = "Results"
	For I=0 to UBound(arrPath)-1
		If (I=0) Then
			strResultPath = arrPath(I)
		Else
			strResultPath = strResultPath + "\" + arrPath(I)
		End If
	Next
	strResultPath = strResultPath & "\"
	
 	Set objFso= createObject("Scripting.filesystemobject")  ' Script to delete existing results data
 	For each file in  objFso.GetFolder(strResultPath).Files
 		file.delete
 	Next
 	For each folder in  objFso.GetFolder(strResultPath).SubFolders
 		folder.delete
 	Next
 	
	'Save Script Path to variable strScriptPath
	arrPath(UBound(arrPath)-1) = "Scripts"
	For I=0 to UBound(arrPath) -1
		If (I=0) Then
			strScriptPath = arrPath(I)
		Else
			strScriptPath = strScriptPath + "\" + arrPath(I)
		End If
	Next
	strScriptPath = strScriptPath  & "\"
	
    	'Save Lib Path to variable sAppLibPath
	arrPath(UBound(arrPath)-1) = "Library"
	For I=0 to UBound(arrPath) -1
		If (I=0) Then
			strLibPath = arrPath(I)
		Else
			strLibPath = strLibPath + "\" + arrPath(I)
		End If
	Next
	
	'Save TestData Path to variable strTestDataPath
	arrPath(UBound(arrPath)-1) = "TestData"
	For I=0 to UBound(arrPath) -1
		If (I=0) Then
			strTestDataPath = arrPath(I)
		Else
			strTestDataPath = strTestDataPath + "\" + arrPath(I)
		End If
	Next
	strTestDataPath = strTestDataPath & "\"	
	
'	'Save ObjectRepository Path to variable strObjRepPath
'	arrPath(UBound(arrPath)-1) = "COR"
'	For I=0 to UBound(arrPath)-1
'		If (I=0) Then
'			strObjRepPath = arrPath(I)
'		Else
'			strObjRepPath = strObjRepPath + "\" + arrPath(I)
'		End If
'	Next
'	strObjRepPath1 = strObjRepPath & "\" & "CommonObjectRepository.tsr"
'	
'	'Loading the repository file
'	If  Instr(1, UCase(Environment.Value ("TestName")), "WS") = 0 Then
'		RepositoriesCollection.RemoveAll() 
'		RepositoriesCollection.Add(strObjRepPath1)  
'	End If
'	
'	'Loading the Controller File PathDriver Path
'	arrPath(UBound(arrPath)-1) = "Controller.xlsx"
'	For I=0 to UBound(arrPath)-1
'		If (I=0) Then
'			strContPath = arrPath(I)
'		Else
'			strContPath = strContPath + "\" + arrPath(I)
'		End If
'	Next
	
'	'Get the Browser from Controller File matching the test name
'	Set fsoTestSuiteController = CreateObject("Scripting.FileSystemObject")
'	Set sTestSuiteControllerFileSpec = fsoTestSuiteController.GetFile (strContPath)
'	Set objExcel = CreateObject("Excel.Application")
'	Set objWorkbook = objExcel.Workbooks.Open(sTestSuiteControllerFileSpec.Path)
'	intRow = 2	
'	Do Until objExcel.Cells(intRow,1).Value = ""
'		Set sAppController 	= objWorkbook.Sheets("Controller")
'		strTestname  = sAppController.Cells(intRow, 3).Value
'		If strTestname = Environment.Value ("TestName") Then
'			strBrowser  = UCASE(sAppController.Cells(intRow, 2).Value)
'                 If INSTR(1,UCASE(strBrowser),"IE") > 0 Then
'					 strBrowser = "IE"
'				 End If
'				 If INSTR(1,UCASE(strBrowser),"CHROME") > 0 Then
'					 strBrowser = "CHROME"
'				 End If
'				 If INSTR(1,UCASE(strBrowser),"FIREFOX") > 0 Then
'					 strBrowser = "FIREFOX"
'				 End If
'			Exit Do
'		End If
'		intRow = intRow +1 
'	Loop
'	objWorkbook.Close '  Close the excel report
'	objExcel.Quit
'	Set objExcel = Nothing
'	Set objWorkbook = Nothing
'	Set fsoTestSuiteController = nothing
'	Set sTestSuiteControllerFileSpec = nothing
'	If strBrowser = "" Then
'		strBrowser = "CHROME"
'	End If
'	systemutil.CloseProcessByName("EXCEL.EXE")
'	'-----------------------------------------------------'
'	'---------------Find The Test URL---------------------'
'	'-----------------------------------------------------'
'	
'	Set objExcel = CreateObject("Excel.Application") 
'	objExcel.Visible = True 
'	Set controllerExcel = objExcel.Workbooks.Open( strContPath ) 
'	Set controllerObj = controllerExcel.Sheets("Controller")
'	rowCountController = controllerObj.UsedRange.Rows.Count
'	Set environmentObj = controllerExcel.Sheets("Environment")
'	rowCountEnvironment = environmentObj.UsedRange.Rows.Count
'	testName = Environment.Value("TestName")
'	For contRow = 2 to rowCountController
'		If testName =  controllerObj.cells(contRow ,3).value Then
'			env = controllerObj.cells(contRow,7).value
'			For envRow = 2 to rowCountEnvironment
'				If env = environmentObj.cells(envRow, 1).value Then
'					strURL = environmentObj.cells(envRow, 2).value
'				End If
'			Next
'			Exit For
'		End If
'	Next
'	controllerExcel.close
'	'objExcel.Quit
'	Set objExcel = nothing
'	Set controllerExcel = nothing
'	Set controllerObj = nothing
'	Set environmentObj = nothing
'	'systemutil.CloseProcessByName("EXCEL.EXE")
'	'-----------------------------------------------------'
'
'	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
'	Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")
'	For Each objProcess in colProcess
'		If LCase(objProcess.Name) = LCase("EXCEL.EXE") OR LCase(objProcess.Name) = LCase("EXCEL.EXE *32") Then
'	        objProcess.Terminate()
'	        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
'		End If
'	Next
'	For Each objProcess in colProcess
'		If LCase(objProcess.Name) = LCase("WerFault.exe") OR LCase(objProcess.Name) = LCase("WerFault.exe *32") Then
'        objProcess.Terminate()
'        ' MsgBox "- ACTION: " & objProcess.Name & " terminated"
'		End If
'	Next

End Function










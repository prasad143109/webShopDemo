public sprevTestCaseID,sTestCaseID,screenShotOption
' ***********************************************************************************************
'
' 			R E P O R T   L I B R A R Y   F U N C T I O N S 
'
' ***********************************************************************************************
'
'1. CreateResultFile
'2. LogResult
'3. TestCaseExecutiveSummary
' ***********************************************************************************************

Set objBrowser = Description.Create
Set objWindow  = Description.Create
objBrowser("creationtime").value = "0"
objWindow("regexpwndtitle").value = "Microsoft Internet Explorer|Windows Internet Explorer|Mozilla Firefox"

Public  sResultFile, iStartTime, iErrImageNumber

iErrImageNumber = 1


' ================================================================================================
'  NAME			    : CreateResultFile
'  DESCRIPTION 	  	: This function creates HTML test script execution result. Result file name 
'			           created is same as the test script name. This function is called at the 
'			           begining of the test script
' ================================================================================================

Public Function CreateResultFile()
	Dim objfso, MyFile
	iStartTime = Now              'initialise starttime of TC execution
	Set objfso = CreateObject("Scripting.FileSystemObject")
	' Create Module folder if it doesn't exists
'	If (fso.FolderExists (strResultPath) = False) Then
'		Set fFolder = fso.CreateFolder (strResultPath)
'	End If
		
' Create Test Result file from TestName 
	sResultFile = strResultPath & Environment.Value("TestName") & "_" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "_" & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".html"
	Set MyFile = objfso.CreateTextFile(sResultFile,True)
	
	MyFile.WriteLine("<html><head><title>Test Script Execution Report</title></head>")
	MyFile.WriteLine("<body><table border='1' width='100%' bordercolorlight='#C0C0C0' cellspacing='0' cellpadding='0' bordercolordark='#C0C0C0' bordercolor='#C0C0C0' height='88'>")
	MyFile.WriteLine("<tr><td width='100%' colspan='4' height='28' bgcolor='#C0C0C0'><p align='center'><b><font face='Verdana' size='4' color='#000080'>Test Script Execution Report</font></b></td></tr>")
	MyFile.WriteLine("<tr><td width='16%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test Script Path</font></b></td>")
	MyFile.WriteLine("<td width='84%' height='25' colspan='3'><p style='margin-left: 5'><font face='Verdana' size='2'>")
	
	MyFile.WriteLine(Environment.Value("TestDir") & "</font></td></tr>")
	MyFile.WriteLine("<tr><td width='16%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test Case Name</font></b></td>")
	MyFile.WriteLine("<td width='84%' height='25' colspan='3'><font face='Verdana' size='2'>&nbsp;" & Environment.Value("TestName") & "</font></td></tr>	")
	
	MyFile.WriteLine("<tr><td width='16%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>Test URL</font></b></td>")
	MyFile.WriteLine("<td width='84%' height='25' colspan='3'><font face='Verdana' size='2'>&nbsp;"& strURL &"</font></td></tr>	")
	
	MyFile.WriteLine("</table>")
	MyFile.WriteLine("<p style='margin-left: 5'>&nbsp; </p>")
	
	MyFile.WriteLine("<table border='1' width='100%' bordercolordark='#C0C0C0' cellspacing='0' cellpadding='0' bordercolorlight='#C0C0C0' bordercolor='#C0C0C0' height='91'>")	
	
	'============== Result Column Header =====================
	MyFile.WriteLine("<tr><td width='5%' align='center' bgcolor='#000099' height='35'><b>")
	MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Sl. No.</font></b></td>")
	MyFile.WriteLine("<td width='10%' align='center' bgcolor='#000099' height='35'><b>")
	MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Testcase ID</font></b></td>")
	MyFile.WriteLine("<td width='35%' align='center' bgcolor='#000099' height='35'><b>")
	MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Test Step Description/Expected Result</font></b></td>")
	MyFile.WriteLine("<td width='40%' align='center' bgcolor='#000099' height='35'><b>")
	MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Test Step Actual Result</font></b></td>")
	MyFile.WriteLine("<td width='10%' bgcolor='#000099' height='35' align='center'><b>")
	MyFile.WriteLine("<font face='Verdana' size='2' color='#FFFFFF'>Test Step Status</font></b></td></tr>")
	
	strLogFile.WriteLine sResultFile & " report file created."	
' Close the file
	MyFile.Close
	CreateResultFile = resultfile
End Function

' ================================================================================================
'  NAME				: LogResult
'  DESCRIPTION 	  	: This function is used to write test step status to the HTML test result file.
'  PARAMETERS		: sStatus 	- Status of the step (micPass / micFail)
'					  sTestCaseID 	- Test case ID
'				  	  sTestStep 	- Test Step Name
'	       		      sDescription 	- Test Step Description
' ================================================================================================

Public Function LogResult(sStatus, sTestCaseID,sTestStep, sDescription)
	Const ForAppending = 8
	Const TristateUseDefault = -2
	Dim objfso, file, objTs
	strLogFile.WriteLine sTestCaseID &" , "& sTestStep &" , "&  sDescription
	Set objfso = CreateObject("Scripting.FileSystemObject")
	If (objfso.FolderExists (strResultPath & "ErrorSnapshot") = False) Then
		objfso.CreateFolder (strResultPath & "ErrorSnapshot")
	End If
	
	Set file = objfso.GetFile(sResultFile)
	Set objTs = file.OpenAsTextStream(ForAppending, TristateUseDefault)
	
	if trim(sprevTestCaseID)="" or ucase(trim(sprevTestCaseID))<> ucase(trim(sTestCaseID)) then 
		Environment.Value("TestStepNo")=0
		sprevTestCaseID=sTestCaseID
		objTs.WriteLine("<tr><td width='100%' colspan='5' height='28' bgcolor='#C0C0C0'><p align='center'><b><font face='Verdana' size='4' color='#000080'>"& sTCDescription &" testcase </font></b></td></tr>")
	End If
	Environment.Value("TestStepNo") = Environment.Value("TestStepNo") + 1
	objTs.WriteLine("<tr><td width='5%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & Environment.Value("TestStepNo") & "</font></td>")
	objTs.WriteLine("<td width='10%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & sTestCaseID & "</font></td>")
	objTs.WriteLine("<td width='35%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & sTestStep & "</font></td>")
	objTs.WriteLine("<td width='40%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>")
	
	' Replace vbcrlf in sDescription to <br>
	arrDesc = Split (sDescription, vbcrlf)
	
	For Each arrElement In arrDesc
		objTs.WriteLine(arrElement & "<br>")
	Next
	objTs.WriteLine("</font></td>")
    If sStatus = micPass OR sStatus = micDone Then
    	Reporter.ReportEvent micPass, sTestStep, sDescription
    	If (Ucase(Trim(screenShotOption)) = "ALWAYS" ) Then
			sImgRelativePath = "ErrorSnapshot\" & Environment.Value("TestName") & "_" & Month(Now) & "_" & Day(Now) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & "_PASS_STEP_" & Environment.Value("TestStepNo") & ".png"
			sErrImage1 = strResultPath & sImgRelativePath
			Desktop.CaptureBitmap sErrImage1, True		' Capture Desktop Snapshot.So that the browser will be displayed with its Address bar, which we miss by Browser().CaptureBitmap If image by name sErrImage already exist, then override		
			objTs.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#05A251'><a href='" & sImgRelativePath & "'>PASS </a></font></td></tr>")	    		
		Else
			objTs.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#05A251'> PASS </font></b></td></tr>")
		End If
        objTs.Close
	ElseIf sStatus = micFail Then
		Reporter.ReportEvent micFail, sTestStep, sDescription
		If (Browser(objBrowser).Exist(5)) Then
			If (Ucase(Trim(screenShotOption))= "ALWAYS" OR Ucase(Trim(screenShotOption))= "ONERROR") Then 
				sImgRelativePath = "ErrorSnapshot\" & Environment.Value("TestName") & "_" & Month(Now) & "_" & Day(Now) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & "_FAIL_STEP_" & Environment.Value("TestStepNo") & ".png"
				sErrImage1 = strResultPath & sImgRelativePath
				Desktop.CaptureBitmap sErrImage1, True		' Capture Desktop Snapshot.So that the browser will be displayed with its Address bar, which we miss by Browser().CaptureBitmap If image by name sErrImage already exist, then override
				objTs.WriteLine("<td width='10' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#FF0000'><a href='" & sImgRelativePath & "'>FAIL </a></font></td></tr>")
				iErrImageNumber = iErrImageNumber + 1
			Else
				objTs.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#FF0000'>FAIL</font></td></tr>")
			End If
		Else
			objTs.WriteLine("<td width='10%' height='25' align='center'><p style='margin-left: 5'><font face='Verdana' size='2' color='#FF0000'>FAIL</font></td></tr>")
		End If
	    objTs.Close
	End If
End Function

' ================================================================================================
'  NAME			: TestCaseExecutiveSummary
'  DESCRIPTION 	  	: This function is used to create test script executive summary. This function 
'					  is called from test script at the end of the test script
'  PARAMETERS		:
' ================================================================================================

Function TestCaseExecutiveSummary ()
	Const ForAppending = 8
	Const TristateUseDefault = -2
	Dim objfso, file, objTs
	Set objfso = CreateObject("Scripting.FileSystemObject")
	Set file = objfso.GetFile(sResultFile)
	Set objTs = file.OpenAsTextStream(ForAppending, TristateUseDefault)

	objTs.WriteLine("</table>")
	objTs.WriteLine("<p>&nbsp;</p>")
	objTs.WriteLine("<table border='1' width='52%' bordercolorlight='#C0C0C0' cellspacing='0' cellpadding='0' bordercolordark='#C0C0C0' bordercolor='#C0C0C0' height='88'>")
	objTs.WriteLine("<tr><td width='53%' colspan='2' height='28' bgcolor='#C0C0C0'><p align='center'><b><font face='Verdana' size='4' color='#000080'>")
	objTs.WriteLine("Test Script Execution Summary</font></b></td></tr>")
	
	objTs.WriteLine("<tr><td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2' >")
	objTs.WriteLine("Total Testcase Count</font></b></td>")
	
	objTs.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2' >&nbsp;" & Environment.Value("TotalTestcases") & "</font></td></tr>")
	objTs.WriteLine("<tr> <td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#05A251'>")
	objTs.WriteLine("Total Pass Count</font></b></td>")
	
	objTs.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'color='#05A251'>&nbsp;" & Environment.Value("TotalPassedTCs") & "</font></td></tr>")
	objTs.WriteLine("<tr> <td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2' color='#FF0000'>")
	objTs.WriteLine("Total Fail Count</font></b></td>")
	
	objTs.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'color='#FF0000'>&nbsp;" & Environment.Value("TotalFailedTCs") & "</font></td></tr>")
	objTs.WriteLine("<tr><td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>")
	objTs.WriteLine("Start Time</font></b></td>")
	objTs.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & iStartTime & "</font></td></tr>")
	objTs.WriteLine("<tr> <td width='17%' height='25'><p style='margin-left: 5'><b><font face='Verdana' size='2'>End Time</font></b></td>")
	objTs.WriteLine("<td width='36%' height='25'><p style='margin-left: 5'><font face='Verdana' size='2'>&nbsp;" & Now & "</font></td></tr></table>")
	objTs.Close
	
	objExcel.ActiveWorkbook.SaveAs strTestDataPath
	killProcess "excel.exe"
End Function 

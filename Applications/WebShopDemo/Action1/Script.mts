	
Public oEnvironment,oUrl,oUsername,oPassword
On error resume next
Set objfso= createObject("Scripting.fileSystemObject")
Environment.Value("TestStepNo")=0
Environment.Value("TotalTestcases")=0
Environment.Value("TotalPassedTCs")=0
Environment.Value("TotalFailedTCs")=0
Environment.Value("Applicationpath")=objfso.GetParentFolderName(Environment.Value("TestDir")) & "\"
Environment.Value("FrameworkPath")=objfso.GetParentFolderName(Environment.Value("Applicationpath")) & "\"
Environment.Value("FunLibraryPath")=Environment.Value("FrameworkPath") & "FunctionLibrary" & "\"
Environment.Value("ObjectRepositoryPath")=Environment.Value("FrameworkPath")  & "objectRepository" & "\"
Environment.Value("ResultPath")=Environment.Value("FrameworkPath") & "Results" & "\"
Environment.Value("TestDataPath")=Environment.Value("FrameworkPath") & "TestData" & "\"
Environment.Value("LogsPath")=Environment.Value("FrameworkPath") & "Logs" & "\"

'Load library files
set libraryFiles= objfso.GetFolder(Environment.Value("FunLibraryPath")).Files
If NOT ISObject(libraryFiles) Then
	ExitTest
End If
For each libFile in libraryFiles
	If Lcase(split(libFile.name,".")(1))="vbs" OR Lcase(split(libFile.name,".")(1))= "qfl" Then
		LoadFunctionLibrary libFile.path
	End If 
Next

'Load object repository
RepositoriesCollection.RemoveAll()
set objRepFiles= objfso.GetFolder(Environment.Value("ObjectRepositoryPath")).Files
If NOT ISObject(libraryFiles) Then
	ExitTest
End If
For each repFile in libraryFiles
	If Lcase(split(repFile.name,".")(1))="tsr" Then
		RepositoriesCollection.Add(repFile.path)  
	End If 
Next
' initialize the framework directory
call Initialization()
' create the Html report format
call CreateResultFile()

'get configsheet Fielditem Indexes
Set oEnvironment=oConfigSheet.range("1:1").find("Environment")
Set oUsername=oConfigSheet.range("1:1").find("username")
Set oPassword=oConfigSheet.range("1:1").find("password")
Set oUrl=oConfigSheet.range("1:1").find("URL")
Set oscreenShotOption=oConfigSheet.range("1:1").find("ScreenShotOption")
' verify the field column availability in configsheet
ColNames="Environment,username,password,URL,ScreenShotOption"
If HeaderVerify(oConfigSheet,ColNames)=false Then 
	strLogFile.WriteLine "Invalid columns in Input Sheet"&oExcelSheet.Name
	ExitTest
End If 

'get TCExecutionsheet Fielditem Indexes
Set oTCID=oTCExecSheet.range("1:1").find("TestcaseID")
Set oTCDescription=oTCExecSheet.range("1:1").find("TestcaseDescription")
Set oBrowser=oTCExecSheet.range("1:1").find("Browser")
Set oExecFlag=oTCExecSheet.range("1:1").find("ExecuteFlag")
Set oEnvironment=oTCExecSheet.range("1:1").find("Environment")
Set oStatus=oTCExecSheet.range("1:1").find("Status")
Set oComments=oTCExecSheet.range("1:1").find("Comments")
' verify the field column availability in TCExecutionsheet
ColNames="TestcaseID,TestcaseDescription,Browser,ExecuteFlag,Environment,Status,Comments"
If HeaderVerify(oTCExecSheet,ColNames)=false Then 
	strLogFile.WriteLine "Invalid columns in Input Sheet"&oTCExecSheet.Name
	ExitTest
End If
'get TCFieldsheet Fielditem Indexes
Set oTCID=oTCFieldSheet.range("1:1").find("TestcaseID")
Set oTCFieldItems=oTCFieldSheet.range("1:1").find("FieldItemNames")
' verify the field column availability in TCFieldsheet
ColNames="TestcaseID,FieldItemNames"
If HeaderVerify(oTCFieldSheet,ColNames)=false Then 
	strLogFile.WriteLine "Invalid columns in Input Sheet"&oTCFieldSheet.Name
	ExitTest
End If

sCommentCol=oComments.column
sStatusCol= oStatus.column

Set oDictFieldItem=CreateObject("Scripting.Dictionary")

For row = 2 to oTCExecSheet.usedrange.rows.count
	Environment.Value("Comments")=""
	Environment.Value("Status")=""
	comments=""
	If oTCExecSheet.cells(row,1).value="" Then
		Exit for
	End If
	strExecFlag=oTCExecSheet.cells(row,oExecFlag.column).value
	If Ucase(Trim(strExecFlag))="Y" Then
		oDictFieldItem.RemoveAll
		Environment.Value("TotalTestcases")=Environment.Value("TotalTestcases")+1
		For iLoop=1 to 1											' Loop used to exit current record if any error occurs
			sTestCaseID=oTCExecSheet.cells(row,oTCID.column).value
			sTCDescription=oTCExecSheet.cells(row,oTCDescription.column).value
			strBrowserType=oTCExecSheet.cells(row,oBrowser.column).value
			strEnvironment=oTCExecSheet.cells(row,oEnvironment.column).value
			strLogFile.WriteLine sTestCaseID &" testcase execution started"
		
			set cfEnvRow =oConfigSheet.range("A1:A"&oConfigSheet.usedrange.rows.count).find(Ucase(strEnvironment))
			If IsEmpty(cfEnvRow.Row) Then
				comments="Environment:-" & strEnvironment & " details are not present in the configSheet"
				strLogFile.WriteLine comments
				LogResult micFail,sTestCaseID, "Retrive Login Credentials", comments
				Environment.Value("Comments")=comments
				Environment.Value("Status")="FAILED"
				sTestCaseSetupFlag=False
			Else
				strURL=oConfigSheet.cells(cfEnvRow.row,oUrl.column).value
				screenShotOption=oConfigSheet.cells(cfEnvRow.row,oscreenShotOption.column).value
				Environment.Value("UID")=oConfigSheet.cells(cfEnvRow.row,oUsername.column).value	
				Environment.Value("PWD")=oConfigSheet.cells(cfEnvRow.row,oPassword.column).value
				
				set tcFieldRow =oTCFieldSheet.range("A1:A"&oTCFieldSheet.usedrange.rows.count).find(sTestCaseID)
				If IsEmpty(tcFieldRow.Row) Then
					comments="TestCase ID:-" & sTestCaseID & " is not present in the TestCaseFields inputsheet"
					strLogFile.WriteLine comments
					LogResult micFail,sTestCaseID, "Retrive testcase Input fieldItems", comments
					Environment.Value("Comments")=comments
					Environment.Value("Status")="FAILED"
					sTestCaseSetupFlag=False
				Else
					fieldItemString=oTCFieldSheet.cells(tcFieldRow.row,oTCFieldItems.column).value
					IF Trim(fieldItemString)<>"" Then 
						fieldItemArray=split(fieldItemString,",")	
						For itr = 0 To Ubound(fieldItemArray)
							Set tcInputField =oTestdataSheet.range("A1:A"&oTestdataSheet.usedrange.rows.count).find(fieldItemArray(itr))
							If IsEmpty(tcInputField.Row) Then
								comments="Field Item :-" & fieldItemArray(itr) & " is not present in the TestData inputsheet"
								strLogFile.WriteLine comments
								LogResult micFail,sTestCaseID, "Retrive testcase  fieldItem values", comments
								Environment.Value("Comments")=comments
								Environment.Value("Status")="FAILED"
								sTestCaseSetupFlag=False
								Exit for
							Else
								oDictFieldItem.Add fieldItemArray(itr), oTestdataSheet.cells(tcInputField.Row,2)
							End If 
						Next
						sTestCaseSetupFlag=True
					Else
						sTestCaseSetupFlag=True
					End If
				End If
			End If	
			
			If sTestCaseSetupFlag=True Then
				Select Case sTestCaseID
					Case "Demo001"
						Call initializeBrowser
						If launchBrowser(strURL)=False Then Exit for
						If applicationLogin(Environment.Value("UID"),Environment.Value("PWD"))=False Then Exit for
						Environment.Value("Status")="PASSED" 
					Case "Demo002"
						Call initializeBrowser
						Environment.Value("ExpectedResult") ="Login was unsuccessful. Please correct the errors and try again."
						Environment.Value("UID")=oDictFieldItem.Item("invalid_username")
						Environment.Value("PWD")=oDictFieldItem.Item("invalid_password")
						If launchBrowser(strURL)=False Then Exit for
						If applicationLogin(Environment.Value("UID"),Environment.Value("PWD"))=False Then 
							If Environment.Value("ExpectedResult")= Environment.Value("ErrorMessage") Then
							Environment.Value("Status")="PASSED"	
							End If
						End If	
					Case "Demo003"
						Call initializeBrowser
						Environment.Value("ExpectedResult")="Please enter a valid email address."
						Environment.Value("UID")=oDictFieldItem.Item("wrong_username")
						Environment.Value("PWD")=oDictFieldItem.Item("wrong_password")
						If launchBrowser(strURL)=False Then Exit for
						If applicationLogin(Environment.Value("UID"),Environment.Value("PWD"))=False Then 
							If Environment.Value("ExpectedResult")= Environment.Value("ErrorMessage") Then
							Environment.Value("Status")="PASSED"	
							End If
						End If	
					Case "Demo004"
						Call initializeBrowser
						Environment.Value("ExpectedResult")="Please verify login credentials once, it should not be emply"
						Environment.Value("UID")=""
						Environment.Value("PWD")=""
						If launchBrowser(strURL)=False Then Exit for
						If applicationLogin(Environment.Value("UID"),Environment.Value("PWD"))=False Then 
							If Environment.Value("ExpectedResult")= Environment.Value("ErrorMessage") Then
							Environment.Value("Status")="PASSED"	
							End If
						End If	
						
					Case "Demo005"
						Call initializeBrowser
						If launchBrowser(strURL)=False Then Exit for
						If applicationLogin(Environment.Value("UID"),Environment.Value("PWD"))=False Then Exit for
						If emptyShoppingCart()=False Then Exit for
						Environment.Value("Status")="PASSED" 
					Case "Demo006"
						Environment.Value("itemQuantity")=oDictFieldItem.Item("item_Quantity").value
						Environment.Value("FirstName")=oDictFieldItem.Item("Address_FirstName").value
						Environment.Value("LastName")=oDictFieldItem.Item("Address_LastName").value
						Environment.Value("Country")=oDictFieldItem.Item("Address_Country").value
						Environment.Value("City")=oDictFieldItem.Item("Address_City").value
						Environment.Value("Address1")=oDictFieldItem.Item("Address_Address1").value
						Environment.Value("ZIPCode")=oDictFieldItem.Item("Address_zipcode").value
						Environment.Value("PhoneNumber")=oDictFieldItem.Item("Address_PhoneNumber").value
						Environment.Value("PaymentMethod")=oDictFieldItem.Item("Address_PaymentMethod").value
						Environment.Value("ShippingMethod")=oDictFieldItem.Item("Address_ShippingMethod").value
						Call initializeBrowser
						If launchBrowser(strURL)=False Then Exit for
						If applicationLogin(Environment.Value("UID"),Environment.Value("PWD"))=False Then Exit for
						If emptyShoppingCart()=False Then Exit for
						If addItemToCartandVerifyDetails()=False Then Exit for
						If enterBillingAddressDetails(oDictFieldItem)=False Then Exit for
						Environment.Value("Status")="PASSED"
					Case else
						LogResult micFail,sTestCaseID, "TC should be available", "Provided testcase is not available, please check it once"
				End Select
			End If	
		Next
		If Environment.Value("Status")="PASSED" Then 
			Environment.Value("TotalPassedTCs")=Environment.Value("TotalPassedTCs")+1
		Else
			Environment.Value("Status")="FAILED"
			Environment.Value("TotalFailedTCs")=Environment.Value("TotalFailedTCs")+1
		End If 
		updateTestCaseStatus oTCExecSheet,row,sCommentCol,sStatusCol
		Call Web_CloseSession()
	End If
Next

' appends the execution summary
TestCaseExecutiveSummary()
 @@ hightlight id_;_Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebRadioGroup("paymentmethod")_;_script infofile_;_ZIP::ssf64.xml_;_

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

call Initialization()
call CreateResultFile()

Set oEnvironment=oConfigSheet.range("1:1").find("Environment")
Set oUsername=oConfigSheet.range("1:1").find("username")
Set oPassword=oConfigSheet.range("1:1").find("password")
Set oUrl=oConfigSheet.range("1:1").find("URL")
Set oscreenShotOption=oConfigSheet.range("1:1").find("ScreenShotOption")

ColNames="Environment,username,password,URL,ScreenShotOption"
If HeaderVerify(oConfigSheet,ColNames)=false Then 
	strLogFile.WriteLine "Invalid columns in Input Sheet"&oExcelSheet.Name
	ExitTest
End If 

Set oTCID=oTCExecSheet.range("1:1").find("TestcaseID")
Set oTCDescription=oTCExecSheet.range("1:1").find("TestcaseDescription")
Set oBrowser=oTCExecSheet.range("1:1").find("Browser")
Set oExecFlag=oTCExecSheet.range("1:1").find("ExecuteFlag")
Set oEnvironment=oTCExecSheet.range("1:1").find("Environment")
Set oStatus=oTCExecSheet.range("1:1").find("Status")
Set oComments=oTCExecSheet.range("1:1").find("Comments")
ColNames="TestcaseID,TestcaseDescription,Browser,ExecuteFlag,Environment,Status,Comments"
If HeaderVerify(oTCExecSheet,ColNames)=false Then 
	strLogFile.WriteLine "Invalid columns in Input Sheet"&oTCExecSheet.Name
	ExitTest
End If

Set oTCID=oTCFieldSheet.range("1:1").find("TestcaseID")
Set oTCFieldItems=oTCFieldSheet.range("1:1").find("FieldItemNames")
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
		For iLoop=1 to 1
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
				Environment.Value("TotalFailedTCs")=Environment.Value("TotalFailedTCs")+1
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
					Environment.Value("TotalFailedTCs")=Environment.Value("TotalFailedTCs")+1
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
								Environment.Value("TotalFailedTCs")=Environment.Value("TotalFailedTCs")+1
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
						Call Web_CloseSession()
						Environment.Value("Status")="PASSED" 
					Case "Demo002"
						Environment.Value("ExpectedResult") ="Login was unsuccessful. Please correct the errors and try again."
						Environment.Value("UID")=oDictFieldItem.Item("invalid_username")
						Environment.Value("PWD")=oDictFieldItem.Item("invalid_password")
						If launchBrowser(strURL)=False Then Exit for
						If applicationLogin(Environment.Value("UID"),Environment.Value("PWD"))=False Then 
							If Environment.Value("ExpectedResult")= Environment.Value("ErrorMessage") Then
							Environment.Value("Status")="PASSED"	
							End If
						End If	
						Web_CloseSession()
					Case "Demo003"
						Environment.Value("ExpectedResult")="Please enter a valid email address."
						Environment.Value("UID")=oDictFieldItem.Item("wrong_username")
						Environment.Value("PWD")=oDictFieldItem.Item("wrong_password")
						If launchBrowser(strURL)=False Then Exit for
						If applicationLogin(Environment.Value("UID"),Environment.Value("PWD"))=False Then 
							If Environment.Value("ExpectedResult")= Environment.Value("ErrorMessage") Then
							Environment.Value("Status")="PASSED"	
							End If
						End If	
						Web_CloseSession()	
					Case "Demo004"
						Environment.Value("ExpectedResult")="Please verify login credentials once, it should not be emply"
						Environment.Value("UID")=""
						Environment.Value("PWD")=""
						If launchBrowser(strURL)=False Then Exit for
						If applicationLogin(Environment.Value("UID"),Environment.Value("PWD"))=False Then 
							If Environment.Value("ExpectedResult")= Environment.Value("ErrorMessage") Then
							Environment.Value("Status")="PASSED"	
							End If
						End If	
						Web_CloseSession()
						
					Case "Demo006"
						Call launchBrowser(strBrowserType,strURL)
						Call login(Environment.Value("UID"),Environment.Value("PWD"))
						Environment.Value("Status")="PASSED"
						emptyShoppingCart()
						updateTestCaseStatus oTCExecSheet,row,sCommentCol,sStatusCol
						Web_CloseSession()
					Case "Demo004"
						Call launchBrowser(strBrowserType,strURL)
						Call login(Environment.Value("UID"),Environment.Value("PWD"))
						addItemToCartandVerifyDetails
						enterBillingAddressDetails
						Web_CloseSession()
					Case else
						LogResult micFail,sTestCaseID, "TC should be available", "Provided testcase is not available, please check it once"
				End Select
			Else
				Environment.Value("TotalFailedTCs")=Environment.Value("TotalFailedTCs")+1
			End If	
		Next
		If Environment.Value("Status")="PASSED" Then 
			Environment.Value("TotalPassedTCs")=Environment.Value("TotalPassedTCs")+1
		Else
			Environment.Value("Status")="FAILED"
			Environment.Value("TotalFailedTCs")=Environment.Value("TotalFailedTCs")+1
		End If 
		updateTestCaseStatus oTCExecSheet,row,sCommentCol,sStatusCol
	End If
Next
TestCaseExecutiveSummary()

'systemutil.Run "chrome.exe","www.google.com"
'Set oDesc= description.Create
'oDesc("xpath").value="//input[@type='submit']"
'Browser("title:=Google").Page("title:=Google").Sync
''oDesc("Class Name").value="WebButton"
'Set objcount= Browser("title:=Google").Page("title:=Google").ChildObjects(oDesc)
'For itr =0 to objcount.count-1
'
'objcount(itr).getroproperty("value")
'	objcount(itr).click
'	wait(5)
'Next
'msgbox objcount.count


'*************************************************
'	strvalue="atest@gmail.com"
'	on error resume next
'	verifyPatternMatch=false
'	set reg= new RegExp
'	With reg
'		.pattern="^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
'		.ignorecase=true
'		.global=false
'	End With
'	If  reg.Test(strvalue) Then
'		verifyPatternMatch=true
'	End If

'***********************************************

 'LogResult micFail, "Login application", "unsuccesfull"
' sTestCaseID="Login12345"
'Call launchBrowser(strBrowserType,strURL)
'data=array("atest@gmail.com","123456")
'Call login(data(0),data(1))
'Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
'Set oDesc= description.Create
'oDesc("xpath").value="//input[@value='Add to cart']"
'Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("tab_books").Click
'Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
''oDesc("Class Name").value="WebButton"
'Set objcount= Browser("DemoWebShopHome").Page("DemoWebShopHome").ChildObjects(oDesc)
'For itr =0 to objcount.count-1
'
'objcount(itr).getroproperty("value")
'	objcount(itr).click
'	wait(5)
'Next
'
'Browser("DemoWebShopHome").Page("DemoWebShopHome").ChildObjects()
'
'sTestCaseID="Logout12345"
'Call launchBrowser(strBrowserType,strURL)
'Call login("atest@gmail.com","123456")
'strAppLibPath = strLibPath & "\" & "AppLib.vbs"
'
'strReportLibPath = strLibPath & "\" & "ReportLib.vbs"
'
'
'Browser("DemoWebShop").Page("Welcome to The Hub").Sync
'Browser("DemoWebShop").Navigate "http://demowebshop.tricentis.com/" @@ hightlight id_;_1052750_;_script infofile_;_ZIP::ssf5.xml_;_
'Browser("DemoWebShop").Page("DemoWebShop").Link("link_Login").Click @@ script infofile_;_ZIP::ssf6.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Login").WebEdit("Email").Set "atest@gmail.com" @@ script infofile_;_ZIP::ssf7.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Login").WebEdit("Password").SetSecure "5f6966267bbd16d98f3724765835d9f36e25" @@ script infofile_;_ZIP::ssf8.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Login").WebCheckBox("RememberMe").Set "ON" @@ script infofile_;_ZIP::ssf9.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Login").WebCheckBox("RememberMe").Set "OFF" @@ script infofile_;_ZIP::ssf10.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Login").WebButton("Log in").Click @@ script infofile_;_ZIP::ssf11.xml_;_
'Browser("DemoWebShop").Page("DemoWebShop").Link("atest@gmail.com").Click @@ script infofile_;_ZIP::ssf12.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Account").Link("Shopping cart (0)").Click @@ script infofile_;_ZIP::ssf13.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").Link("Books").Click @@ script infofile_;_ZIP::ssf14.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Books").WebButton("Add to cart").Click @@ script infofile_;_ZIP::ssf15.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Books").Link("Shopping cart (1)").Click @@ script infofile_;_ZIP::ssf16.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebCheckBox("removefromcart").Set "ON" @@ script infofile_;_ZIP::ssf17.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Computing and Internet").Click @@ script infofile_;_ZIP::ssf18.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebButton("Update shopping cart").Click @@ script infofile_;_ZIP::ssf19.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Your Shopping Cart is").Click @@ script infofile_;_ZIP::ssf20.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Your Shopping Cart is").Click @@ script infofile_;_ZIP::ssf21.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Your Shopping Cart is").Click @@ script infofile_;_ZIP::ssf22.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").Link("Books").Click @@ script infofile_;_ZIP::ssf23.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Books").Image("Picture of Computing and").Click @@ script infofile_;_ZIP::ssf24.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebElement("Qty: $(document).ready(functio").Click @@ script infofile_;_ZIP::ssf25.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebEdit("addtocart_13.EnteredQuantity").Set "3" @@ script infofile_;_ZIP::ssf26.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebElement("Price: 10.00").Click @@ script infofile_;_ZIP::ssf27.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebButton("Add to cart").Click @@ script infofile_;_ZIP::ssf28.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebElement("Computing and Internet").Click @@ script infofile_;_ZIP::ssf29.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Computing").Link("Shopping cart (3)").Click @@ script infofile_;_ZIP::ssf30.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("30.00").Click @@ script infofile_;_ZIP::ssf31.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("30.00_2").Click @@ script infofile_;_ZIP::ssf32.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Sub-Total:").Click @@ script infofile_;_ZIP::ssf33.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Enter your destination").Click @@ script infofile_;_ZIP::ssf34.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebButton("Checkout").Click @@ script infofile_;_ZIP::ssf35.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebButton("close").Click @@ script infofile_;_ZIP::ssf36.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebCheckBox("termsofservice").Set "ON" @@ script infofile_;_ZIP::ssf37.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebButton("Checkout").Click @@ script infofile_;_ZIP::ssf38.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("checkout-step-billing").Click @@ script infofile_;_ZIP::ssf39.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebList("billing_address_id").Select "test test, test, test test, Algeria" @@ script infofile_;_ZIP::ssf40.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebList("billing_address_id").Select "New Address" @@ script infofile_;_ZIP::ssf41.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("Company:").Click @@ script infofile_;_ZIP::ssf42.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebList("billing_address_id").Select "test test, test, test test, Algeria" @@ script infofile_;_ZIP::ssf43.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue").Click @@ script infofile_;_ZIP::ssf44.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("checkout-step-shipping").Click @@ script infofile_;_ZIP::ssf45.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue_2").Click @@ script infofile_;_ZIP::ssf46.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebRadioGroup("shippingoption").Select "Next Day Air___Shipping.FixedRate" @@ script infofile_;_ZIP::ssf47.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue_3").Click @@ script infofile_;_ZIP::ssf48.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue_4").Click @@ script infofile_;_ZIP::ssf49.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue_5").Click @@ script infofile_;_ZIP::ssf50.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("atest dummy").Click @@ script infofile_;_ZIP::ssf51.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("Phone: 9974652536").Click @@ script infofile_;_ZIP::ssf52.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("Total: 30.00").Click @@ script infofile_;_ZIP::ssf53.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Confirm").Click @@ script infofile_;_ZIP::ssf54.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout_2").WebElement("Your order has been successful").Click @@ script infofile_;_ZIP::ssf55.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout_2").WebElement("Order number: 787185").Click @@ script infofile_;_ZIP::ssf56.xml_;_
'Browser("DemoWebShop").Page("Demo Web Shop. Checkout_2").WebButton("Continue").Click @@ script infofile_;_ZIP::ssf57.xml_;_
'Browser("DemoWebShop").Page("DemoWebShop").Link("Log out").Click @@ script infofile_;_ZIP::ssf58.xml_;_

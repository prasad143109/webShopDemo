'Option Explicit
Public strTestDir,strResultPath, strScriptPath, strFunLibraryPath,oTCExecSheet,oTCFieldSheet, strTestDataPath, strObjRepPath, strCommonLibPath, strContPath, strLogFile,strBrowser , strURL
Public objDictBrowser,strBrowserType ,objExcel,objWorkbook,oTestdataSheet,oConfigSheet
Public sTestCaseID ,sTCDescription,sEnvironment ,oDiscObj
Const objSyncwaitTime=10,pageSyncwaitTime=60 ,ForReading = 1, ForWriting = 2 , ForAppending=8
on error resume next
' ***********************************************************************************************
'
' 			C O M M O N   L I B R A R Y   F U N C T I O N S 
'
' ***********************************************************************************************

'1.Initialization
'2.releaseObects
'3.booleanDBConnection
'4.booleanDBQuery
'5.booleanDBDisconnect

' =========================================================================================================
'  NAME			: Initialization
'  DESCRIPTION 	  	: 1. This function is used to kill process before starting execution.
'					  2. Creates global variables which stores location path of strTestDir,strFunLibraryPath
'					  	 strObjRepPath,strResultPath,strLogPath,strTestDataPath,Browser
'  PARAMETERS		: nil
' =========================================================================================================

Public Function Initialization()
	Initialization=False
	On error resume next
	For itr = 1 To 1 
		killProcess "excel.exe"
		killProcess "iexplore.exe"
		killProcess "chrome.exe"
		killProcess "firefox.exe"
		
	' Defining browser types and set default browser type
		Set objDictBrowser= CreateObject("Scripting.dictionary")
		objDictBrowser.Add "IE","iexplore.exe"
		objDictBrowser.Add "CHROME","chrome.exe"
		objDictBrowser.Add "FIREFOX","firefox.exe"
		
		
	'Get the Directory of the framework
		Set objfso= createObject("Scripting.fileSystemObject")
		strTestDir= Environment.Value ("TestDir")
		strFunLibraryPath= Environment.Value("FunLibraryPath")
		strObjRepPath= Environment.Value("ObjectRepositoryPath")
		strResultPath=Environment.Value("ResultPath")
		strLogPath=Environment.Value("LogsPath")
		strTestDataPath= Environment.Value("TestDataPath") & "TestData.xlsx"
		
		
	' Script to delete existing results data
		If NOT objfso.FolderExists(strLogPath) Then
			objfso.CreateFolder(strResultPath)
		End If
		Set strLogFile= objfso.CreateTextFile(strLogPath&"Log.txt",ForWriting,True)
		If objfso.FolderExists(strResultPath) Then
			strLogFile.WriteLine strResultPath&" Folder exists."
			For each file in  objfso.GetFolder(strResultPath).Files
				file.delete
			Next
			For each folder in  objfso.GetFolder(strResultPath).SubFolders
				folder.delete
			Next
			If err.Number<>0 Then
				strLogFile.WriteLine err.description&" Error occured , hence Exiting from script."
				Exit for
			End If
		Else
			objfso.CreateFolder(strResultPath)
			strLogFile.WriteLine strResultPath&" Folder is not available, Hence creating new folder."
		End If
		
		
	'Get testdata file
		Set objExcel = CreateObject("Excel.Application")
		If objfso.FileExists(strTestDataPath) Then
			Set objWorkbook = objExcel.Workbooks.Open(strTestDataPath)
			Set oTestdataSheet = objWorkbook.Sheets("TestData")
			Set oConfigSheet = objWorkbook.Sheets("Config")
			Set oTCExecSheet = objWorkbook.Sheets("TestCaseExecution")
			Set oTCFieldSheet = objWorkbook.Sheets("TestCaseFields")
			objExcel.Visible =false
			objExcel.DisplayAlerts=false
			
			IF NOT IsObject(oTestdataSheet)Then
				strLogFile.WriteLine "TestData sheet is not exists, hence Exiting from script."
				Exit for
			End If
			IF NOT IsObject(oConfigSheet)Then
				strLogFile.WriteLine "Config sheet is not exists, hence Exiting from script."
				Exit for
			End If
			IF NOT IsObject(oTCExecSheet)Then
				strLogFile.WriteLine "TestExecution sheet is not exists, hence Exiting from script."
				Exit for
			End If
			IF NOT IsObject(oTCFieldSheet)Then
				strLogFile.WriteLine "TestCase Fields sheet is not exists, hence Exiting from script."
				Exit for
			End If
		Else
			strLogFile.WriteLine strTestDataPath&" TestData File is not exists, hence Exiting from script."
			Exit for
		End If
		Initialization=True
	Next
	If Err.Number<>0 Then
		Initialization=False
		Err.clear
		releaseObects()
		ExitTest
	End If
	
End Function

' =========================================================================================================
'  NAME			: releaseObects
'  DESCRIPTION 	  	: This function is used to release the memory for all memory created in test.
'  PARAMETERS		: 	 
' =========================================================================================================

Function releaseObects()
	Set strLogFile=Nothing
	Set objDictBrowser=Nothing
	set objfso= Nothing
	set objExcel= Nothing
	Set objWorkbook = Nothing
	Set oTestdataSheet = Nothing
	Set oConfigSheet = Nothing
	Set oTCExecSheet = Nothing
	Set oTCFieldSheet = Nothing
End Function

' =========================================================================================================
'  NAME			: initializeBrowser
'  DESCRIPTION 	  	: This function is used to define browser environment for Execution , Deafult is Chrome.
'  PARAMETERS		: nil
' =========================================================================================================
Sub initializeBrowser()
	If objDictBrowser.Exists(Trim(ucase(strBrowserType))) Then
		strBrowserType=strBrowserType
		strLogFile.WriteLine "Executing Script in "&strBrowserType&"."
	Else
		strLogFile.WriteLine "Browser "& strBrowserType &" not available , Hence Executing Script in CHROME."
		strBrowserType="CHROME"
	End If
End Sub

' =========================================================================================================
'NAME				: HeaderVerify
'DESCRIPTION 	  	: This function is used to verify the header fielditems in the inputsheet 
'PARAMETERS			: oExcelSheet - Excelsheet object
'					   ColNames  - columns to be verified
'RETURN 			: True if all present , else false.
' =========================================================================================================
Public Function HeaderVerify(oExcelSheet,ColNames)
	HeaderVerify= True
	If instr(ColNames,",")>0 then
    	arrTemp=split(ColNames,",")
    Else
    	arrTemp=Array(ColNames)
    end if
	With oExcelSheet
		Bcolfound=true         
		for itr =lbound(arrTemp) to ubound(arrTemp)
			set tempObj=.Range("1:1").Find(arrTemp(itr))
			if tempObj is nothing then
			Bcolfound=False
			exit for
			end if
		next
		If Bcolfound=False then
			comments = "Invalid columns in Input Sheet"&oExcelSheet.Name
			appendComments(comments)
			LogResult micFail,sTestCaseID, "verify the inputsheet field columns", comments
			HeaderVerify = False
		end if
	End With
End Function
' =========================================================================================================
'NAME				: killProcess
'DESCRIPTION 	  	: This function is used to kill the process from the taskmanager 
'PARAMETERS			: ProcessName 
' =========================================================================================================
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

' =========================================================================================================
'NAME				: Web_CloseSession
'DESCRIPTION 	  	: This function is used to close the browser set to the 'strBrowserType' variable
' =========================================================================================================
Function Web_CloseSession( )
	On Error Resume Next
	Set WshShell1 = CreateObject("WScript.Shell")
	WshShell1.Run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255", 1, True
	Set WshShell1=nothing
	KillProcess objDictBrowser.Item(strBrowserType)
End Function

' =========================================================================================================
'NAME				: Web_CloseSession
'DESCRIPTION 	  	: This function is used to close the browser set to the 'strBrowserType' variable
' =========================================================================================================
Function launchBrowser(byVal strURL)
	launchBrowser=False
	On error resume Next
	Dim mode_Maximized, mode_Minimized
	mode_Maximized = 3 'Open in maximized mode
	mode_Minimized = 2 'Open in minimized mode
	If Trim(strURL)="" Then
		comments="Provided url is Empty"
		LogResult micFail,sTestCaseID, "Launch DemoWebShop Url", comments
		appendComments(comments)
		Exit Function
	End If
	Web_CloseSession()
	SystemUtil.Run objDictBrowser.Item(strBrowserType) ,strURL , , ,mode_Maximized
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	If Browser("DemoWebShopHome").Page("DemoWebShopHome").Exist(pageSyncwaitTime) Then
		sPageTitle= Browser("DemoWebShopHome").Page("DemoWebShopHome").GetROProperty("title")
		If sPageTitle="Demo Web Shop" Then
			launchBrowser=True
			comments=sPageTitle&" page is displayed"
			LogResult micPass,sTestCaseID, "Launch DemoWebShop Url", comments
		Else
			comments=sPageTitle&" page is displayed"
			LogResult micFail,sTestCaseID, "Launch DemoWebShop Url", comments
		End If
	Else
		comments="Demo Web Shop page is not displayed"
		LogResult micFail,sTestCaseID, "Launch DemoWebShop Url", comments
		appendComments(comments)		
	End If
	logout()
	If Err.Number<> 0 Then
		launchBrowser=False
		comments="Error occured while launching url "&strURL&", error description is "&Err.Description
		Err.clear
		LogResult micFail,sTestCaseID, "Launch DemoWebShop Url", comments
		appendComments(comments)		
	End If
End Function

' =========================================================================================================
'NAME				: applicationLogin
'DESCRIPTION 	  	: This function is used to close the browser set to the 'strBrowserType' variable
' =========================================================================================================

Function applicationLogin(byVal username,byVal password)
	On Error resume next
	applicationLogin=False
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	
	pattern="^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"	' Patter to verify email format
	If Trim(username)="" OR Trim(password)="" Then
		Environment.Value("ErrorMessage")="Please verify login credentials once, it should not be emply"
		comments=Environment.Value("ErrorMessage")
		appendComments(comments)
		If Environment.Value("ExpectedResult")=Environment.Value("ErrorMessage") Then
			LogResult micPass,sTestCaseID, "Login DemoWebShop", comments
		Else
			LogResult micFail,sTestCaseID, "Login DemoWebShop", comments
		End If
		Exit Function	
	ElseIf NOT verifyPatternMatch(Trim(username) , pattern, true,false) Then
		Environment.Value("ErrorMessage")="Please enter a valid email address."
		comments=Environment.Value("ErrorMessage")
		appendComments(comments)
		If Environment.Value("ExpectedResult")=Environment.Value("ErrorMessage") Then
			LogResult micPass,sTestCaseID, "Login DemoWebShop", comments
		Else
			LogResult micFail,sTestCaseID, "Login DemoWebShop", comments
		End If
		Exit Function
	End If
	screenName="Demo WebShop Home"
	fieldName="Login Link"
	Set obj= Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("link_Login")
	If clickObject(sTestCaseID,screenname, obj,fieldName)<> true Then Exit function
	
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	sWelcomeNote=Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebElement("web_LoginWelcome").GetROProperty("innertext")
	
	If sWelcomeNote="Welcome, Please Sign In!" Then
		comments=sWelcomeNote& & " is verified successfully."
		appendComments(comments)
		LogResult micPass,sTestCaseID, "verify Login welcome page", comments
	Else
		comments=sWelcomeNote& & " verification is failed."
		appendComments(comments)
		LogResult micFail,sTestCaseID, "verify Login welcome page", comments
		Exit Function
	End If
	screenName="Demo Web shop Login"
	fieldName="username"
	Set obj= Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebEdit("txt_Email")
	If setText(sTestCaseID,screenname, obj,username, fieldName)<> true Then Exit function
	
	if(Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebElement("web_InvalidEmail").Exist(objSyncwaitTime)) Then 
		Environment.Value("ErrorMessage")=Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebElement("web_InvalidEmail").GetROProperty("innertext") 
		comments=Environment.Value("ErrorMessage")
		appendComments(comments)
		If Environment.Value("ExpectedResult")=Environment.Value("ErrorMessage") Then
			LogResult micPass,sTestCaseID, "Login DemoWebShop", comments
		Else
			LogResult micFail,sTestCaseID, "Login DemoWebShop", comments
		End If
		Exit Function
	End If
	fieldName="password"
	Set obj= Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebEdit("txt_Password")
	If setText(sTestCaseID,screenname, obj,crypt.Encrypt(password), fieldName)<> true Then Exit function
	fieldName="Login Button"
	Set obj= Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebButton("btn_Login")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	
	if(Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebElement("web_InvalidLogin").Exist(objSyncwaitTime)) Then 
		Environment.Value("ErrorMessage")=Browser("DemoWebShopLogin").Page("DemoWebShopLogin").WebElement("web_InvalidLogin").GetROProperty("innertext") 
		comments=Environment.Value("ErrorMessage")
		appendComments(comments)
		If Environment.Value("ExpectedResult")=Environment.Value("ErrorMessage") Then
			LogResult micPass,sTestCaseID, "Login DemoWebShop", comments
		Else
			LogResult micFail,sTestCaseID, "Login DemoWebShop", comments
		End If
		Exit Function
	End If
	sEmailID=Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("link_userID").GetROProperty("text")
	
	If lcase(sEmailID)=Lcase(Environment.Value("UID")) Then
		comments=sEmailID & " is verified successfully."
		appendComments(comments)
		LogResult micPass,sTestCaseID, "verify login user accountID", comments
		applicationLogin=True
	Else
		comments=sEmailID & "verification is failed."
		appendComments(comments)
		LogResult micFail,sTestCaseID, "verify login user accountID", comments
		Exit Function
	End If
	If Err.Number<> 0 Then
		applicationLogin=False
		comments="Error occured while logging into application "&strURL&", error description is "&Err.Description
		Err.clear
		LogResult micFail,sTestCaseID, "Login DemoWebShop Application", comments
		appendComments(comments)		
	End If
End Function
' =========================================================================================================
'NAME				: logout
'DESCRIPTION 	  	: This function is used to close the browser set to the 'strBrowserType' variable
' =========================================================================================================
Function logout()
	On Error Resume Next
 	screenname="Demo WebShop Home"
	If Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("link_Logout").Exist(objSyncwaitTime) Then
		fieldName="WebShop Logout"
		Set obj= Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("link_Logout")
		If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
		Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
		If Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("link_Login").Exist(objSyncwaitTime) Then
			comments="Application is logged out successfully."
			LogResult micPass,sTestCaseID, "Logout Application", comments
		Else
			comments="Application is not logged out successfully."
			LogResult micFail,sTestCaseID, "Logout Application", comments
			Environment.Value("Status")="FAILED"
		End If
	End If
End Function

Function verifyEmail()
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	value= Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("link_userID").GetROProperty("value")
	If Ucase(Trim(value))=Environment.Value("UID") Then
		LogResult micPass,sTestCaseID, "Email verification", "UserID "& value & " is verified successfully"
	Else
		LogResult micFail,sTestCaseID, "Email verification", "UserID "& Environment.Value("UID") & " verification failed. displayed userID value is "&value
	End If
End Function

Function emptyShoppingCart()
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	screenname="shopping Cart"
	fieldName="ShoppingCart Link"
	Set obj= Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("btn_Shoppingcart")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	ItemsinCart= split(split(obj.GetROProperty("innertext"),"(")(1),")")(0)
	If cint(ItemsinCart)=0 Then
		sMessage= Browser("b_demoWebShopShopping").Page("p_demoWebShopShopping").WebElement("Web_EmptyCartMessage").GetROProperty("innertext")
		comments="Message displayed " & sMessage
		If trim(sMessage)="Your Shopping Cart is empty!" Then
			LogResult micPass,sTestCaseID, "empty Shopping Cart", comments
		Else
			LogResult micFail,sTestCaseID, "empty Shopping Cart", comments
		End If 
		appendComments(comments)
	Else
		Set obj= Browser("b_demoWebShopShopping").Page("p_demoWebShopShopping").WebTable("tbl_shoppingCartTable")
		tbl_rows= obj.GetROProperty("rows")
		For itr = 2 to tbl_rows
			obj.ChildItem(itr,1,"WebCheckBox",0).set "ON"
		Next
		fieldName=" Update shopping Cart"
		Set obj= Browser("b_demoWebShopShopping").Page("p_demoWebShopShopping").WebButton("btn_updateShoppingCart")
		If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
		comments = tbl_rows-1& " item entries removed from the cart"
		appendComments(comments)
		sMessage= Browser("b_demoWebShopShopping").Page("p_demoWebShopShopping").WebElement("Web_EmptyCartMessage").GetROProperty("innertext")
		comments="Message displayed " & sMessage
		If trim(sMessage)="Your Shopping Cart is empty!" Then
			LogResult micPass,sTestCaseID, "empty Shopping Cart", comments
		Else
			LogResult micFail,sTestCaseID, "empty Shopping Cart", comments
		End If
		appendComments(comments)
	End If

End Function

Function addItemToCartandVerifyDetails()
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	screenname="webShop Home"
	fieldName="Books Tab"
	Set obj= Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("tab_books")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	
	Set objDesc= description.Create
	objDesc("xpath").value="//input[@value='Add to cart']/../../../preceding-sibling::div[@class='picture']//img"
	Set objcount= Browser("DemoWebShopHome").Page("DemoWebShopHome").ChildObjects(objDesc)
	If objcount.count>0 Then
		objcount(0).highlight
		Environment.Value("BookName")=Trim(split(objcount(0).getROProperty("title"),"Show details for")(1))
		screenname="webShop Books"
		fieldName =Environment.Value("BookName") & " Item Image Link"
		Set obj= objcount(itr)
		If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Else
		comments="Add to Cart buttons not enabled for the shopping Cart."
		LogResult micFail,sTestCaseID, "select Item from Cart", comments
		Exit function
	End If
	
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	cartItemName= Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebElement("web_BookName").getRoproperty("innertext")
	comments="Item: "&cartItemName&" retrieved from the itemCart"
	LogResult micPass,sTestCaseID, "retrive the ItemName added in Cart", comments
	
	If Lcase(Trim(cartItemName))=Lcase(Environment.Value("BookName")) Then
		comments="Item: "&cartItemName&" selected as per selection."
		LogResult micPass,sTestCaseID, "verify the Item in Cart", comments
	Else
		comments="Item: "&cartItemName&" not matched with as per selection."
		LogResult micFail,sTestCaseID, "verify the Item in Cart", comments
	End If
	
	itemPrice=Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebElement("web_OriginalPrice").GetROProperty("innertext")
	comments="Item Price: '"&itemPrice&"' retrieved from the itemCart"
	LogResult micPass,sTestCaseID, "retrive the Item Price added in Cart", comments
	
	Environment.Value("itemQuantity")=10
	itemQuantity=Environment.Value("itemQuantity")
	fieldName="Item Quantity"
	Set obj= Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebEdit("txt_itemQuantity")
	If setText(sTestCaseID,screenname, obj,itemQuantity, fieldName)<> true Then Exit function
	
	fieldName ="addTo Cart"
	Set obj= Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebButton("btn_addToCart")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	saddToCartMessage=Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebElement("web_saveCartMessage").GetROProperty("innertext")
	If instr(1,saddToCartMessage,"The product has been added to your shopping cart")>0 Then
		comments= saddToCartMessage& " Message is verfied successfully."
		LogResult micPass,sTestCaseID, "verify the addTo Cart save message", comments
	Else
		comments= saddToCartMessage& " Message is displayed and addTo Cart save message verification is failed."
		LogResult micFail,sTestCaseID, "verify the addTo Cart save message", comments
	End If
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("btn_Shoppingcart").Click
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	
'	Set obj=Browser("b_demoWebShopShopping").Page("p_demoWebShopShopping").WebTable("tbl_subTotal")
'	rowCount= obj.GetROProperty("rows")
'	For itr = 1 To rowCount
'	
'		 msgbox obj.GetCellData(itr,1)
'		 msgbox obj.GetCellData(itr,2)
'	Next
Browser("b_demoWebShopShopping").Page("p_demoWebShopShopping").WebCheckBox("chk_termsofservice").Set "ON"
Browser("b_demoWebShopShopping").Page("p_demoWebShopShopping").WebButton("btn_checkout").Click

'	Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebEdit("txt_itemQuantity").Set "10"
'	msgbox Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebElement("web_BookName").GetROProperty("innertext")
'	price =Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebElement("web_OriginalPrice").GetROProperty("innertext")
'	msgbox price
'	Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebButton("btn_addToCart").Click
'	Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").Sync
'	msgbox Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebElement("web_saveCartMessage").GetROProperty("innertext")

End Function

Function enterBillingAddressDetails()

	username="test"
	country="India"
	Environment.Value("FirstName")="firstname"
	Environment.Value("LastName")="Lastname"
	Environment.Value("Country")="India"
	Environment.Value("City")="Hyderabad"
	Environment.Value("Address1")="Hyderabad"
	Environment.Value("ZIPCode")="518401"
	Environment.Value("PhoneNumber")="1234567890"
	
	getFullAddress()
	screenName="Demo Web shop CheckOut"
	set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebList("LST_billingAddressID")
	If obj.Exist(5) Then
		fieldName="Billing Address"
		If instr(1,obj.GetROProperty("all items"),sFullAddress)>0 Then
			If setText(sTestCaseID,screenname, obj,sFullAddress, fieldName)<> true Then Exit function
		Else
			addAddress()
			If setText(sTestCaseID,screenname, obj,sFullAddress, fieldName)<> true Then Exit function
		End If
	Else
		comments= "Billing address List object is not available"
		LogResult micFail,sTestCaseID, "Select the New Address value from List", comments
	End If
	
	

'
End Function
sub getFullAddress()
	sFullAddress= Environment.Value("FirstName")&" "& _
	Environment.Value("LastName")&", "&Environment.Value("City")&", "& _
	Environment.Value("Address1")&" "&Environment.Value("ZIPCode")&", "&Environment.Value("Country")
End sub
Function addAddress()
	On error resume Next
	addAddress=false
	screenName="Demo WebShop CheckOut"
	set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebList("LST_billingAddressID")
	If obj.Exist(5) Then
		fieldName="Billing Address"
		If setText(sTestCaseID,screenname, obj,"New Address", fieldName)<> true Then Exit function
		Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").Sync
		
		fieldName="First Name"
		Set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebEdit("txt_billingFirstName")
		If setText(sTestCaseID,screenname, obj,Environment.Value("FirstName"), fieldName)<> true Then Exit function
		fieldName="Last Name"
		Set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebEdit("txt_billingLastName")
		If setText(sTestCaseID,screenname, obj,Environment.Value("LastName"), fieldName)<> true Then Exit function
		fieldName="Email ID"
		Set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebEdit("txt_emailID")
		If setText(sTestCaseID,screenname, obj,Environment.Value("UID"), fieldName)<> true Then Exit function
		fieldName="Country"
		Set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebList("LST_billingCountryID")
		If setText(sTestCaseID,screenname, obj,Environment.Value("Country"), fieldName)<> true Then Exit function
		fieldName="City"
		Set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebEdit("txt_City")
		If setText(sTestCaseID,screenname, obj,Environment.Value("City"), fieldName)<> true Then Exit function
		fieldName="Address1"
		Set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebEdit("txt_Address1")
		If setText(sTestCaseID,screenname, obj,Environment.Value("Address1"), fieldName)<> true Then Exit function
		fieldName="ZIP Code"
		Set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebEdit("txt_billingZIPCode")
		If setText(sTestCaseID,screenname, obj,Environment.Value("ZIPCode"), fieldName)<> true Then Exit function
		
		fieldName="Phone Number"
		Set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebEdit("txt_billingPhoneNumber")
		If setText(sTestCaseID,screenname, obj,Environment.Value("PhoneNumber"), fieldName)<> true Then Exit function
		fieldName="Continue"
		Set obj= Browser("Demo Web Shop. Checkout").Page("Demo Web Shop. Checkout").WebButton("BTN_Continue")
		If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
		
	Else
		comments= "Billing address List object is not available"
		LogResult micFail,sTestCaseID, "Select the New Address value from List", comments
		Exit Function
	End If
	If err.Number<>0 Then
		comments= "Error occured, could not be able to add address and Error description is:-"&Err.Description
		err.clear
		LogResult micFail,sTestCaseID, "Select the New Address value from List", comments
	End If
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

Function initializeCredentials(byVal environment)
	For row = 2 to oConfigSheet.rows.count
		If oConfigSheet.cells(row,1).value="" Then
			Exit for
		End If
		If Ucase(oConfigSheet.cells(row,oEnvironment.column))=Ucase(environment) then
			strURL=oConfigSheet.cells(row,oUrl.column).value
			Environment.Value("UID")=oConfigSheet.cells(row,oUsername.column).value	
			Environment.Value("PWD")=oConfigSheet.cells(row,oPassword.column).value
			Exit for
		End If
	Next
End Function



	
	
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
'  NAME			: updateTestCaseStatus
'  DESCRIPTION 	  	: This function used to update the final status and comments to testcase 
'  PARAMETERS		: testDataSheet - sheet reference
'					: recordRow    testcase row number
'					: commentCol   index of the comment column
'					: commentCol   index of the status column
' ================================================================================================
Function updateTestCaseStatus(byRef testDataSheet, byVal recordRow, ByVal commentCol,ByVal statusCol)
	testDataSheet.cells(recordRow,commentCol)=Environment.Value("Comments")
	testDataSheet.cells(recordRow,statusCol)=Ucase(Environment.Value("Status"))
End Function

' ================================================================================================
'  NAME			: updateTestCaseStatus
'  DESCRIPTION 	  	: This function used to update the final status and comments to testcase 
'  PARAMETERS		: comments - comments to be appended in the environment variable.
' ================================================================================================
Function appendComments(byVal comments)
	If Trim(Environment.Value("Comments"))="" Then
		Environment.Value("Comments")=comments
	Else
		Environment.Value("Comments")=Environment.Value("Comments") & vbcrlf & comments
	End If
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










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
' **********************************************************************************************

'1.	Initialization
'2.	releaseObects
'3.	initializeBrowser
'4.	HeaderVerify
'5.	killProcess
'6.	Web_CloseSession
'7.	launchBrowser
'8.	applicationLogin
'9.	logout
'10.	verifyEmail
'11.	emptyShoppingCart
'12.	addItemToCartandVerifyDetails
'13.	enterBillingAddressDetails
'14.	getFullAddress
'15.	addAddress
'16.	verifyPatternMatch
'17.	setText
'18.	clickObject
'19.	appendComments
'20.	updateTestCaseStatus
'
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
		processList=Array("excel.exe","iexplore.exe","chrome.exe","firefox.exe")
		For each process in processList
			killProcess process
		Next
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
	logout()
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
		comments=sWelcomeNote& " is verified successfully."
		appendComments(comments)
		LogResult micPass,sTestCaseID, "verify Login welcome page", comments
	Else
		comments=sWelcomeNote& " verification is failed."
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

' =========================================================================================================
'NAME				: emptyShoppingCart
'DESCRIPTION 	  	: This function is empty the cart and verify the empty items in cart Message.
' =========================================================================================================

Function emptyShoppingCart()
	On Error Resume Next
	emptyShoppingCart=False
	screenname="shopping Cart"
	fieldName="ShoppingCart Link"
	Set obj= Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("btn_Shoppingcart")
	ItemsinCart= split(split(obj.GetROProperty("innertext"),"(")(1),")")(0)
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopShopping").Page("DemoWebShopShopping").Sync
	
	If ItemsinCart="0" Then
		sMessage= Browser("DemoWebShopShopping").Page("DemoWebShopShopping").WebElement("Web_EmptyCartMessage").GetROProperty("innertext")
		comments="Message displayed " & sMessage
		If Trim(sMessage)="Your Shopping Cart is empty!" Then
			LogResult micPass,sTestCaseID, "empty Shopping Cart", comments
			emptyShoppingCart=True
		Else
			LogResult micFail,sTestCaseID, "empty Shopping Cart", comments
			Environment.Value("Status")="FAILED"
		End If 
		appendComments(comments)
	Else
		Set obj= Browser("DemoWebShopShopping").Page("DemoWebShopShopping").WebTable("tbl_shoppingCartTable")
		tbl_rows= obj.GetROProperty("rows")
		For itr = 2 to tbl_rows
			obj.ChildItem(itr,1,"WebCheckBox",0).set "ON"
		Next
		fieldName=" Update shopping Cart"
		Set obj= Browser("DemoWebShopShopping").Page("DemoWebShopShopping").WebButton("btn_updateShoppingCart")
		If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
		Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
		comments = tbl_rows-1& " item entries removed from the cart"
		LogResult micPass,sTestCaseID, "empty Shopping Cart", comments
		sMessage= Browser("DemoWebShopShopping").Page("DemoWebShopShopping").WebElement("Web_EmptyCartMessage").GetROProperty("innertext")
		comments="Message displayed " & sMessage
		If trim(sMessage)="Your Shopping Cart is empty!" Then
			LogResult micPass,sTestCaseID, "empty Shopping Cart", comments
			emptyShoppingCart=True
		Else
			LogResult micFail,sTestCaseID, "empty Shopping Cart", comments
			Environment.Value("Status")="FAILED"
		End If
		appendComments(comments)
	End If
	If Err.Number<> 0 Then
		emptyShoppingCart=False
		Environment.Value("Status")="FAILED"
		comments="Error occured while emptying Items in the cart, error description is "&Err.Description
		Err.clear
		LogResult micFail,sTestCaseID, "empty Shopping Cart", comments
		appendComments(comments)		
	End If
End Function
' =========================================================================================================
'NAME				: addItemToCartandVerifyDetails
'DESCRIPTION 	  	: This function is used to add items to the cart and verify the pricing details.
' =========================================================================================================
Function addItemToCartandVerifyDetails()
	On Error Resume Next
	addItemToCartandVerifyDetails=False
	itemQuantity=Environment.Value("itemQuantity")
	if NOT isNumeric(itemQuantity) Then
		comments="Provided item quantity is not valid, Please enter valid Quantity."
		LogResult micFail,sTestCaseID, "Add items to Cart", comments
		appendComments(comments)
		Exit Function
	ElseIf (itemQuantity<1) Then
		comments="Please provide item quantity more than 0."
		LogResult micFail,sTestCaseID, "Add items to Cart", comments
		appendComments(comments)
		Exit Function
	End If 
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
		Randomize
		itemIndex=cint((objcount.count-1)*Rnd)
		objcount(itemIndex).highlight
		Environment.Value("BookName")=Trim(split(objcount(itemIndex).getROProperty("title"),"Show details for")(1))
		screenname="webShop Books"
		fieldName =Environment.Value("BookName") & " Item Image Link"
		Set obj= objcount(itemIndex)
		If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Else
		comments="Add to Cart buttons not enabled for the shopping Cart."
		LogResult micFail,sTestCaseID, "select Item from Cart", comments
		appendComments(comments)
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
	appendComments(comments)
	
	Environment.Value("itemQuantity")=itemQuantity
	fieldName="Item Quantity"
	Set obj= Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebEdit("txt_itemQuantity")
	If setText(sTestCaseID,screenname, obj,itemQuantity, fieldName)<> true Then Exit function
	
	fieldName ="addTo Cart"
	Set obj= Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebButton("btn_addToCart")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	
	itemTotal= itemQuantity*itemPrice
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	saddToCartMessage=Browser("DemoWebShopHomeItemCart").Page("DemoWebShopHomeItemCart").WebElement("web_saveCartMessage").GetROProperty("innertext")
	If instr(1,saddToCartMessage,"The product has been added to your shopping cart")>0 Then
		comments= saddToCartMessage& " Message is verfied successfully."
		LogResult micPass,sTestCaseID, "verify the addTo Cart save message", comments
	Else
		comments= saddToCartMessage& " Message is displayed and addTo Cart save message verification is failed."
		LogResult micFail,sTestCaseID, "verify the addTo Cart save message", comments
		Exit Function
	End If
	
	fieldName="ShoppingCart Link"
	Set obj= Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("btn_Shoppingcart")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	
	Set obj=Browser("DemoWebShopShopping").Page("DemoWebShopShopping").WebTable("tbl_subTotal")
	rowCount= obj.GetROProperty("rows")
	bsubTotalFound=False
	For itr = 1 To rowCount
		 If Trim(obj.GetCellData(itr,1))="Sub-Total:" Then
		 	subTotal= obj.GetCellData(itr,2)
		 	If cdbl(subTotal)= cdbl(itemTotal) Then
		 		comments= "subTotal: "&subTotal& " and ItemTotal:"&itemTotal&" are matched."
				LogResult micPass,sTestCaseID, "verify the item subtotal", comments
				appendComments(comments)
			Else
				comments= "subTotal: "&subTotal& " and ItemTotal:"&itemTotal&" are not matched."
				LogResult micFail,sTestCaseID, "verify the item subtotal", comments
				appendComments(comments)
				Exit Function
		 	End If
		 	bsubTotalFound=True
		 End If 
		 if bsubTotalFound=True Then Exit for
	Next
	If bsubTotalFound=False Then
		comments= "subTotal column not available in the table"
		LogResult micPass,sTestCaseID, "verify the item subtotal", comments
		appendComments(comments)
		Exit Function
	End If
	
	Set obj= Browser("DemoWebShopShopping").Page("DemoWebShopShopping").WebCheckBox("chk_termsofservice")
	If NOT obj.GetROProperty("checked") Then
		Browser("DemoWebShopShopping").Page("DemoWebShopShopping").WebCheckBox("chk_termsofservice").Set "ON"	
	End If
	fieldName="Checkout"
	Set obj= Browser("DemoWebShopShopping").Page("DemoWebShopShopping").WebButton("btn_checkout")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
	addItemToCartandVerifyDetails=True
	
	If Err.Number<> 0 Then
		addItemToCartandVerifyDetails=False
		comments="Error occured while addItem to cart and verifyDetails, error description is "&Err.Description
		Err.clear
		LogResult micFail,sTestCaseID, "addItem to cart and verifyDetails", comments
		appendComments(comments)		
	End If
End Function
' =========================================================================================================
'NAME				: enterBillingAddressDetails
'DESCRIPTION 	  	: This function is used to Add new address if not available in the list.
'			  If address available then selects from the drop-downlist
'PARAMETRES		: Dictionary Object containing address related information.
' =========================================================================================================
Function enterBillingAddressDetails(byRef dictAddressObj)
	On error resume next
	enterBillingAddressDetails=False
	For Each key in  dictAddressObj.Keys
		If Trim(dictAddressObj.item(key))=""  Then
			comments= "Billing Address data fields should not be empty, Please check data."
			LogResult micFail,sTestCaseID, "verify the Billing Address data fields", comments
			appendComments(comments)
			Exit Function
		End If
	Next
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync
	sFullAddress=getFullAddress()
	screenName="Demo Web shop CheckOut"
	set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebList("LST_billingAddressID")
	If obj.Exist(pageSyncwaitTime) Then
		fieldName="Billing Address"
		If instr(1,obj.GetROProperty("all items"),sFullAddress)>0 Then
			If setText(sTestCaseID,screenname, obj,sFullAddress, fieldName)<> true Then Exit function
		Else
			If addAddress()=False Then Exit function
			If setText(sTestCaseID,screenname, obj,sFullAddress, fieldName)<> true Then Exit function
		End If
	Else
		comments= "Billing address List object is not available"
		LogResult micFail,sTestCaseID, "Select the New Address value from List", comments
		appendComments(comments)
		Exit Function
	End If
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync
	fieldName="Billing Address Continue"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebButton("BTN_BillingContinue")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync
	fieldName="Shipping Address"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebList("LST_shippingAddressID")
	If instr(1,obj.GetROProperty("all items"),sFullAddress)>0 Then
		If setText(sTestCaseID,screenname, obj,sFullAddress, fieldName)<> true Then Exit function
	Else
		comments= "Shipping address is not available in the List"
		LogResult micFail,sTestCaseID, "Select the Shipping Address value from List", comments
		appendComments(comments)
		Exit Function
	End If
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync
	fieldName="Shipping Address Continue"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebButton("BTN_ShippingAddressContinue")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync
	fieldName="Shipping Option"
	set obj=Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebRadioGroup("rad_ShippingOption")
	If instr(1,obj.GetROProperty("all items"),Environment.Value("ShippingMethod"))>0 Then
		If setText(sTestCaseID,screenname, obj,Environment.Value("ShippingMethod"), fieldName)<> true Then Exit function
	Else
		comments= "Provided shipping option is not available in the List."
		LogResult micFail,sTestCaseID, "Select the Shipping value from List", comments
		appendComments(comments)
		Exit Function
	End If
	fieldName="Shipping Method Continue"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebButton("BTN_ShippingMethodContinue")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync
	
	fieldName="Payment Option"
	set obj=Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebRadioGroup("rad_paymentmethod")
	If instr(1,obj.GetROProperty("all items"),Environment.Value("PaymentMethod"))>0 Then
		If setText(sTestCaseID,screenname, obj,Environment.Value("PaymentMethod"), fieldName)<> true Then Exit function
	Else
		comments= "Provided shipping option is not available in the List."
		LogResult micFail,sTestCaseID, "Select the Shipping value from List", comments
		appendComments(comments)
		Exit Function
	End If
	fieldName="Payment Method Continue"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebButton("BTN_PaymentMethodContinue")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync
	strPaymentInfo= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebElement("web_PaymentInfo").GetROProperty("innertext")
	If Trim(strPaymentInfo)="You will pay by COD" Then
		comments= "Paymentinformation verification is success. Payment info message is "&strPaymentInfo
		LogResult micPass,sTestCaseID, "Payment info verification", comments
		appendComments(comments)
	Else
		comments= "Paymentinformation verification is failed. Payment info message is "&strPaymentInfo
		LogResult micFail,sTestCaseID, "Payment info verification", comments
		appendComments(comments)
		Exit Function
	End If
	fieldName="Payment Info Continue"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebButton("BTN_PaymentInfoContinue")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync
	
	fieldName="Confirm Order"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebButton("BTN_ConfirmOrder")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync
	
	strOrderStatus= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebElement("web_orderStatus").GetROProperty("innertext")
	If Trim(strOrderStatus)="Your order has been successfully processed!" and Trim(strOrderStatus)<>"" Then
		comments= "Order status verification is success. Order status message is "&strOrderStatus
		LogResult micPass,sTestCaseID, "Order Status verification", comments
		appendComments(comments)
	Else
		comments= "Order Status verification is failed. Order Status message is "&strOrderStatus
		LogResult micFail,sTestCaseID, "Order Status verification", comments
		appendComments(comments)
		Exit Function
	End If
	
	strOrderNumber= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebElement("web_OrderNumber").GetROProperty("innertext")
	strOrderNumber= Trim(split(strOrderNumber,"Order number:")(1))
	If Trim(strOrderNumber)<>"" and isNumeric(strOrderNumber) Then
		comments= "Order Number generated. Order Number is "&strOrderNumber
		LogResult micPass,sTestCaseID, "verify the Order number", comments
		appendComments(comments)
	Else
		comments= "Order Number is not generated, "
		LogResult micFail,sTestCaseID, "verify the Order number", comments
		appendComments(comments)
		Exit Function
	End If
	fieldName="BTN_Continue"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebButton("BTN_Continue")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync

	enterBillingAddressDetails=True
	If Err.Number<> 0 Then
		enterBillingAddressDetails=False
		comments="Error occured while Process cart Order, error description is "&Err.Description
		Err.clear
		LogResult micFail,sTestCaseID, "Process cart Order", comments
		appendComments(comments)		
	End If
End Function
' =========================================================================================================
'NAME				: getFullAddress
'DESCRIPTION 	  	: This function is used to get addressList format to select the address.
' =========================================================================================================
Function getFullAddress()
	getFullAddress= Environment.Value("FirstName")&" "& _
	Environment.Value("LastName")&", "&Environment.Value("Address1")&", "& _
	Environment.Value("City")&" "&Environment.Value("ZIPCode")&", "&Environment.Value("Country")
End Function
' =========================================================================================================
'NAME				: addAddress
'DESCRIPTION 	  	: This function is used add the address record to select for shipping
' =========================================================================================================
Function addAddress()
	On error resume Next
	addAddress=false
	screenName="Demo WebShop CheckOut"
	set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebList("LST_billingAddressID")
	If NOT obj.Exist(pageSyncwaitTime) Then
		comments= "Billing address List object is not available"
		LogResult micFail,sTestCaseID, "Select the New Address value from List", comments
		appendComments(comments)
		Exit Function
	End If
	fieldName="Billing Address"
	If setText(sTestCaseID,screenname, obj,"New Address", fieldName)<> true Then Exit function
	Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").Sync

	fieldName="First Name"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebEdit("txt_billingFirstName")
	If setText(sTestCaseID,screenname, obj,Environment.Value("FirstName"), fieldName)<> true Then Exit function
	fieldName="Last Name"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebEdit("txt_billingLastName")
	If setText(sTestCaseID,screenname, obj,Environment.Value("LastName"), fieldName)<> true Then Exit function
	fieldName="Email ID"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebEdit("txt_emailID")
	If setText(sTestCaseID,screenname, obj,Environment.Value("UID"), fieldName)<> true Then Exit function
	fieldName="Country"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebList("LST_billingCountryID")
	If instr(1,obj.GetROProperty("all items"),Environment.Value("Country"))>0 Then
		If setText(sTestCaseID,screenname, obj,Environment.Value("Country"), fieldName)<> true Then Exit function
	Else
		comments= Environment.Value("Country")& " country is not  available in the List."
		LogResult micFail,sTestCaseID, "Select the country from List", comments
		appendComments(comments)
		Exit Function
	End If
	fieldName="City"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebEdit("txt_City")
	If setText(sTestCaseID,screenname, obj,Environment.Value("City"), fieldName)<> true Then Exit function
	fieldName="Address1"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebEdit("txt_Address1")
	If setText(sTestCaseID,screenname, obj,Environment.Value("Address1"), fieldName)<> true Then Exit function
	fieldName="ZIP Code"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebEdit("txt_billingZIPCode")
	If setText(sTestCaseID,screenname, obj,Environment.Value("ZIPCode"), fieldName)<> true Then Exit function
	
	fieldName="Phone Number"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebEdit("txt_billingPhoneNumber")
	If setText(sTestCaseID,screenname, obj,Environment.Value("PhoneNumber"), fieldName)<> true Then Exit function
	fieldName="Billing address Continue"
	Set obj= Browser("DemoWebShopCheckout").Page("DemoWebShopCheckout").WebButton("BTN_BillingContinue")
	If clickObject(sTestCaseID,screenname, obj, fieldName)<> true Then Exit function
	addAddress=True
	If err.Number<>0 Then
		addAddress=false
		comments= "Error occured, could not be able to add address and Error description is:-"&Err.Description
		err.clear
		LogResult micFail,sTestCaseID, "Add new address to the customer", comments
		appendComments(comments)
	End If
End Function
' =========================================================================================================
'NAME				: verifyPatternMatch
'DESCRIPTION 	  	: This function is used to verify the pattern match in given string
'PARAMETRES		: strvalue - string to be verified the pattern
'			  pattern - pattern which needs to be verified in the strvalue
'			  ignorecaseFlag - denotes casesensitivity
'			  globalFlag  - pattern search criteria in the strvalue
' =========================================================================================================
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
' =========================================================================================================
'NAME				: setText
'DESCRIPTION 	  	: This function is used to sets the provided strvalue to the text/List/radionbuttons.
'PARAMETRES		: TestCaseID - testcase id of the current TC id
'			  screenName- screenname where objects resides
'			  obj- actual object to be entered/select
'			  strvalue- value to be entered/select
'			  sFieldName- actual field name
' =========================================================================================================
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
			Case "WebList","WebRadioGroup"
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
' =========================================================================================================
'NAME				: clickObject
'DESCRIPTION 	  	: This function is used to clicks the provided object.
'PARAMETRES		: TestCaseID - testcase id of the current TC id
'			  screenName- screenname where objects resides
'			  obj- actual object to be clicked
'			  sFieldName- actual field name
' =========================================================================================================
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











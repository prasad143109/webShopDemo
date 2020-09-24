Public strURL,strBrowser
strURL="http://demowebshop.tricentis.com/"
strBrowser="CHROME"
strTestDir= Environment.Value("TestDir")
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
sTestDir = Environment.Value ("TestDir")  'Finding Script UFT file directory
arrPath = Split(sTestDir, "\")
arrPath(UBound(arrPath)-1) = "FunctionLibrary1"
For I=0 to UBound(arrPath)-1
	If (I=0) Then
		strLibPath = arrPath(I)
	Else
		strLibPath = strLibPath + "\" + arrPath(I)
	End If
Next

strCommLibPath = strLibPath & "\" & "CommonLib1.vbs"
LoadFunctionLibrary strCommLibPath
strCommLibPath = strLibPath & "\" & "ReportLib1.vbs"
LoadFunctionLibrary strCommLibPath
call Initialization
'msgbox strBrowserType
 call CreateResultFile
 'LogResult micFail, "Login application", "unsuccesfull"
 sTestCaseID="Login12345"
Call launchBrowser(strBrowserType,strURL)
data=array("atest@gmail.com","123456")
Call login(data(0),data(1))
Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
Set oDesc= description.Create
oDesc("xpath").value="//input[@value='Add to cart']"
Browser("DemoWebShopHome").Page("DemoWebShopHome").Link("tab_books").Click
Browser("DemoWebShopHome").Page("DemoWebShopHome").Sync
'oDesc("Class Name").value="WebButton"
Set objcount= Browser("DemoWebShopHome").Page("DemoWebShopHome").ChildObjects(oDesc)
For itr =0 to objcount.count-1

objcount(itr).getroproperty("value")
	objcount(itr).click
	wait(5)
Next

Browser("DemoWebShopHome").Page("DemoWebShopHome").ChildObjects()

sTestCaseID="Logout12345"
Call launchBrowser(strBrowserType,strURL)
Call login("atest@gmail.com","123456")
strAppLibPath = strLibPath & "\" & "AppLib.vbs"

strReportLibPath = strLibPath & "\" & "ReportLib.vbs"


Browser("DemoWebShop").Page("Welcome to The Hub").Sync
Browser("DemoWebShop").Navigate "http://demowebshop.tricentis.com/" @@ hightlight id_;_1052750_;_script infofile_;_ZIP::ssf5.xml_;_
Browser("DemoWebShop").Page("DemoWebShop").Link("link_Login").Click @@ script infofile_;_ZIP::ssf6.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Login").WebEdit("Email").Set "atest@gmail.com" @@ script infofile_;_ZIP::ssf7.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Login").WebEdit("Password").SetSecure "5f6966267bbd16d98f3724765835d9f36e25" @@ script infofile_;_ZIP::ssf8.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Login").WebCheckBox("RememberMe").Set "ON" @@ script infofile_;_ZIP::ssf9.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Login").WebCheckBox("RememberMe").Set "OFF" @@ script infofile_;_ZIP::ssf10.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Login").WebButton("Log in").Click @@ script infofile_;_ZIP::ssf11.xml_;_
Browser("DemoWebShop").Page("DemoWebShop").Link("atest@gmail.com").Click @@ script infofile_;_ZIP::ssf12.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Account").Link("Shopping cart (0)").Click @@ script infofile_;_ZIP::ssf13.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").Link("Books").Click @@ script infofile_;_ZIP::ssf14.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Books").WebButton("Add to cart").Click @@ script infofile_;_ZIP::ssf15.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Books").Link("Shopping cart (1)").Click @@ script infofile_;_ZIP::ssf16.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebCheckBox("removefromcart").Set "ON" @@ script infofile_;_ZIP::ssf17.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Computing and Internet").Click @@ script infofile_;_ZIP::ssf18.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebButton("Update shopping cart").Click @@ script infofile_;_ZIP::ssf19.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Your Shopping Cart is").Click @@ script infofile_;_ZIP::ssf20.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Your Shopping Cart is").Click @@ script infofile_;_ZIP::ssf21.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Your Shopping Cart is").Click @@ script infofile_;_ZIP::ssf22.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").Link("Books").Click @@ script infofile_;_ZIP::ssf23.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Books").Image("Picture of Computing and").Click @@ script infofile_;_ZIP::ssf24.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebElement("Qty: $(document).ready(functio").Click @@ script infofile_;_ZIP::ssf25.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebEdit("addtocart_13.EnteredQuantity").Set "3" @@ script infofile_;_ZIP::ssf26.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebElement("Price: 10.00").Click @@ script infofile_;_ZIP::ssf27.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebButton("Add to cart").Click @@ script infofile_;_ZIP::ssf28.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Computing").WebElement("Computing and Internet").Click @@ script infofile_;_ZIP::ssf29.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Computing").Link("Shopping cart (3)").Click @@ script infofile_;_ZIP::ssf30.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("30.00").Click @@ script infofile_;_ZIP::ssf31.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("30.00_2").Click @@ script infofile_;_ZIP::ssf32.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Sub-Total:").Click @@ script infofile_;_ZIP::ssf33.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebElement("Enter your destination").Click @@ script infofile_;_ZIP::ssf34.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebButton("Checkout").Click @@ script infofile_;_ZIP::ssf35.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebButton("close").Click @@ script infofile_;_ZIP::ssf36.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebCheckBox("termsofservice").Set "ON" @@ script infofile_;_ZIP::ssf37.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Shopping").WebButton("Checkout").Click @@ script infofile_;_ZIP::ssf38.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("checkout-step-billing").Click @@ script infofile_;_ZIP::ssf39.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebList("billing_address_id").Select "test test, test, test test, Algeria" @@ script infofile_;_ZIP::ssf40.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebList("billing_address_id").Select "New Address" @@ script infofile_;_ZIP::ssf41.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("Company:").Click @@ script infofile_;_ZIP::ssf42.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebList("billing_address_id").Select "test test, test, test test, Algeria" @@ script infofile_;_ZIP::ssf43.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue").Click @@ script infofile_;_ZIP::ssf44.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("checkout-step-shipping").Click @@ script infofile_;_ZIP::ssf45.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue_2").Click @@ script infofile_;_ZIP::ssf46.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebRadioGroup("shippingoption").Select "Next Day Air___Shipping.FixedRate" @@ script infofile_;_ZIP::ssf47.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue_3").Click @@ script infofile_;_ZIP::ssf48.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue_4").Click @@ script infofile_;_ZIP::ssf49.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Continue_5").Click @@ script infofile_;_ZIP::ssf50.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("atest dummy").Click @@ script infofile_;_ZIP::ssf51.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("Phone: 9974652536").Click @@ script infofile_;_ZIP::ssf52.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebElement("Total: 30.00").Click @@ script infofile_;_ZIP::ssf53.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout").WebButton("Confirm").Click @@ script infofile_;_ZIP::ssf54.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout_2").WebElement("Your order has been successful").Click @@ script infofile_;_ZIP::ssf55.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout_2").WebElement("Order number: 787185").Click @@ script infofile_;_ZIP::ssf56.xml_;_
Browser("DemoWebShop").Page("Demo Web Shop. Checkout_2").WebButton("Continue").Click @@ script infofile_;_ZIP::ssf57.xml_;_
Browser("DemoWebShop").Page("DemoWebShop").Link("Log out").Click @@ script infofile_;_ZIP::ssf58.xml_;_

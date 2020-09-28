###################################################################
# Chace Gibson											          #
# 9/24/2020                                                       #
# pulls all users information from the address book in outlook    #
# any one can run this script and get the same outcome            #
###################################################################


function Get-AllUsers {
	[Microsoft.Office.Interop.Outlook.Application] $outlook = New-Object -ComObject Outlook.Application
	$users = $outlook.Session.GetGlobalAddressList().AddressEntries

	$heading = "DISPLAY NAME?FIRST NAME?LAST NAME?EMAIL?USER NAME?CITY?DEPARTMENT?TITLE?PHONE NUMBER?ZIP CODE?ADDRESS?MANAGER"
	$heading | Out-File -append AllUsers.csv


	foreach($user in $users){
		$name = $user.Name
		$firstName = $user.GetExchangeUser().FirstName
		$lastName = $user.GetExchangeUser().LastName
		$email = $user.GetExchangeUser().PrimarySmtpAddress
		$userName = $user.GetExchangeUser().Alias
		$city = $user.GetExchangeUser().city
		$department = $user.GetExchangeUser().Department
		$title = $user.GetExchangeUser().JobTitle
		$phone = $user.GetExchangeUser().BusinessTelephoneNumber
		$zip = $user.GetExchangeUser().PostalCode
		$address = $user.GetExchangeUser().StreetAddress
		$manager = $user.GetExchangeUser().Manager.Name
		
		$row = $name + "?" + $firstName + "?" + $lastName + "?" + $email + "?" + $userName + "?" + $city + "?" + $department + "?" + $title + "?" + $phone + "?" + $zip + "?" + $address + "?" + $manager
		echo $row
		$row | Out-File -append AllUsers.csv
	}
}



echo "Make sure you update your address book in outlook"
echo "	Go to outlook-> send and receive tab-> send receive groups-> Download address book"
echo ""
echo "Once the program has ran it will create a CSV file and everything will be separated by question marks"
echo "In excel you can separate everything into cells by the question mark"
echo "To do this select column A-> click the data tab-> click text to columns-> delimited-> "
echo "Uncheck tab, check other and add a question mark-> finish"
echo ""
echo "Helpful formulas to use to reference data"
echo "The Match function will take the value of A1 in sheet AllUsers and look for it in row E returning the row number"
echo "=MATCH(A1,MainAllUsers!$ E$ 1:$ E$ 6000,0)"
echo "Index will return any value you want with a given row and column"
echo "=INDEX(MainAllUsers!$ A$ 1:$ E$ 6000,PREVIOUS MATCH FUNCTION VALUE,ROW NEEDING RETRIVED)"
echo ""
echo "PRESS ENTER TO GENERATE USER DATA"

$pause = Read-Host

Get-AllUsers

$pause = Read-Host



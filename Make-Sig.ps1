# Still need to create the  search and replace feature.
# Function SearchandRep($findText, $searchTerm, $document)

#Need to make a copy of the Docx template for later use.



$Username = ""
# Pull the user's properties from AD
$sigDetails = (Get-ADUser -identity $Username -Properties Name,Title,Company,StreetAddress,City,State,PostalCode,OfficePhone,Mobile,Fax)

#Set variables for the details required
$sigName = $sigDetails.Name
$sigTitle = $sigDetails.Title
$sigCompany = $sigDetails.Company
$sigStreet = $sigDetails.StreetAddress
$sigCity = $sigDetails.City
$sigState = $sigDetails.State
$sigPostCode = $sigDetails.PostalCode
$sigPhone = $sigDetails.OfficePhone
$sigMobile = $sigDetails.Mobile
$sigFax = $sigDetails.Fax
# $sigEmail = $sigDetails.EmailAddress
# $sigWeb = $sigDetails.wWWHomePage

# This is a Test output only Not needed for  the final script
Write-Output $sigName $sigTitle $sigCompany $sigStreet $sigCity $sigState $sigPostCode $sigPhone $sigMobile $sigFax



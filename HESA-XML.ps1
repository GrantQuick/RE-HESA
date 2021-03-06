#####################################################
# Parameters
# Taken from config.json
#####################################################

# RE exported csv file paths
$config = Get-Content ".\config.json" | ConvertFrom-Json
$bioPath = ($config | Select-Object -Property "bioPath").bioPath
$regNoPath = ($config | Select-Object -Property "regNoPath").regNoPath
$addressPath = ($config | Select-Object -Property "addressPath").addressPath
$emailPath = ($config | Select-Object -Property "emailPath").emailPath
$mobilePath = ($config | Select-Object -Property "mobilePath").mobilePath
$phonePath = ($config | Select-Object -Property "phonePath").phonePath
$husidPath = ($config | Select-Object -Property "husidPath").husidPath

# HESA coding files
$countryCodePath = ($config | Select-Object -Property "countryCodePath").countryCodePath

# XML output file
$generatedFile = ($config | Select-Object -Property "generatedFile").generatedFile

# RECID (17071, unlikely to change)
$recIdValue = ($config | Select-Object -Property "recIdValue").recIdValue

# UKPRN
$ukPrnValue = ($config | Select-Object -Property "ukPrnValue").ukPrnValue

# Value should be one of:
# CODE      LABEL
# Dec       December to February    Cohort A
# Mar       March to May            Cohort B
# Jun       June to August          Cohort C
# Sep       September to November   Cohort D
$censusValue = ($config | Select-Object -Property "censusValue").censusValue

<#
Add any countries to the list in config.json for which the country name in
RE may correspond to, but not match exactly, the country name as
listed in the C17071 valid-entries.csv list of countries and codes
#>
$countryList = ($config | Select-Object -Property "countryList").countryList


#####################################################
# Function definitions
#####################################################

<#
Check the country field matches one of the
values in the country code file from HESA and
add the HESA country code to each address object.
#>
function Get-AddressAndCountryCodes([string]$addressPathP, [string]$countryCodePathP){

    $countryCodeFile = Import-Csv -Path $countryCodePathP
    $addressFile = Import-Csv -Path $addressPathP

    $addresses = @()

    foreach ($address in $addressFile)
    {
        # Handle any blank addresses
        if($address.Country -eq '')
        {
            $address.Country = 'BLANK'
        }

        # Rename the countries as per the valid-entries.csv country codes
        if( $countryList.PSobject.Properties.Name -contains $($address.Country) )
        {
            $address.Country = $countryList.$($address.Country)
        }

        # Add a new property to the address object for the countrycode
        $address | Add-Member -NotePropertyName CountryCode -NotePropertyValue ''
        foreach ($code in $countryCodeFile)
        {
            if ($address.Country -eq $code.Label)
            {
                $address.CountryCode = $code.Code
                break > $null
            }
        }
        $addresses += $address
    }

    # Get the list of countries that do not have valid codes
    $testCountries = $addresses | Where-Object -Property CountryCode -eq '' | Select-Object -Property Country | Group-Object Country | Select-Object -Property @{N='CountryName';E={$_.Name}}, Count

    # Count them
    $unmatchedCountries = ($testCountries | Measure-Object -sum Count).sum

    Clear-Host

    # Check to see if the user wants to continue with the xml generation
    if($unmatchedCountries -gt 0)
    {

        write-host "Addresses without valid country codes:"
        Write-Host ($testCountries | Format-Table | Out-String)

        $confirmation = Read-Host "$unmatchedCountries address(es) cannot be matched to valid country codes. Do you wish to continue generating the xml? [y/n]"
        while($confirmation -ne "y")
        {
            if ($confirmation -eq 'n') {exit}
            $confirmation = Read-Host "Please press y or n [y/n]"
        }
    }
	return $addresses
}

# Test mobile/ukphone values
function Test-Phones ([object[]]$phoneFile,[string]$phoneType)
{
    $phones = @()

    foreach ($phoneNo in $phoneFile)
    {
        $phoneNo | Add-Member -NotePropertyName TestResult -NotePropertyValue ''
        $testResult = switch ( $phoneNo."Phone Number" )
        {
            {$($phoneNo."Phone Number") -notmatch '^\d{11}$' -and $phoneType -ne "International"} { "Not 11 digits long or contains invalid character" }
            {$($phoneNo."Phone Number") -notmatch '^\d{1,17}$' -and $phoneType -eq "International"} { "Not up to 17 digits long or contains invalid character" }
            {
                ($phoneType -eq "Mobile") -and 
                -not ($($phoneNo."Phone Number").StartsWith('071')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('072')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('073')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('074')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('075')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('07624')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('077')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('078')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('079'))
            } { "Not one of the valid mobile numbers as defined by HESA" }
            {
                ($phoneType -eq "Phone") -and 
                -not ($($phoneNo."Phone Number").StartsWith('01')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('02')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('03')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('04')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('05')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('06')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('08')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('09')) -and 
                -not ($($phoneNo."Phone Number").StartsWith('070'))
            } { "Not one of the valid phone numbers as defined by HESA" }
            {
                ($phoneType -eq "International") -and -not ($($phoneNo."Phone Number").StartsWith('00'))
            } { "International phone number does not start with 00" }
            default { "Pass" }
        }
        $phoneNo.TestResult = $testResult
        $phones += $phoneNo
    }

    # Get the list of phones that don't meet the HESA criteria
    $testPhones = $phones | Where-Object -Property TestResult -ne 'Pass' | Select-Object -Property TestResult | Group-Object TestResult | Select-Object -Property @{N='TestResult';E={$_.Name}}, Count
     # Count them
    $failedPhones = ($testPhones | Measure-Object -sum Count).sum

    Clear-Host

    # Check to see if the user wants to continue with the xml generation
    if($failedPhones -gt 0)
    {
        if ($phoneType -eq "Phone")
        {
            $headingString = "UKTEL(s) that will be rejected by HESA:"
            $getInputString = "$failedPhones UKTEL(s) do not meet the criteria specified by HESA. Do you wish to continue generating the xml? [y/n]"
        }
        if ($phoneType -eq "Mobile")
        {
            $headingString = "UKMOB(s) that will be rejected by HESA:"
            $getInputString = "$failedPhones UKMOB(s) do not meet the criteria specified by HESA. Do you wish to continue generating the xml? [y/n]"
        }
        if ($phoneType -eq "International")
        {
            $headingString = "INTTEL(s) that will be rejected by HESA:"
            $getInputString = "$failedPhones INTTEL(s) do not meet the criteria specified by HESA. Do you wish to continue generating the xml? [y/n]"
        }

        write-host $headingString
        Write-Host ($testPhones | Format-Table | Out-String)

        $confirmation = Read-Host $getInputString
        while($confirmation -ne "y")
        {
            if ($confirmation -eq 'n')
            {
                Write-Host "`n"
                Write-Host "Rejected numbers:"
                $phones | Where-Object -Property TestResult -ne 'Pass'| Select-Object -Property "Constituent ID","Phone Number","Phone Comments"
                exit
            }
            $confirmation = Read-Host "Please press y or n [y/n]"
        }
    }
    return $phones
}

# Add an XML sub-element
function Add-SubElement ($xmlDocument, $rootNode, $elementName, $elementValue)
{
    $xmlSubElt = $xmlDocument.CreateElement($elementName)
    $xmlSubText = $xmlDocument.CreateTextNode($elementValue)
    [void]$xmlSubElt.AppendChild($xmlSubText)
    [void]$rootNode.AppendChild($xmlSubElt)
}

# Get a list of matching phones
function Get-Phones ([object[]]$phoneArr,[string]$constituentID) {

    $valueList = @()
    foreach ($phoneItem in $phoneArr)
    {
        if ($($phoneItem."Constituent ID") -eq $constituentID)
        {
            $valueList += $($phoneItem."Phone Number")
        }
    }
    return $valueList
}

# Lose any spaces from the phone numbers (NXT automatically adds spaces to phone numbers)
function Remove-Spaces ($array_of_phone_objects)
{
    foreach($phone in $array_of_phone_objects){
        $phone.'Phone number' =  $($phone.'Phone number').replace(' ','')
    }
    return $array_of_phone_objects
}

#####################################################
# Let's go
#####################################################

# Test files
$testFail = $false
$filePaths = @{}
$filePaths.Add('Bio',$bioPath)
$filePaths.Add('Registration No.',$regNoPath)
$filePaths.Add('Addresses',$addressPath)
$filePaths.Add('Emails',$emailPath)
$filePaths.Add('Mobiles',$mobilePath)
$filePaths.Add('Phones',$phonePath)
$filePaths.Add('Country Codes',$countryCodePath)
$filePaths.Add('HUSIDs',$husidPath)

# Check the source files exist
foreach ($sourceFile in $filePaths.values)
{
    if (-not (Test-Path $sourceFile))
    {
        Write-Host "Please check the location for the" $($filePaths.GetEnumerator() | Where-Object { $_.Value -eq $sourceFile }).Key "csv is a valid file path"
        $testFail = $true
    }
}

if ($testFail)
{
    break  > $null
}

$addressesAndCodes = Get-AddressAndCountryCodes $addressPath $countryCodePath

# Filter out non-UK addresses and sort by preferred = 'Yes'
$ukAddresses = $addressesAndCodes |
    Where-Object {($_.CountryCode -eq 'XF')  `
    -or ($_.CountryCode -eq 'XH') `
    -or ($_.CountryCode -eq 'XI') `
    -or ($_.CountryCode -eq 'XG') `
    -or ($_.CountryCode -eq 'XK') `
    } | Sort-Object -Property "Constituent ID", Preferred -Descending

# Import the phone files
$mobileFile = Import-Csv -Path $mobilePath
$phoneFile = Import-Csv -Path $phonePath

$ukMobiles = $mobileFile | Where-Object -Property "Phone Comments" -ne "International"
$ukMobiles = Remove-Spaces ($ukMobiles)
$ukPhones = $phoneFile | Where-Object -Property "Phone Comments" -ne "International"
$ukPhones = Remove-Spaces ($ukPhones)
$intMobiles = $mobileFile | Where-Object -Property "Phone Comments" -eq "International"
$intLandlines = $phoneFile | Where-Object -Property "Phone Comments" -eq "International"
$internationalPhones = [array]$intMobiles + [array]$intLandlines
$internationalPhones = Remove-Spaces ($internationalPhones)

# Check and return valid phone numbers
$mobileNumbers = Test-Phones $ukMobiles "Mobile"
$phoneNumbers = Test-Phones $ukPhones "Phone"
$internationalNumbers = Test-Phones $internationalPhones "International"

# Import the remaining files
$bioFile = Import-Csv -Path $bioPath
$regNoFile = Import-Csv -Path $regNoPath
$emailFile = Import-Csv -Path $emailPath
$husidFile = Import-Csv -Path $husidPath

# Document creation
[xml]$xmlDoc = New-Object system.Xml.XmlDocument
[void]$xmlDoc.LoadXml("<?xml version=`"1.0`" encoding=`"utf-8`"?><GRADUATEOUTCOMESCONTACTDETAILSRecord></GRADUATEOUTCOMESCONTACTDETAILSRecord>")

# Create Provider Node
$xmlEltProv = $xmlDoc.CreateElement("Provider")

# RECID
Add-SubElement $xmlDoc $xmlEltProv "RECID" $recIdValue

# UKPRN
Add-SubElement $xmlDoc $xmlEltProv "UKPRN" $ukPrnValue

# CENSUS
Add-SubElement $xmlDoc $xmlEltProv "CENSUS" $censusValue

# Create Graduate Node and its sub elements
$i = 0
foreach ($bio in $bioFile)
{

    # Create the actual graduate node
    $xmlElt = $xmlDoc.CreateElement("Graduate")

    # UKMOB - Get a list of phones for this constituent
    $ukmobValues = Get-Phones $mobileNumbers $bio.'Constituent ID'

    # UKTEL - Get a list of mobiles for this constituent
    $uktelValues = Get-Phones $phoneNumbers $bio.'Constituent ID'

    # INTTEL - Get a list of mobiles for this constituent
    $inttelValues = Get-Phones $internationalNumbers $bio.'Constituent ID'

    # EMAIL - Get a list of emails for this constituent
    $emailValues = Get-Phones $emailFile $bio.'Constituent ID'

    # HUSID
    foreach ($husid in $husidFile)
    {
        if ($($husid."Constituent ID") -eq $($bio.'Constituent ID'))
        {
            Add-SubElement $xmlDoc $xmlElt "HUSID" $($husid.Alias)
            break > $null
        }
    }

    # # FEPUSID
    # Add-SubElement $xmlDoc $xmlElt "FEPUSID" $($bio."Constituent ID")

    # OWNSTU
    foreach ($regNo in $regNoFile)
    {
        if ($($regNo."Constituent ID") -eq $($bio.'Constituent ID'))
        {
            Add-SubElement $xmlDoc $xmlElt "OWNSTU" $($regNo.Alias)
            break > $null
        }
    }

    # COUNTRY
    $cCode = ''
    foreach ($countryCode in $addressesAndCodes)
    {
        if ($countryCode.Preferred -eq 'Yes' -and $($countryCode."Constituent ID") -eq $($bio.'Constituent ID'))
        {
            $cCode = $($countryCode.CountryCode)
            Add-SubElement $xmlDoc $xmlElt "COUNTRY" $cCode
            break  > $null
        }
    }

    # EMAIL - Add any values that were found for the current constituent
    foreach ($emailAddr in $emailValues)
    {
        Add-SubElement $xmlDoc $xmlElt "EMAIL" $emailAddr
    }

    # FNAMES
    Add-SubElement $xmlDoc $xmlElt "FNAMES" $($($bio.'First Name') + ' ' + $($bio.'Middle Name')).Trim()

    # FNMECHANGE
    # Not recorded in RE by Aberystwyth University
    # Add-SubElement $xmlDoc $xmlElt "FNMECHANGE" "TODO: Value for FNMECHANGE"

    # GRADSTATUS
    # Provide a value of 02 if there is no valid email or telephone record for a graduate
    $gradStatus = ''
    if ($bio.Deceased -eq 'Yes')
    {
        $gradStatus = '01'
        Add-SubElement $xmlDoc $xmlElt "GRADSTATUS" $gradStatus
    }
    elseif (($ukmobValues.Length -eq 0) -and ($uktelValues.Length -eq 0) -and ($inttelValues.Length -eq 0) -and ($emailValues.Length -eq 0) )
    {
        $gradStatus = '02'
        Add-SubElement $xmlDoc $xmlElt "GRADSTATUS" $gradStatus
    }

    # INTTEL - Add any values that were found for the current constituent
    foreach ($inttelValue in $inttelValues)
    {
        Add-SubElement $xmlDoc $xmlElt "INTTEL" $inttelValue
    }

    # SNAMECHNGE
    if ($bio.'Maiden Name' -ne '')
    {
        Add-SubElement $xmlDoc $xmlElt "SNAMECHNGE" $($bio.'Maiden Name')
    }

    # SURNAME
    Add-SubElement $xmlDoc $xmlElt "SURNAME" $($bio.Surname)

    # UKTEL - Add any values that were found for the current constituent
    foreach ($uktelValue in $uktelValues)
    {
        Add-SubElement $xmlDoc $xmlElt "UKTEL" $uktelValue
    }

    # UKMOB - Add any values that were found for the current constituent
    foreach ($ukmobValue in $ukmobValues)
    {
        Add-SubElement $xmlDoc $xmlElt "UKMOB" $ukmobValue
    }

    # PostalAddress
    # DOES NOT FILTER OUT INVALID/PREVIOUS ADDRESSES
    # Include all addresses. Enforces a maximum of 2 addresses as per HESA spec.
    # If more than 2 addresses exist for a constituent, the source file should
    # be edited manually to only include the 2 most pertinent addresses.

    if ($gradStatus -eq '02' )
    {
        $j = 1
        foreach ($ukAddr in $ukAddresses)
        {
            if ($j -gt 2) {continue}
            
            if ($($ukAddr."Constituent ID") -eq $($bio.'Constituent ID'))
            {
                $xmlEltPostal = $xmlDoc.CreateElement("PostalAddress")
                
                if ($($ukAddr."Address Line 1")) { Add-SubElement $xmlDoc $xmlEltPostal "ADDRESSLN1" $($ukAddr."Address Line 1") }
                if ($($ukAddr."Address Line 2")) { Add-SubElement $xmlDoc $xmlEltPostal "ADDRESSLN2" $($ukAddr."Address Line 2") }
                if ($($ukAddr."Address Line 3")) { Add-SubElement $xmlDoc $xmlEltPostal "ADDRESSLN3" $($ukAddr."Address Line 3") }
                if ($($ukAddr."Address Line 4")) { Add-SubElement $xmlDoc $xmlEltPostal "ADDRESSLN4" $($ukAddr."Address Line 4") }
                if ($($ukAddr."Address Line 5")) { Add-SubElement $xmlDoc $xmlEltPostal "ADDRESSLN5" $($ukAddr."Address Line 5") }
                if ($($ukAddr.City)) { Add-SubElement $xmlDoc $xmlEltPostal "ADDRESSLN6" $($ukAddr.City) }
                if ($($ukAddr.Postcode)) { Add-SubElement $xmlDoc $xmlEltPostal "POSTCODE" $($ukAddr.Postcode) }
                
                # Add the Postal Address to the Graduate node
                [void]$xmlElt.AppendChild($xmlEltPostal);
                $j++
            }
        }
    }

    # Add the Provider and all sub elements
    [void]$xmlEltProv.AppendChild($xmlElt)

    $i += 1
    # Clear-Host
    Write-Host "Processing record $i of $($bioFile.Count)"
    # Write-Progress -Activity "Generating XML" -status "Processing record $i" -percentComplete ($i / $($bioFile.Count)*100)
}


# Add the node to the document
[void]$xmlDoc.LastChild.AppendChild($xmlEltProv);

# Store to a file
$xmlDoc.Save($generatedFile)

if (Test-Path $generatedFile) {
    Write-Host "XML successfully output to $generatedFile"
}
else {
    Write-Host "There was an error creating the output file"
}
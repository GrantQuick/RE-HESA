## RE-HESA
A script for generating the XML for the C17071 HESA Graduate Outcomes submission in the schema defined by C17071.xsd, using data stored in Raiser's Edge.

## Getting Started
These instructions will describe how to extract the relevant data from Raiser's Edge (RE) and subsequently use the script to translate the data into the XML schema defined by HESA.

### Configuring PowerShell to run scripts
You will need to be able to run PowerShell scripts on your computer. You may initially need to alter the execution policy for your PowerShell installation in order to run scripts. To do this, start PowerShell as an administrator then run the following command:
```
Set-ExecutionPolicy RemoteSigned
```

### Script requirements
The script uses several files from data stored in RE. The data is extracted using the RE Query tool. The script needs the following files in order to run:
* A CSV containing biographical information (Bio file).
* A CSV containing the registration number of the student (Registration Number file).
* A CSV containing the HESA unique student identifier of the student (HUSID file).
* A CSV containing valid addresses for the graduates (Addresses file).
* A CSV containing valid email addresses for the graduates (Emails file).
* A CSV containing valid mobile phone numbers for the graduates (Mobiles file).
* A CSV containing valid landline numbers for the graduates (Phones file).
* The CSV of HESA valid country codes, downloadable from here: https://www.hesa.ac.uk/5272e752-eeca-4a78-8e51-f10da7363972

### Raiser's Edge source file requirements
The source files need to contain the following fields as a minimum:
* Bio file
    1. Constituent ID
    2. First Name
    3. Middle Name
    4. Surname
    5. Maiden Name
    6. Deceased
* Registration Number file
    1. Constituent ID
    2. Alias (This is the student registration number. In my organisation the student registration number is stored in RE as an Alias. If this is not the case in your organisation, the registration number should still be output into the csv with a column heading of Alias)
* HUSID file
    1. Constituent ID
    2. Alias (This is the value of the HESA unique student identifier. In my organisation the student registration number is stored in RE as an Alias. If this is not the case in your organisation, the HUSID should still be output into the csv with a column heading of Alias)
* Addresses file
    1. Constituent ID
    2. Address Line 1
    3. Address Line 2
    4. Address Line 3
    5. Address Line 4
    6. Address Line 5
    7. City
    8. Postcode
    9. Country
    10. Preferred
* Emails file
    1. Constituent ID
    2. Phone Type
    3. Phone Number
    4. Phone Inactive
    5. Phone Is Primary?
* Mobiles file
    1. Constituent ID
    2. Phone Type
    3. Phone Number
    4. Phone Inactive
    5. Phone Is Primary?
    6. Phone Comments
* Phones file
    1. Constituent ID
    2. Phone Type
    3. Phone Number
    4. Phone Inactive
    5. Phone Is Primary?
    6. Phone Comments

### Creating the source files
It is recommended that as part of the source file generation, an initial base query is created in RE to produce the dataset of constituents who should appear on any given HESA submission. This base query can be used as the source for each of the subsequent queries to ensure that each will return data on the same group of people. For example, you may wish to create a base query that returns all constituents from the class of 2017 and use that as the foundation for all other queries. To do this, when creating the other queries in RE, go to Tools > Query Options > Record Processing > Tick Select from query and choose your base query.

As an example, the following query definitions will produce the necessary output for use by the PowerShell script:

**Base Query**
- Query type: Constituent
- Query format: Dynamic
- Criteria: Primary Education Class of equals 2017

**Bio Query**
- Query type: Constituent
- Query format: Dynamic
- Select from: Base Query
- Output: Constituent ID, First Name, Middle Name, Surname, Maiden Name, Deceased

**Reg No Query**
- Query type: Constituent
- Query format: Dynamic
- Select from: Base Query
- Criteria: Alias Type equals Registration Number
- Output: Constituent ID, Alias Type, Alias

**HUSID Query**
- Query type: Constituent
- Query format: Dynamic
- Select from: Base Query
- Criteria: Alias Type equals HUSID
- Output: Constituent ID, Alias Type, Alias

**Address Query**
- Query type: Constituent
- Query format: Dynamic
- Select from: Base Query
- Output: Constituent ID, Address Line 1, Address Line 2, Address Line 3, Address Line 4, Address Line 5, City, Postcode, Country, Address Type, Preferred, Valid Date From, Valid Date To

**Emails Query**
- Query type: Constituent
- Query format: Dynamic
- Select from: Base Query
- Criteria: Phone Type equals Email AND Phone Number not blank
- Output: Constituent ID, Phone Type, Phone Number, Phone Inactive, Phone Is Primary?

**Mobiles Query**
- Query type: Constituent
- Query format: Dynamic
- Select from: Base Query
- Criteria: Phone Type equals Mobile AND Phone Number not blank
- Output: Constituent ID, Phone Type, Phone Number, Phone Inactive, Phone Is Primary?, Phone Comments

**Phones Query**
- Query type: Constituent
- Query format: Dynamic
- Select from: Base Query
- Criteria: Phone Type equals Phone AND Phone Number not blank
- Output: Constituent ID, Phone Type, Phone Number, Phone Inactive, Phone Is Primary?, Phone Comments

All queries with the exception of the base query will need to be saved as CSV files (by clicking the Export button in the query window) and stored locally in an appropriate folder.

### Configuring the PowerShell script
There are a number of parameters that will need to be configured at the beginning of the PowerShell script:
* $bioPath - set this to the full file path of your exported bio.csv file
* $regNoPath - set this to the full file path of your exported regNo.csv file
* $addressPath - set this to the full file path of your exported address.csv file
* $emailPath - set this to the full file path of your exported email.csv file
* $mobilePath - set this to the full file path of your exported mobile.csv file
* $phonePath - set this to the full file path of your exported phone.csv file
* $husidPath - set this to the full file path of your exported HUSID.csv file
* $countryCodePath - set this to the full file path of the CSV of HESA valid country codes, downloadable from here: https://www.hesa.ac.uk/5272e752-eeca-4a78-8e51-f10da7363972
* $generatedFile - set this to the full file path of the output file you would like to create
* $ukPrnValue - set this to the UK PRN value for your institution, available via https://www.ukrlp.co.uk/
* $censusValue - set this to the value of the submission period. Valid entries are available here: https://www.hesa.ac.uk/collection/c17071/a/census
* $countryList - add any countries to this list for which the country name in RE may correspond to, but not match exactly, the country name as listed in the C17071 valid-entries.csv list of countries and codes. 

### Running the PowerShell script
Run the script by:
1. Opening up a PowerShell terminal
2. Navigating to the folder containing the script
3. Entering the following
```
.\HESA-XML.ps1
```
The script will iterate through the constituent data in the csv files and produce an XML output file in the location you specified.

## Phone Numbers - Additional Information
In order to meet HESA submission requirements, phone numbers have to be classified as either UK Mobiles, UK Landlines or International Phones (either Mobiles or Landlines). In my organisation we have not historically identified phone numbers as being domestic or international. For the purposes of HESA submissions, for the most recent graduates we have begun to record this information within the Comments field for each international phone number. The Powershell script processes the Mobiles.csv and Phones.csv, and separates out any numbers that have the word **International** in the Comments field into INTTELs, while the remaining phones and mobiles (i.e. those without International in the Comments field) will go into UKTEL and UKMOB elements of the resulting XML.

## Known Issues
The script does not currently handle:
* Invalid data - some cursory checks on the data is performed by the script, but not all of HESAs requirements have bben implemented. It is assumed that manual validation of the data returned by the RE queries will be necessary in all cases.


## Authors
* **Grant Quick** - *Initial work* - [GrantQuick](https://github.com/GrantQuick)

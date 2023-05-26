# DefaultSig
A PowerShell script to enforce an Outlook signature using AD information.

This is a rehash of a script we found out on the Internet and updated.  The original can be found here:  http://www.danielclasson.com/powershell-script-to-set-outlook-signature-in-office-2010-and-office-2013-using-information-populated-from-active-directory/

Basically the script uses an MS Word doc as a template.  That document is copied to the users local profile and then a stream replace is performed using MS Word to swap out the target words in the template with data from Active Directory.  It then sets some registry keys to apply the signature to Outlook.

We use a .docx template, with all formatting in place as well as the company logo and confidentiality notices.  The stream replace keeps the formatting in place and only updates the words.  An example of the template header is as follows:

DisplayNameDesignation

Title
Company
Department
StreetAddress
City, STATECODE PostalCode
TelephoneNumber
FaxNumber
MobileNumber

I made a lot of changes to the original.  This version detects the PowerShell version and responds accordingly.  It also stores user data in an XML file to compare to current AD data in the event of a change.

We keep some information in the AD user record in the "extentionAttributes".  You should edit the script to conform to your AD configuration or you'll get blanks.

With this script there are numerous ways to trigger a refresh. 

Delete any of the files in the users profile
Select the "-force $true" command line option.
Change the script version number
The "-console $true" option shows all operation output.


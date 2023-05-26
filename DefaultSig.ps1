Param(
    [boolean]$Console = $False,                #--[ Set to true to enable local console result display. Defaults to false ]--
    [boolean]$Force = $False,                  #--[ Set to true to force a signature update ]--
    [boolean]$Remove = $False                  #--[ Set to true to clean all settings from local system ]--
)

<#==============================================================================
         File Name : DefaultSig.ps1
   Original Author : Daniel Classon
                   : http://www.danielclasson.com/powershell-script-to-set-outlook-signature-in-office-2010-and-office-2013-using-information-populated-from-active-directory/
                   :
       Description : Script to set Outlook 2010/2013 e-mail signature using Active Directory information
                   :
             Notes : This script will set the Outlook 2010/2013 e-mail signature on the local client
                   : using Active Directory information.  The template is created with a Word document,
                   : where images can be inserted and AD values can be provided.
                   :
         Operation : To force a refresh, change the version number below, delete ANY file on an individual user's profile,
                   : or use the "force" options on the command line or in the XML file.
                   :
          Warnings : None
                   :  
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources including but
                   :   not limited to the following:  Original Author Daniel Classon.
                   :
    Last Update by : Kenneth C. Mazie    
   Version History : v2.0 - 09-24-14 - Original
    Change History : v2.1 - 00-00-00 - Modified for use in environment by Andy Niel
                   : v3.0 - 10-16-17 - Refactored to compact script.  Added XML file, removed flag folders.  kmazie
                   : v3.1 - 10-28-17 - Fixed missing office and fax designations.  kmazie
                   : v3.2 - 10-30-17 - Fixed extra attribute display
                   : v3.3 - 10-31-17 - Disabled cell-phone display
                   : v3.4 - 11-01-17 - Minor tweak to correct cell-phone display (again...)
                   : v3.5 - 11-07-17 - Expanded file check to verify that all files exist.  Any missing file forces update.
                   : v4.0 - 11-17-17 - Major update.  Added tracking of user data to identify changes to user AD data.
                   : v4.1 - 11-20-17 - Fixed issues caused by Win7 defaulting to PS v2.
                   : #>               
                     $Version = "4.1"              <#--[ Denotes current version.  !!! CHANGING THIS NUMBER WILL TRIGGER A SIGNATURE UPDATE !!!
                   :
==============================================================================#>
clear-host

#--[ Misc Variables ]-----------------------------------------------------------
If (([string]$PSVersionTable.PSVersion.Major) -lt 4){
    $Domain = $Env:UserDNSDomain   #--[ For PowerShell v3 and below ]--
}Else{
    $Domain = (Get-ADDomain).DNSroot
}

#$Console = $true
#$Force = $true

$SigSource = "\\$Domain\netlogon\template.docx"
$OutlookPath = 'C:\Program Files (x86)\Microsoft Office\Office15\Outlook.exe'
$regkeypath = "HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings"
$keys = "NewSignature","ReplySignature"
$Date = $((Get-Date).ToString('yyyy-MM-dd'))
$AppData = $ENV:AppData
$SigPath = '\Microsoft\Signatures'
$LocalSignaturePath = $AppData + $SigPath

#--[ Determine local user, gather data from AD ]----------------------------------
$Script:UserName = $env:username
$Script:Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
$Script:Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Script:Searcher.Filter = $Filter
$Script:ADUserPath = $Searcher.FindOne()
$Script:ADUser = $ADUserPath.GetDirectoryEntry()
#$Script:ADEmailAddress = $ADUser.mail
#$Script:ADModify = $ADUser.whenChanged

#--[ Populate temporary array with users AD values ]------------------------------
$Script:ADAttributes = @()                                                                             
$Script:ADAttributes += ,@("DisplayName",$ADUser.DisplayName)
$Script:ADAttributes += ,@("Designation",$ADUser.extensionAttribute1)
$Script:ADAttributes += ,@("Title",$ADUser.title)
$Script:ADAttributes += ,@("Company",$ADUser.company)
$Script:ADAttributes += ,@("Department",$ADUser.extensionAttribute4)
$Script:ADAttributes += ,@("StreetAddress",$ADUser.extensionAttribute5)
$Script:ADAttributes += ,@("City",$ADUser.l)
$Script:ADAttributes += ,@("Statecode",$ADUser.st)
$Script:ADAttributes += ,@("PostalCode",$ADUser.PostalCode)
$Script:ADAttributes += ,@("TelephoneNumber",$ADUser.TelephoneNumber)
$Script:ADAttributes += ,@("MobileNumber",$ADUser.mobile)  
$Script:ADAttributes += ,@("FaxNumber",$ADUser.facsimileTelephoneNumber)
$Script:ADAttributes += ,@("MLS",$ADUser.extensionAttribute2)

$ExtraAttrib = @{
    "Designation" = ""
    "TelephoneNumber" = "Office: "  
    "MobileNumber" = "Mobile: "
    "FaxNumber" = "Fax: "
    "MLS" = "MLS:"
}

If ($Console){Write-Host "`n--[ Outlook Signature Script ]--`n" -ForegroundColor Cyan}

#--[ Check for Outlook to skip running on machines without MS Office ]-----------------------
If (!(Test-Path -path $OutlookPath)){exit}                                                         

#--[ Check for existence of all required files ]----------------------------------
$FileTypes = @("docx","htm","rtf","txt","xml")
ForEach ($Type in $FileTypes){
    If (!(Test-Path -Path "$LocalSignaturePath\$SignatureName.$Type" -PathType Leaf -ErrorAction SilentlyContinue)){
        If ($Console){Write-Host "--"($Type.ToUpper())"file was not detected.  Forcing signature file refresh --`n" -ForegroundColor Red}
        $Force = $true
    }   
}

#--[ Read XML data file in user profile ]---------------------------------
If (Test-Path -Path "$LocalSignaturePath\$SignatureName.xml" -PathType Leaf -ErrorAction SilentlyContinue){
    [XML]$XMLData = Get-Content -Path "$LocalSignaturePath\$SignatureName.xml"  -ErrorAction SilentlyContinue
    $MisMatch = $False
    $SignatureName = $XMLData.SigData.ScriptData.SigName
    If ($Console){
        Write-Host "Runtime values from current environment:" -ForegroundColor Cyan
       If (([string]$PSVersionTable.PSVersion.Major) -lt 4){
            Write-Host "PSVersion          = "$PSVersionTable.PSVersion" (PS Version is old.  Using v3 methods)" -ForegroundColor Red
        }Else{
            Write-Host "PSVersion          = "$PSVersionTable.PSVersion -ForegroundColor Green
        }
        Write-Host "Date               = "$Date -ForegroundColor Yellow
        Write-Host "Version            = "$Version  -ForegroundColor Yellow
        Write-Host "`nData read from existing XML file:" -ForegroundColor Cyan
        Write-Host "Date               = "$XMLData.SigData.ScriptData.Date -ForegroundColor Yellow
        Write-Host "Version            = "$XMLData.SigData.ScriptData.Version -ForegroundColor Yellow
        Write-Host "Remove             = "$XMLData.SigData.ScriptData.Remove -ForegroundColor Yellow
        Write-Host "Signature Name     = "$SignatureName -ForegroundColor Yellow
        Write-Host "`nUser data compare:     ActiveDirectory              Existing XML file:"  -ForegroundColor Cyan
    }
    foreach ($Item in $ADAttributes){
        write-host $Item[0].Padright(18," ")"= " -NoNewline -ForegroundColor Yellow
        If (!([string]::IsNullOrEmpty($XMLData.SigData.UserData.($Item[0])))){
            Write-Host (($Item[1]).Value).Padright(30," ") -ForegroundColor Yellow -NoNewline
        }Else{
            Write-Host "".Padright(30," ") -NoNewline
        }
        Write-Host $XMLData.SigData.UserData.($Item[0]) -NoNewline -ForegroundColor Magenta
       
        If (($Item[1]) -ne $XMLData.SigData.UserData.($Item[0])){
            $MisMatch = $True
            Write-Host "   MISMATCH" -ForegroundColor Red
            $Force = $true
        }Else{
            Write-Host ""
        }       
    }
   
    If ($Console){
        Write-Host "`nUser data comparison results:" -ForegroundColor Cyan
    }   
    If ($XMLData.SigData.ScriptData.Version -ne $Version){
        $Force = $true
        If ($Console){Write-Host "Script version     = MISMATCH" -ForegroundColor Red}
    }Else{
        If ($Console){Write-Host "Script version     = MATCH" -ForegroundColor Green}
    }       
    If ($XMLData.SigData.ScriptData.Remove -ne "false"){
        If ($Console){Write-Host "Remove Flag   = TRUE " -ForegroundColor Green}
        $Remove = $true
    }
    If ($MisMatch){
        Write-Host "User Data          = MISMATCH" -ForegroundColor Red
    }Else{
        Write-Host "User Data          = MATCH" -ForegroundColor Green
    }
}

#--[ Delete and recreate the signature files ]--
If ($Force){ 
    If ($Console){Write-Host "`n---[ Refreshing Signature Files ]---`n" -ForegroundColor White}
    remove-item -path $localSignaturePath\* -recurse -force | Out-Null                                #--[ Purge all existing local files ]--
      Copy-Item "$Sigsource" "$LocalSignaturePath\$signaturename.docx" -Recurse -Force  | Out-Null      #--[ Copy default template to local system ]--

    #--[ Insert AD user data ]--
      $MSWord = New-Object -ComObject word.application
      $MSWord.Visible = $false
      $fullPath = $LocalSignaturePath + '\'+$SignatureName + '.docx'
      $MSWord.Documents.Open($fullPath)  | Out-Null

    #--[ Create and/or populate the registry keys for original and reply emails ]----------------
      If ($Console){Write-Host "--- Updating registry files ---" -ForegroundColor White}
    If (test-path 'HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings\NewSignature') {
            Set-ItemProperty 'HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName | Out-Null
      }Else{
            New-ItemProperty 'HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force | Out-Null
      }

      If (test-path 'HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings\ReplySignature') {
            Set-ItemProperty 'HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName  | Out-Null
      }Else{
            New-ItemProperty 'HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force  | Out-Null
      }

    #--[ Perform a "in-memory" search and replace on the template ]-----------------------------
    If ($Console){Write-Host "--- Performing in-memory MS Word stream replace on template ---`n" -ForegroundColor White}
    If ($Console){Write-Host "--- New active user settings ---`nNote: Magenta items are from AD extention attributes." -ForegroundColor White}
      foreach ($Item in $ADAttributes){
        If ($ExtraAttrib.ContainsKey($Item[0])){  
            If ($Console){write-host ($Item[0].PadRight(18,' ')) -for  magenta -NoNewline }
            If (!([string]::IsNullOrEmpty($Item[1]))){ 
                If ($Console){
                    write-host " ="($ExtraAttrib.($Item[0])) -NoNewline -ForegroundColor red
                    write-host $Item[1] -foreground yellow
                } 
                If ($Item[0] -eq "MobileNumber"){  #--[ Forces listed attributes to be bypassed ]--
                    $MSWord.Selection.Find.Execute($Item[0], $False, $False, $False, $False, $False, $True, 1, $False, "", 2) | Out-Null
                }Else{
                    $MSWord.Selection.Find.Execute($Item[0], $False, $False, $False, $False, $False, $True, 1, $False, (($ExtraAttrib.($Item[0]))+$Item[1]), 2) | Out-Null
                }
            }Else{
                If ($Console){write-host " = " -for cyan }
                $MSWord.Selection.Find.Execute($Script:Item[0], $False, $False, $False, $False, $False, $True, 1, $False, "", 2) | Out-Null
            }
        }Else{   
            If ($Console){
                write-host (($Item[0]).PadRight(18,' ')) -foreground yellow -NoNewline       
                write-host " =" $Item[1] -foreground cyan
            }
            $MSWord.Selection.Find.Execute($Script:Item[0], $False, $False, $False, $False, $False, $True, 1, $False, $Script:Item[1].tostring(), 2) | Out-Null
        }   
            #--[ Search exectution format: $MSWord.Selection.Find.Execute($Script:Item[0],$MatchCase,$MatchWholeWord,$MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,$Wrap,$Format,$Script:Item[1].tostring(),$ReplaceAll)     
      }

      #--[ Save the modified template to the local system in multiple formats ]--------------------
      If ($Console){Write-Host "`n--- Writing new HTML signature file to local system ---" -ForegroundColor White}
    $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");          #--[ HTML format ]--
      $path = $LocalSignaturePath + '\'+$SignatureName + ".htm"
   
    If ($Console){Write-Host "--- Writing new RTF  signature file to local system ---" -ForegroundColor White}
      $MSWord.ActiveDocument.saveas([ref]$path, [ref]$saveFormat)                                        
      $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF");           #--[ RTF format ]--
      $path = $LocalSignaturePath + '\'+$SignatureName + ".rtf"
      $MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat)
   
    If ($Console){Write-Host "--- Writing new TEXT signature file to local system ---" -ForegroundColor White}
      $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText");          #--[ TXT format ]--
      $path = $LocalSignaturePath + '\'+$SignatureName + ".txt"
      $MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$SaveFormat)
      $MSWord.ActiveDocument.Close()
    $MSWord.Quit()                                                                                      #--[ Close the original template ]--

    #--[ Create a new XML config file ]----------------------------------------------------------
    If ($Console){Write-Host "`n--- Writing new XML file to local system ---`n" -ForegroundColor White}
    $XmlWriter = New-Object System.XMl.XmlTextWriter("$LocalSignaturePath\$SignatureName.xml",$Null)    #--[ Create the XML Document ]--
    $xmlWriter.Formatting = "Indented"                                                                  #--[ Set The Formatting ]--
    $xmlWriter.Indentation = "4"
    $xmlWriter.WriteStartDocument()                                                                     #--[ Write the XML Decleration ]--
    $XSLPropText = "type='text/xsl' href='style.xsl'"                                                   #--[ Set the XSL ]--
    $xmlWriter.WriteProcessingInstruction("xml-stylesheet", $XSLPropText)
    #--[ Script data for comparison at next check ]--
    $xmlWriter.WriteStartElement("SigData")                                                             #--[ Write the Root Element ]--
    $xmlWriter.WriteStartElement("ScriptData")
    $xmlWriter.WriteElementString("Remove",$False)                                                      #--[ Write the Data ]--
    $xmlWriter.WriteElementString("Version",$Version)
    $xmlWriter.WriteElementString("Date",$Date)
    $xmlWriter.WriteElementString("SigName",$SignatureName)
    $xmlWriter.WriteEndElement() | out-null                                                             #--[ Close ScriptDataElement ]--
    #--[ User data from AD for comparison at next check ]--
    $xmlWriter.WriteStartElement("UserData")                                                            #--[ Write the UserData Element ]--
    $xmlWriter.WriteElementString("DisplayName",$ADUser.DisplayName)
    $xmlWriter.WriteElementString("Designation",$ADUser.extensionAttribute1)
    $xmlWriter.WriteElementString("Title",$ADUser.title)
    $xmlWriter.WriteElementString("Company",$ADUser.company)
    $xmlWriter.WriteElementString("Department",$ADUser.extensionAttribute4)
    $xmlWriter.WriteElementString("StreetAddress",$ADUser.extensionAttribute5)
    $xmlWriter.WriteElementString("City",$ADUser.l)
    $xmlWriter.WriteElementString("Statecode",$ADUser.st)
    $xmlWriter.WriteElementString("PostalCode",$ADUser.PostalCode)
    $xmlWriter.WriteElementString("TelephoneNumber",$ADUser.TelephoneNumber)
    $xmlWriter.WriteElementString("MobileNumber",$ADUser.mobile)  
    $xmlWriter.WriteElementString("FaxNumber",$ADUser.facsimileTelephoneNumber)
    $xmlWriter.WriteElementString("MLS",$ADUser.extensionAttribute2)
    $xmlWriter.WriteEndElement() | out-null  
    $xmlWriter.WriteEndElement() | out-null                                                             #--[ Close RootElement ]--  
    $xmlWriter.WriteEndDocument()                                                                       #--[ End the XML Document ]--
    $xmlWriter.Finalize                                                                                 #--[ Finish The Document ]--
    $xmlWriter.Flush | Out-Null
    $xmlWriter.Close()
}Else{
    If ($Console){Write-Host "`n--- Nothing to do ---" -ForegroundColor Green}
}

If ($Remove){   #--[ Move all files to backup. Clear registry keys ]--
      ForEach ($key in $keys){
            $value = (Get-ItemProperty $regkeypath).$key -eq $null
            If ($value -eq $False) {Remove-ItemProperty -path $regkeypath -name $key}
      }
      copy-item $LocalSignaturePath "$AppData\Microsoft\Signatures Backup" -recurse -force
      remove-item -path $localSignaturePath\* -recurse -force
}

If ($Console){Write-host "`n--- Completed ---" -ForegroundColor red }


<#
.SYNOPSIS
	This script is a template that allows you to extend the toolkit with your own custom functions.
.DESCRIPTION
	The script is automatically dot-sourced by the AppDeployToolkitMain.ps1 script.
.NOTES
    Toolkit Exit Code Ranges:
    60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
    69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
    70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK 
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
)

##*===============================================
##* VARIABLE DECLARATION
##*===============================================

# Variables: Script
[string]$appDeployToolkitExtName = 'PSAppDeployToolkitExt'
[string]$appDeployExtScriptFriendlyName = 'App Deploy Toolkit Extensions'
[version]$appDeployExtScriptVersion = [version]'1.5.0'
[string]$appDeployExtScriptDate = '06/11/2015'
[hashtable]$appDeployExtScriptParameters = $PSBoundParameters

##*===============================================
##* FUNCTION LISTINGS
##*===============================================


#region Function Write-SWIDTag
Function Write-SWIDTag {
<#
.SYNOPSIS
	Writes an SWIDTag xmlfile as specified in ISO/IEC 19770 around Software assetmanagement .DESCRIPTION
	Writes an SWIDTag xmlfile as specified in ISO/IEC 19770 around Software assetmanagement
    Which can help alot in software asset management and license compliance .PARAMETER SWIDTagFilename 
    eg: regid.1995-08.com.techsmith Snagit 12.swidtag saved in $env:ProgramData 
.PARAMETER EntitlementRequired
	Does the software require a license to be used. 
    Optional - Default is $true 
.PARAMETER ProductTitle
    Title of the software installed. 
    Optional - Default is $appname when used in powershell app deployment toolkit 
.PARAMETER ProductVersion
    Version of the software in x.y.z.w format, where x is major, y is minor, z is build and w is review. 
    Optional - Default is $appversion when used in powershell app deployment toolkit
    If your product version has fewer levels, complete with 0 for any level (minor, build or review) you don't have.
.PARAMETER CreatorName
	Name of the software creator. 
    Optional - Default is $appvendor when used in powershell app deployme
	Path of the SWIDTag file to create. 
    Optional - Default $env:ProgramData + '\' + $CreatorRegid + '\' + $creatorRegid + ' ' + $ProductTitle + ' ' + $ProductVersion +'.swidtag'
.PARAMETER CreatorRegid
	Regid of the software creator. Regid's in ISO/IEC 19770 are defined as the word regid. followed by the date you owned a particular DNS domain in YYYY-MM format
    followed by the domain name in reverse formatting. eg: regid.1991-06.com.microsoft 
    Mandatory field
.PARAMETER LicensorName
	Name of the software licensor. Optional - Default is $appvendor when used in powershell app deployment toolkit 
.PARAMETER LicensorRegid
	Regid of the software Licensor. Regid's in ISO/IEC 19770 are defined as the word regid. followed by the date you owned a particular DNS domain in YYYY-MM format
    followed by the domain name in reverse formatting. eg: regid.1991-06.com.microsoft
    Optional - Default is $CreatorRegid value 
.PARAMETER SoftwareUniqueID
    UniqueID assigned to the software.
    Can be set in GUID format 8HexCharacters-4HexCharacters-4HexCharacters-4HexCharacters-12HexCharacters, the format used by MSI productid's
    eg: BDFD9ADC-3F97-4A8A-A533-987B21776449
    Optional - Default is $producttitle + ' ' +  $productversion
.PARAMETER TagCreatorRegid
    Regid of the TagCreator. Regid's in ISO/IEC 19770 are defined as the word regid. followed by the date you owned a particular DNS domain in YYYY-MM format
    followed by the domain name in reverse formatting. eg: regid.1991-06.com.microsoft
    Optional - Default is $CreatorRegid value



.EXAMPLE
	Write-SWIDTag -CreatorRegid 'regid.2008-03.be.oscc' 
    Write a software id tag using the minimum set of parameters to specify, other parameters will use the powershell app deployment toolkit variables for software title, version, vendor to generate the tag.
    This example will not work outside the Powershell app deployment toolkit .EXAMPLE
    Write-SWIDTag -ProductTitle "SWIDTagHandler" -ProductVersion "1.2.3.4" -CreatorName "OSCC" -CreatorRegid "regid.2008-03.be.oscc" -SoftwareUniqueid "a6ca313e-c7ae-447c-9ee0-bd872278c166" 
    Write a software id tag using the minimum set of parameters when used outside of the powershell app deployment toolkit
    This generates a tag for a product called Swidtaghandler with version 1.2.3.4 by Software vendor OSCC, oscc's regid is regid.2008-30.BE.OSCC .NOTES
    You can generate a guid in powershell using this code [guid]::NewGuid(), don't generate a new one for every install. The idea of the guid is that is the same on every machine you write this swidtag to
    If this is to install an msi, the ps app deployment toolkit containts a function to get the productcode from that msi, please use that productcode as your guid, the following command returns the productcode from an msi.
    Get-MsiTableProperty -Path RelativePathToMsiFile | select productcode .NOTES
    the regid as definied in ISO19770 is the literal string regid, followed by the date the domain was first registered in YYYY-MM format, followed by the domain reversed
    eg: regid.1991-06.com.microsoft
    You can typically find the registration date using whois

.LINK
	http://psappdeploytoolkit.codeplex.com
.LINK
    http://www.scug.be/thewmiguy
.LINK
    http://www.oscc.be
.LINK
    https://technet.microsoft.com/en-us/library/gg681998.aspx#BKMK_WhatsNewSP1
.LINK
    http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/
#>
	[CmdletBinding()]
	Param (

		[Parameter(Mandatory=$false,
        HelpMessage='Does the software license require an entitlement also known as software usage right or license?')]
		[boolean]$EntitlementRequired = $true,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Name of the software to add to the swidtag, default is the $appName variable from the PS app deploy toolkit.')]
        [Alias('SoftwareName','ApplicationName','Name')]
		[ValidateNotNullorEmpty()]
        [string] $ProductTitle = $appName_withspace, 

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='32 or 64 bit architecture, default is the $appArch variable from the PS app deploy toolkit.')]
        [Alias('AppArchitecture','ApplicationArchitecture','BitNess')]
		[ValidateSet('x86','X86','x64','X64')]
        [string] $ProductArch = $appArch,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Version of the software to add to the swidtag, in 4 digit format separated by the . character')]
        [Alias('SoftwareVersion','ApplicationVersion','Version')]
		#[ValidatePattern("^\d+\.\d+\.\d+\.\d+$")]
        [string] $ProductVersion = $appVersion,  

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Path, including filename of the swidtag to generate.')]
        [Alias('Filename')]
        [string] $SWIDTagFilename = $env:ProgramData + '\' + $CreatorRegid + '\' + $creatorRegid + ' ' + $ProductTitle + ' ' + $ProductVersion + ' ' + $productArch +'.swidtag', 

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Vendor of the software to add to the swidtag, default is the $appvendor variable in the PS app deploy toolkit.')]
        [Alias('SoftwareVendor','Vendor','Manufacturer')]
		[ValidateNotNullorEmpty()]
        [string] $CreatorName = $appVendor_withspace, 

        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Regid of the Vendor of the software to add to the swidtag, format is: regid.YYYY-MM.FirstLevelDomain.SecondLeveldomain, eg: regid.1991-06.com.microsoft the date is the date when the domain referenced was first registered')]
        [Alias('Regid','VendorRegid')]
        #[ValidatePattern("^regid\.(19|20)\d\d-(0[1-9]|1[012])\.[A-Za-z]{2,6}\.[A-Za-z0-9-]{1,63}.+$")]
        [string]$CreatorRegid,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='LicensorName of the software to add to the swidtag.')]
		[ValidateNotNullorEmpty()]
        [string] $LicensorName = $CreatorName, 

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Licensor Regid of the software to add to the swidtag, format is identical to $CreatorRegid')]
        #[ValidatePattern("^regid\.(19|20)\d\d-(0[1-9]|1[012])\.[A-Za-z]{2,6}\.[A-Za-z0-9-]{1,63}.+$")]
        [string]$LicensorRegid = $creatorRegid,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Uinqueid to add to the swid tag, format is: 8 hex chars-4 hex chars-4 hex chars-4 hex chars-12 hex chars. If this is an MSI install use the MSI ProductID, guid can but does not have to be enclosed in {}')]
        [Alias('UniqueId','GUID')]
        #[ValidatePattern("^{?[A-Z0-9]{8}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{12}}?$")]
		[string]$SoftwareUniqueid=$ProductTitle  + ' ' + $ProductVersion,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Tag creator name of the software to add to the swidtag.')]
		[ValidateNotNullorEmpty()]
        [string] $TagCreatorName = $CreatorName, 

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='TagCreator Regid of the software to add to the swidtag, format is identical to $CreatorRegid')]
        #[ValidatePattern("^regid\.(19|20)\d\d-(0[1-9]|1[012])\.[A-Za-z]{2,6}\.[A-Za-z0-9-]{1,63}.+$")]
        [string]$TagCreatorRegid = $creatorRegid,

		[Parameter(Mandatory=$false)]
		[ValidateNotNullOrEmpty()]
		[boolean]$ContinueOnError = $true
	)
	
	Begin {
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
	}
	Process {
		Try {
                # this is where the document will be saved:
                $Path = $SWIDTagFilename.Replace(' ','_')
                $SWIDTagFolder = split-path $Path

				Write-Log -Message "Testing whether [$SWIDTagFolder] exists, if not path is recursively created." -Source ${CmdletName}
                if(-not(test-path($SWIDTagFolder)))
                {
				Write-Log -Message "Folder [$SWIDTagFolder] does not exist, recursively creating it." -Source ${CmdletName}
                    New-Item -path $SWIDTagFolder -type directory
				Write-Log -Message "Folder [$SWIDTagFolder] is successfully created." -Source ${CmdletName}
                }
                else
                {
				Write-Log -Message "Folder [$SWIDTagFolder] already exists, no need to create it." -Source ${CmdletName}
                }


 
                # get an XMLTextWriter to create the XML and set the encoding to UTF8
				Write-Log -Message "Creating XMLTextWriter object and set encoding to UTF8." -Source ${CmdletName}
                $encoding = [System.Text.Encoding]::UTF8
                Write-Log -Message "Creating XMLTextWriter file in path $Path." -Source ${CmdletName}
                Write-Log -Message "Creating XMLTextWriter file in path $SWIDTagFileName." -Source ${CmdletName}
                $XmlWriter = New-Object System.XMl.XmlTextWriter($Path,$encoding)
 
                # choose a pretty formatting: (set Indentation to 1 Tab)
				Write-Log -Message "Setting indentation of the file to 1 Tab." -Source ${CmdletName}
                $xmlWriter.Formatting = 'Indented'
                $xmlWriter.Indentation = 1
                $XmlWriter.IndentChar = "`t"

 
                # write the header
				Write-Log -Message "Writing SWIDTag header." -Source ${CmdletName}
                $xmlWriter.WriteStartDocument("true")
                $xmlWriter.WriteProcessingInstruction("xml-stylesheet", "type='text/xsl' href='style.xsl'")
                $XmlWriter.WriteStartElement("swid","software_identification_tag","http://standards.iso.org/iso/19770/-2/2008/schema.xsd")
                $XmlWriter.WriteAttributeString("xsi:schemalocation","http://standards.iso.org/iso/19770/-2/2008/schema.xsd software_identification_tag.xsd")
                $XmlWriter.WriteAttributeString("xmlns:xsi","http://www.w3.org/2001/XMLSchema-instance")
                $XmlWriter.WriteAttributeString("xmlns:ds","http://www.w3.org/2000/09/xmldsig#")
                # write the mandatory elements, and the sccm supported elements of an iso 19770 swidtag
				Write-Log -Message "Writing EntitlementRequired field, settng it to [$EntitleMentRequired]." -Source ${CmdletName}
                if ($EntitlementRequired)
                {
                    $XmlWriter.WriteElementString("swid:entitlement_required_indicator","true")
                }
                else
                {
                    $XmlWriter.WriteElementString("swid:entitlement_required_indicator","false")
                }
				Write-Log -Message "Writing other SWIDTag elements." -Source ${CmdletName}
                Write-Log -Message "Writing ProductTitle, setting it to [$ProductTitle]" -Source ${CmdletName}
                $XmlWriter.WriteElementString("swid:product_title",$ProductTitle)
                $XmlWriter.WriteStartElement("swid:product_version")
                $XmlWriter.WriteElementString("swid:name",$ProductVersion)
                $XmlWriter.WriteStartElement("swid:numeric")
                if (!($productversion.Contains('.')))
                {
                    Write-Log -Message "ProductVersion [$ProductVersion] is not in the correct format. Altering format to comply with SwidTag format, appending .0. New Productversion is [$ProductVersion.0]"
                    $productversion = $productversion + '.0'
                }

                Write-Log -Message "Writing ProductVersion, setting it to [$ProductVersion]" -Source ${CmdletName}
                $splitProductVersion = $ProductVersion.Split('.')
                $XmlWriter.WriteElementString("swid:major",$splitProductversion[0])
                $XmlWriter.WriteElementString("swid:minor",$splitProductversion[1])
                $XmlWriter.WriteElementString("swid:build",$splitProductversion[2])
                $XmlWriter.WriteElementString("swid:review",$splitProductversion[3])
                $XmlWriter.WriteEndElement();
                $XmlWriter.WriteEndElement();
                $XmlWriter.WriteStartElement("swid:software_creator")
                $XmlWriter.WriteElementString("swid:name",$CreatorName)
                $XmlWriter.WriteElementString("swid:regid",$creatorRegid)
                $XmlWriter.WriteEndElement();
                $XmlWriter.WriteStartElement("swid:software_licensor")
                $XmlWriter.WriteElementString("swid:name",$LicensorName)
                $XmlWriter.WriteElementString("swid:regid",$LicensorRegid)
                $XmlWriter.WriteEndElement();
                $XmlWriter.WriteStartElement("swid:software_id")
                if ($SoftwareUniqueid.ToString().StartsWith('{'))
                {
                  $swid = $SoftwareUniqueid.ToString().Substring(1,36)
                  Write-Log -Message "Writing SoftwareuniqueId [$swid] to swid tag." -Source ${CmdletName}
                  $XmlWriter.WriteElementString("swid:unique_id",$SoftwareUniqueid.ToString().Substring(1,36))
                }
                else
                {
                  Write-Log -Message "Writing SoftwareuniqueId [$SoftwareUniqueid] to swid tag." -Source ${CmdletName}
                    $XmlWriter.WriteElementString("swid:unique_id",$SoftwareUniqueid)
                }
                $XmlWriter.WriteElementString("swid:tag_creator_regid",$TagCreatorRegid)
                $XmlWriter.WriteEndElement();
                $XmlWriter.WriteStartElement("swid:tag_creator")
                $XmlWriter.WriteElementString("swid:name",$TagCreatorName)
                $XmlWriter.WriteElementString("swid:regid",$TagCreatorRegid)
                $XmlWriter.WriteEndElement();
                $xmlWriter.Flush()
                $xmlWriter.Close()
		}
		Catch {
			Write-Log -Message "Failed to write swidtag to destination [$SWIDTagFilename]. `n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
			If (-not $ContinueOnError) {
				Throw "Failed to write swidtag to destination [$SWIDTagFilename]: $($_.Exception.Message)"
			}
		}
	}
	End {
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
	}
}
#endregion



##*===============================================
##* END FUNCTION LISTINGS
##*===============================================

##*===============================================
##* SCRIPT BODY
##*===============================================

If ($scriptParentPath) {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] dot-source invoked by [$(((Get-Variable -Name MyInvocation).Value).ScriptName)]" -Source $appDeployToolkitExtName
}
Else {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] invoked directly" -Source $appDeployToolkitExtName
}

##*===============================================
##* END SCRIPT BODY
##*===============================================
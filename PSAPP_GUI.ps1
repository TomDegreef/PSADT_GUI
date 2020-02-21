<#
.SYNOPSIS
	This GUI allows for (basic) easy creation of PsAppDeploymentToolkit packages.
.DESCRIPTION
        - Fully unattended generation of the Deploy-Application.PS1 file
        - Merging of your source binaries with the PS App toolkit on a destination Share (must be UNC)
        - Creating the SCCM Application + Deployment Type
        - Adding a Software ID tag and using this as a detection mechanism.
.NOTES
    FileName:    PSAPP_GUI.ps1
    Blog: 	 http://www.OSCC.Be
    Author:      Tom Degreef
    Twitter:     @TomDegreef
    Email:       Tom.Degreef@OSCC.Be
    Created:     2017-08-23
    Updated:     2018-02-06
    
    Version history
    1.4   - Default DT is now install for system, fixed uninstall but where swidtag wasn't removed properly if there was a space in the appname, Creates intune W32 packages, Browse for MSI,EXE files, Whois lookup fixed
    1.3   - (2018-02-06) Created Preferences TAB with automatic actions for SCCM, Specify your own limiting collection, create security groups in AD and link them to collections
    1.2   - (2017-09-27) SWID-Tags are Cached and will be pre-filled out if vendor is used again, Source & Destination path are cached, removed bug in detection method, added listbox for uninstall functionality
    1.1   - (2017-09-06) Experimental support for SWID-Tag lookup is added
    1.0   - (2017-09-05) Added 2nd Tab for SCCM, Allowing software distribution & Collection Creation
    0.9	  - (2017-08-23) Initial version of the GUI
.LINK 
	Http://www.OSCC.Be
#>

$gui_version = '1.4'
$logfile = "$env:systemdrive\temp\PSAPPgui\PSAPP_Gui.txt"
$logmodule = "$env:systemdrive\temp\PSAPPgui\PSAPP_Gui.log"
$global:packagepath = ""
$global:architecture = ""
$global:fullname=""
$global:Sitecode=""
$global:ADEnabled=$False
$global:dest=""
$global:detectionscript=""
$global:Startlocation=get-location
$global:intunedection=""

$logopath = "$psscriptroot\config\OSCCLogo.jpg"

#Check if OSCC Logging module is installed, if not, install it
if ($PSVersionTable.PSVersion -ge '5.0.0.0')
{
  if (!(Get-PackageProvider | Where-Object {$_.Name -eq 'NuGet'}))
    {  Install-PackageProvider -Name Nuget -force -MinimumVersion 2.8.5.208 }
  if (!(get-module -ListAvailable -Name osccpslogging | Where-Object {$_.Version -ge '1.5.0.1'}))
    { 
    write-host "installing module"
    install-module -Name OSCCPsLogging -MinimumVersion 1.5.0.1 -Force}
}

if (!(test-path "$env:systemdrive\temp\PSAPPgui"))
{
       New-Item -itemtype Directory -path $env:systemdrive\temp\PSAPPgui -Force
}

try {$log = Initialize-CMLogging -LogFileName $logmodule -OMSFileLogLevel off -ConsoleLogLevel off -fileLogLevel info} catch {}
function log-item
{
param(  $logline,
        $severity = "Info" )
      [string]$currentdatetime = get-date
        if (get-module -name OSCCPSLogging)
      {
                Switch ($severity)
                {
                        "Info" {$log.Info($logline)}
                        "Warn" {$log.Warn($logline)}
                        "Error" {$log.Error($logline)}
                }
                
      }
        else
        {
                 Switch ($severity)
                {
                        "Info" {[string]$output = $currentdatetime + " - INFO - " + $logline}
                        "Warn" {[string]$output = $currentdatetime + " - WARNING - " + $logline}
                        "Error" {[string]$output = $currentdatetime + " - ERROR - " + $logline}
                }              
        $output | out-file $Logfile -append
        }
}

Log-Item -logline "Starting Powershell AppDeployment Toolkit GUI version $gui_version"

try 
        {
        $inputXML = Get-Content "$psscriptroot\config\MainWindow.xaml" -ErrorAction stop
        $inputXML = $inputXML -replace "ReplacePath", $logopath
        log-item("GUI XML sucessfully loaded")
        }
catch 
        {
        log-item -logline "error loading GUI XML" -severity "Error"
        log-item -logline "exiting script" -severity "Error"
        exit        
        }

try {
        $inputApptoolkit = get-content "$psscriptroot\config\Deploy-Application.ps1" -ErrorAction stop
        log-item("'Powershell AppDeployment Toolkit' input-file succesfully loaded")
}
catch {
        log-item -logline "error loading 'Powershell AppDeployment Toolkit' input-file " -severity "Error"
        log-item -logline "exiting script" -severity "Error"
        exit        
}

try {
        $Preferences = New-Object -TypeName XML
        $Preferences.load("$psscriptroot\config\Prefs.xml")
        log-item("Loading preferences")
}
catch {
        log-item -logline "Failed loading preferences " -severity "Error"
        log-item -logline "Continuing script but without preferences" -severity "Warn"     
}

try {
        $SWIDtags = New-Object -TypeName XML
        $SWIDtags.load("$psscriptroot\config\swid.xml")
        log-item("Loading SWIDtag cache")
}
catch {
        log-item -logline "Failed loading SWIDtag cache" -severity "Error"
        log-item -logline "Continuing script but without SWIDtag cache" -severity "Warn"     
         Log-Item -logline "Caught an exception:" -severity "Error"
            log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
            log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
}

Function Prefill{

        $wpfCB_SourcePath.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.settings.Save_SP)
        $wpfCB_DestinationPath.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.Settings.Save_DP)
        $wpfCB_SiteCode.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.Settings.Save_SiteCode)
        $wpfCB_DP.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.Settings.Save_DistPoint)
        $wpfCB_DPG.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.Settings.Save_DistGroup)

        #If no sitecode exist in preference, disable autoconnect untill manual connection is made first and the sitecode is saved in preferences
        if ($Preferences.Preferences.SiteCode.SC -ne "")
                {
                log-item -logline "Sitecode is available in preferences, enabling automatic SCCM functionality"        
                $wpfCB_Auto_Sccm.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.SCCM_Prefs.Autoconnect)
                $wpfCB_Auto_Import.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.SCCM_Prefs.AutoImport)
                $wpfCB_Auto_Collection.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.SCCM_Prefs.AutoCollection)
                Try { Import-Module ActiveDirectory -ErrorAction Stop
                        $wpfCB_Auto_ADGroup.IsEnabled = $true
                        $wpfTB_LDAP.IsEnabled = $true
                        #$wpfCB_CreateAD.IsEnabled = $true
                        log-item -logline "Active Directory PS Module found, enabling AD Functionality"
                        $wpfCB_Auto_ADGroup.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.SCCM_Prefs.AutoADGroup)
                        #if ($wpfCB_Auto_ADGroup.IsChecked = $True)
                       # {
                        #        $wpfCB_CreateAD.IsChecked = $True
                       # }
                }
                catch {
                
                        log-item -logline "Active Directory PS Module NOT found, Disabling AD Functionality" -severity warn
                        $wpfCB_Auto_ADGroup.IsEnabled = $False
                        $wpfTB_LDAP.IsEnabled = $false
                       # $wpfCB_CreateAD.IsEnabled = $False
                        }
                } 
        else {
                log-item -logline "Sitecode is NOT available in preferences, disabling automatic SCCM functionality" -severity warn
                $wpfTB_AutoPerform.Content = "Automatic SCCM actions below are not available until a package is manually imported through the GUI interface"
                $wpfTB_AutoPerform.Foreground = "red"
                $wpfCB_Auto_Sccm.IsEnabled = $false
                $wpfCB_Auto_Import.IsEnabled = $false
                $wpfCB_Auto_Collection.IsEnabled = $false
                #$wpfCB_Auto_Sccm.tooltipservice = $true
                #$wpfCB_Auto_Sccm.tooltip = "Setting is currently disabled because SCCM details are unknown. After creating a package, connect to SCCM manually from the app"
                #$wpfCB_Auto_Import.tooltip = "Setting is currently disabled because SCCM details are unknown. After creating a package, connect to SCCM manually from the app"
                $wpfCB_Auto_ADGroup.IsEnabled = $False
                $wpfTB_LDAP.IsEnabled = $false
                #$wpfCB_CreateAD.IsEnabled = $False
        }

        if (($Preferences.Preferences.SelectedDP.DP.Count -gt 1) -or ($Preferences.Preferences.SelectedDPGroup.Dpg.count -gt 1))
        {
                log-item -logline "At least 1 DP or DP Group was used before, enabling Auto-distribution"
                $wpfCB_Auto_Distribute.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.SCCM_Prefs.AutoDistribute)  
        }
        else {
                log-item -logline "No DP or DP Group is known, disabling Auto-distribution"
                $wpfCB_Auto_Distribute.isenabled = $false
        }
        
        $wpfCB_Intunepackage.IsChecked = [System.Convert]::ToBoolean($Preferences.Preferences.Settings.IntunePackage)
        $wpfTB_DefaultLimiting.Text = $Preferences.Preferences.SCCM_Prefs.LimitingCollection
        $wpfTB_LDAP.Text = $Preferences.Preferences.SCCM_Prefs.LDAP 
        $wpfprefix.text = $Preferences.Preferences.Collection.Prefix
        $wpfsuffix.text = $Preferences.Preferences.Collection.Suffix
        $WPFTB_IntuneFolder.Text = $Preferences.Preferences.SCCM_Prefs.IntuneAppfolder

        if ($wpfCB_Auto_ADGroup.IsChecked -and $wpfCB_Auto_ADGroup.IsEnabled)
                {
                      #  $wpfCB_CreateAD.Ischecked = $True
                }

        if ($wpfCB_SourcePath.IsChecked -and $Preferences.Preferences.Sourcepath.SP -ne "")
                {
                        $wpfinputSource.text = $Preferences.Preferences.Sourcepath.sp
                        $wpfinputSource.Tag = 'cleared'
                }
        if ($wpfCB_DestinationPath.IsChecked -and $Preferences.Preferences.Destinationpath.DP -ne "")
                {
                        $wpfInputDestination.text = $Preferences.Preferences.Destinationpath.DP
                        $wpfInputDestination.Tag = 'cleared'
                }  
        if ($Preferences.Preferences.EndMessage.EM -ne "")
                {
                        $wpfTB_EndMessage.Text = $Preferences.Preferences.EndMessage.EM
                }
        if (-not ($wpfCB_Intunepackage.IsChecked))
                {
                       $WPFTB_IntuneFolder.IsEnabled = $False
                }
        log-item -logline "Updated GUI with preset preferences"
}

Function Update_Prefs{
        param($sp,$dp,$em)

        If ($wpfCB_SourcePath.IsChecked)
        { 
                $item = Select-XML -Xml $Preferences -XPath '//Sourcepath[1]'
                $item.Node.SP = $sp
                $Preferences.Save("$psscriptroot\config\Prefs.xml")
        }
        if ($wpfCB_DestinationPath.IsChecked)
        {
                $item = Select-XML -Xml $Preferences -XPath '//Destinationpath[1]'
                $item.Node.DP = $dp
                $Preferences.Save("$psscriptroot\config\Prefs.xml")
        }
        if ($em -ne "")
        {
                $item = Select-XML -Xml $Preferences -XPath '//EndMessage[1]'
                $item.Node.EM = $em
                $Preferences.Save("$psscriptroot\config\Prefs.xml")
        }
        
}
function Check_and_Save_xml
{       Param($name,$year,$month,$domain)

        $output = $SWIDtags.swid.RegId.Vendor | Where-object {$_.Name -eq $Name} | Select-Object -Property Name,Year,month,domain
        if ($output.name -ne $wpfVendor.text)
        { log-item -logline ("Adding RegID information to preferences file for vendor : " + $name)
                $item = Select-XML -Xml $SWIDtags -XPath '//Vendor[1]'
                $newnode = $item.Node.CloneNode($true)
                $newnode.Name = $Name
                $newnode.Year = $Year
                $newnode.Month = $Month
                $newnode.Domain = $Domain
                $Vendor = Select-XML -Xml $SWIDtags -XPath '//RegId'
                $Vendor.Node.AppendChild($newnode)
                $SWIDtags.Save("$psscriptroot\config\swid.xml")
        }
        else {
                log-item -logline ("fRegID information was already available in swid.xml")
        }
}
function lookup-url
{
        param($url)
        
        [Regex]$regidYY = "[1-2]\d{3}"
        [Regex]$regidMM = "[0-1][0-9]"
        
        $verify = $url.split(".")

        $error.clear()

        $headers=@{}
        $headers.Add("Accept", "application/json")
        $headers.Add("authorization", "Token token=9f27053628bc1eed8d0b382ece8fed7a")
        $response = Invoke-WebRequest -Uri "https://jsonwhois.com/api/v1/whois?domain=$url" -Method GET -Headers $headers -UseBasicParsing

        $Response = ConvertFrom-Json -InputObject $response
        log-item  ("Returned value from whois database : $($Response.created_on)")

            if ($error.count -eq 0)
            { 
            $myyear = $Response.created_on.Substring(0,4)
            $mymonth = $Response.created_on.Substring(5,2)
        
                if ($myyear -match $regidyy -and $mymonth -match $regidmm)
                    {
                        $wpfregidYY.Text =  $myyear
                        $wpfregidYY.Tag = 'cleared' 
                        $WPFRegIDYY.background = "green"
                        $wpfregidMM.Text = $mymonth
                        $wpfregidMM.Tag = 'cleared'
                        $wpfregidMM.background = "green"
                        $wpfregidDomain.Text = $verify[1] + "." + $verify[0]
                        $wpfregidDomain.Tag = 'cleared'
                        $wpfregidDomain.background = "green"
                        log-item -logline ("Lookup of " + $wpfVendorLookup.text + " succeeded. Pre-filling data in form") 
                    }
                else
                    {
                        [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                        [Microsoft.VisualBasic.Interaction]::MsgBox("Lookup failed. Please check manually at www.whois.net",'OKOnly,Information',"URL Lookup Failed")
                        log-item -logline ("Lookup failed for " + $wpfVendorLookup.text ) -severity Error
                    }
            }
            else
            {
                [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                [Microsoft.VisualBasic.Interaction]::MsgBox("Lookup failed. Please check manually at www.whois.net",'critical',"URL Lookup Failed")
                log-item -logline ("Lookup failed for " + $wpfVendorLookup.text ) -severity Error
            }

        
}              

Function Fill-swid
{
        log-item -logline ("Checking if vendor information is availablein swid.xml")
        if ($wpfvendor.text -ne '')
        {
                $output = $SWIDtags.swid.RegId.Vendor | Where-object {$_.Name -eq $wpfVendor.text} | Select-Object -Property Name,Year,month,domain
                if ($output.name -ne $wpfVendor.text )
                        { log-item -logline ($wpfVendor.text + " not found as a vendor in preferences file")}
                else { 
                        log-item -logline ($wpfVendor.text + " was found as a vendor in preferences file. Populating fields.")
                        $wpfregidYY.text = $output.year
                        $wpfregidMM.text = $output.month
                        $wpfregidDomain.text = $output.domain
                        $wpfregidyy.background = "green"
                        $wpfregidmm.background = "green"
                        $wpfregidDomain.background = "green"
                        $wpfregidYY.Tag = 'cleared'
                        $wpfregidMM.Tag = 'cleared'
                        $wpfregidDomain.Tag = 'cleared'

                }
        }    
}

Function Compose-Package
{
        Param($source,$destination,$appvendor,$appname,$appversion,$appbitness)

        Switch ($appbitness)
                {
                "0"
                        {
                        $global:architecture = "X86"
                        }
                "1"
                        {
                        $global:architecture = "X64"
                        }
                }

        if ($destination.Endswith("\"))
        {
                $global:packagepath = $destination + $appvendor + "\" + $appname + "_" + $appversion + "_" + $global:architecture
        }
        else
        {
                $global:packagepath = $destination + "\" + $appvendor + "\" + $appname + "_" + $appversion + "_" + $global:architecture
        }

        $global:dest=$destination

        $pct = 0

        new-item -itemtype Directory -path $packagepath -force

        #$WPFPackage_progress.Content = "Copying toolkit to $packagepath"
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFPackage_progress.Content = "Copying toolkit to $packagepath"} ),$null, $null )

        Copy-item -path "$psscriptroot\Toolkit\*" -Destination $packagepath -Recurse -Force
        log-item ("Copying toolkit to $packagepath")

        if (-not (test-path "$packagepath\Files")) {
                Log-Item "Subfolder Files did not exist in toolkit, creating it"
                New-Item -ItemType Container -Path "$packagepath\Files" -force
        }

        if (-not (test-path "$packagepath\SupportFiles")) {
                Log-Item "Subfolder SupportFiles did not exist in toolkit, creating it"
                New-Item -ItemType Container -Path "$packagepath\SupportFiles" -force
        }

        $sourcefiles = Get-ChildItem $source -Recurse -Force
        
        [float]$blocks = (100 / $sourcefiles.Count)
        foreach ($file in $sourcefiles)
                {
                        $size = "{0:N2}" -f (($file.Length)/1MB)
                        $dest = ($file.FullName -replace [regex]::Escape($source), "$packagepath\Files")

                        #$WPFPackage_progress.Content = "Copying file $file with size $size MB"
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFPackage_progress.Content = "Copying file $file with size $size MB"} ),$null, $null )

                       if ($file.PSIsContainer)
                        {
                                $null = New-Item -ItemType Container -Path $dest -force
                                log-item ("Created folder $dest")
                        }
                        if (-not $file.psiscontainer)
                        {
                                Copy-Item $file.fullname -Force -Destination $dest
                                log-item ("Copying $file ($size MB)")
                        }
                        
                        $pct += $blocks
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $wpfprogress.Value = $pct } ),$null, $null )

                }

}

Function Invoke-MemberOnType #Used in the extraction of MSI properties
{
    param
    (
        #The string containing the name of the method to invoke
        [Parameter(Mandatory=$true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$Name, 
        
        #A bitmask comprised of one or more BindingFlags that specify how the search is conducted. 
        #The access can be one of the BindingFlags such as Public, NonPublic, Private, InvokeMethod, GetField, and so on
        [Parameter(Mandatory=$false, Position = 3)]
        [System.Reflection.BindingFlags]$InvokeAttr = "InvokeMethod", 

        #The object on which to invoke the specified member
        [Parameter(Mandatory=$true, Position = 1)]
        [ValidateNotNull()]
        [Object]$Target, 

        #An array containing the arguments to pass to the member to invoke.
        [Parameter(Mandatory=$false, Position = 2)]
        [Object[]]$Arguments = $null
    )

    #Invokes the specified member, using the specified binding constraints and matching the specified argument list.

    $Target.GetType().InvokeMember($Name,$InvokeAttr,$null,$Target,$Arguments)
} 

Function File-Picker
{
        $openFileDialog = New-Object system.windows.forms.openfiledialog
        
        if ($wpfCB_SourcePath.IsChecked -and $Preferences.Preferences.Sourcepath.SP -ne "")
        {
                $openFileDialog.initialDirectory = $wpfinputSource.text 
        }
        else {
                $openFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()   
        }  
        $openFileDialog.title = "Select Executable or MSI"   
        $openFileDialog.filter = "Installable Files (*.exe;*.msi)|*.exe;*.msi|All files (*.*)|*.*" 
        $result = $openFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))

        if($result -eq "OK")   
         {    
                $fullfilename = $OpenFileDialog.filename   
                $filepath =  split-path -path $fullfilename
                $Filename = split-path -path $fullfilename -Leaf
                $Extension = [System.IO.Path]::GetExtension($Filename)
                Log-Item -logline "Filebrowser results :"
                Log-Item -Logline "selected path = $filepath"
                Log-Item -Logline "selected filename = $filename"
                Log-Item -Logline "Filetype = $extension"

                Log-Item -Logline "Clearing all entered fields on Package Tab"
                $wpfappname.text = ""
                $wpfappname.background = "white"
                $wpfvendor.text = ""
                $wpfvendor.Background = "white"
                $WPFAppversion.Text = ""
                $WPFAppversion.Background = "white"
                $WPFAppLanguage.Text = ""
                $WPFAppLanguage.Background = "white"
                $WPFAppRevision.Text = ""
                $WPFAppRevision.Background = "white"
                $wpfVendorLookup.Text = ""
                $wpfVendorLookup.Background = "white"
                $wpfregidYY.Text = "YYYY"
                $wpfregidYY.Background = "white"
                $wpfregidMM.Text = "MM"
                $wpfregidMM.Background = "white"
                $wpfregidDomain.Text = "Com.Domainname"
                $wpfregidDomain.Background = "white"
                $WPFLicense.SelectedIndex = "-1"
                $WPFLicenseText.Background = "white"
                $wpfinputsource.Background = "white"
                $wpfinputdestination.Background = "white"
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFPackage_progress.Content = ""} ),$null, $null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $wpfprogress.value = 0 } ),$null, $null )

                Log-Item -Logline "Clearing all entered fields on SCCM Tab"
                $WPFsitecode.content = ""
                $WPFsitecode.Background = "white"
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfImport_progress.Content = ""} ),$null, $null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $WPFProgress_SCCM.value = 0 } ),$null, $null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfdistribute_progress.Content = ""} ),$null, $null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $WPFPB_Progress_distribute.value = 0 } ),$null, $null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfsccm.background = "white" } ),$null, $null )
                $wpfsccm.IsEnabled = $false
                $wpfdp.items.Clear()
                $wpfdpgroup.items.Clear()
                $WPFCollections_info.Background = "white"
                $WPFCollections_info.Content = ""

            } 
            else { Log-Item -Logline "File-browser Cancelled!"} 

        switch ($extension) {
                ".msi" {
                        Log-Item "MSI file chosen"
                        $type = [Type]::GetTypeFromProgID("WindowsInstaller.Installer") 
                        $installer = [Activator]::CreateInstance($type)
                        $db = Invoke-MemberOnType "OpenDatabase" $installer @($fullfilename,0)
                        $view = Invoke-MemberOnType "OpenView" $db ('SELECT * FROM Property')
                        Invoke-MemberOnType "Execute" $view $null
                        $record = Invoke-MemberOnType "Fetch" $view $null
                    
                        while ($record -ne $null) 
                        {
                            $property = Invoke-MemberOnType "StringData" $record 1 "GetProperty"
                           switch ($property){
                                "ProductCode" {
                                        $value = Invoke-MemberOnType "StringData" $record 2 "GetProperty"
                                        Log-Item -logline "$property = $value"
                                        $WPFUninstallCmd.text = $value.Trim()
                                        $wpfUnInstallType.SelectedIndex = 2
                                }
                                "ProductVersion" {
                                        $value = Invoke-MemberOnType "StringData" $record 2 "GetProperty"
                                        Log-Item -logline "$property = $value"
                                        $WPFAppversion.text = $value.Trim()
                                }
                                "ProductName" {
                                        $value = Invoke-MemberOnType "StringData" $record 2 "GetProperty"
                                        Log-Item -logline "$property = $value"
                                        $wpfappname.text = $value.Trim()
                                }
                                "Manufacturer" {
                                        $value = Invoke-MemberOnType "StringData" $record 2 "GetProperty"
                                        Log-Item -logline "$property = $value"
                                        $wpfvendor.text = $value.Trim()
                                } 
                                "ProductLanguage"{
                                        $value = Invoke-MemberOnType "StringData" $record 2 "GetProperty"
                                        Log-Item -logline "$property = $value"
                                }
                           }
                            $record = Invoke-MemberOnType "Fetch" $view $null
                        }


                        $wpfinstalltype.SelectedIndex = 0
                }

                ".exe" {
                        Log-Item "EXE file chosen"
                        $FullVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($fullfilename).ProductVersion
                        $ProductLanguage = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($fullfilename).Language
                        $Manufacturer = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($fullfilename).CompanyName
                        $ProductName = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($fullfilename).ProductName
                        Log-item "Version : $FullVersion" 
                        $WPFAppversion.text = $FullVersion.Trim()
                        Log-Item "Language : $ProductLanguage"
                        Log-Item "Vendor : $Manufacturer"
                        $wpfvendor.text = $Manufacturer.Trim()
                        Log-Item "Name : $ProductName"
                        $wpfappname.text = $productname.Trim()
                        $wpfinstalltype.SelectedIndex = 1
                        $wpfUnInstallType.SelectedIndex = 1
                        
                }
        }
        
        $wpfinputsource.text = $filepath
        $wpfinstallpath.text = $filename
        $wpfinputSource.Tag = 'cleared' #if not cleared, will clear input source on click. (reported by Sandy)

        Fill-swid

}
Function Execute-Scriptblock
{
        Param($Remote,$SB, $sess)

       switch ($remote)
       {
        "Succeeded"
               {
                       write-host "succeded part"
                Invoke-Command -Session $sess -ScriptBlock $SB
               }
        "Failed"
               {
                log-item "Remote connection failed, not executing current scriptblock" -severity warn
               }
        "Local"
               {
                write-host "local part"
                Invoke-Command -ScriptBlock $SB
               }
       }
}

Function Connect-SCCM
{
        $Connection = $False
        
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfImport_progress.Content = "Connecting to SCCM Environment ..."} ),$null, $null )

        Try
                {
                        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop # Import the ConfigurationManager.psd1 module 
                        $global:Sitecode = (Get-PSDrive -PSProvider CMSite).Name
                        log-item ("SCCM Sitecode = $Sitecode") 
                        Set-Location ($global:Sitecode + ":") # Set the current location to be the site code.
                        #$WPFsitecode.Content = "Connected to Site : " + $global:Sitecode
                        $wpfSiteCode.background = "Green"
                        $Connection = $True
                        $WPFImport_SCCM.isenabled = $true
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFsitecode.Content = "Connected to Site : " + $global:Sitecode} ),$null, $null )
                        
                        if ($wpfCB_SiteCode.IsChecked)
                        {
                                $item = Select-XML -Xml $Preferences -XPath '//SiteCode[1]'
                                $item.Node.SC = $global:Sitecode
                                $Preferences.Save("$psscriptroot\config\Prefs.xml")
                                log-item "Updated Sitecode in Preferences XML"
                        }
                
                }
                Catch
                {
                        log-item ("Failed connecting to the SCCM Environment") -severity "Error"
                        #$wpfImport_progress.Content = "Failed connecting to the SCCM Environment"
                        $WPFConnect_SCCM.Background="Red"
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFsitecode.Content = "Failed connecting to the SCCM Environment"} ),$null, $null )
                # $WPFProgress_SCCM.Foreground="Red"
                # $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $WPFProgress_SCCM.value = 100; } ),$null, $null )
                }
                        
        if ($connection)
        {
                $wpfdp.items.Clear()
                $wpfdpgroup.items.Clear()
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfImport_progress.Content = "Enumerating Distribution Points"} ),$null, $null )
                $DPs = get-cmdistributionpoint -allsite
                foreach ($dp in $dps)
                {
                        $wpfdp.Items.Add($dp.NetworkOsPath.trimstart("\"))
                }
                $DPGs = Get-CMDistributionPointGroup
                foreach ($dpg in $dpgs)
                {
                        $wpfdpgroup.items.add($dpg.name)
                }
                #$wpfImport_progress.Content = "Distribution Points enumerated from site"
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfImport_progress.Content = "Distribution Points enumerated from site"} ),$null, $null )
        
                if ($wpfCB_Auto_Import.IsChecked)
                {
                        Import-SCCM  
                }
                if ($wpfCB_Auto_Collection.IsChecked)
                {
                        Generate-Collections
                }
        }
}

Function Generate-Detection{

        $appwithoutspaces = $wpfappname.text -replace " ","_"
        $appversionwithoutspaces = $WPFappVersion.text -replace " ",""

        $fullregid = "RegID." + $wpfregidYY.text + "-" + $wpfregidMM.text + "." + $wpfregiddomain.text

        $global:detectionscript = "if (test-path 'C:\ProgramData\" + $fullregid + "\" + $fullregid + "_" + $appwithoutspaces + "_" + $appversionwithoutspaces + "_" + $global:architecture +".swidtag') {write-host 'found swidtag'}"
        $global:intunedection =
@"
`$check = Test-Path -Path 'C:\ProgramData\$fullregid\$fullregid`_$appwithoutspaces`_$appversionwithoutspaces`_$global:architecture.swidtag'
if (`$check) {
        write-output "found swidtag"
        Exit 0
    } else {
        exit 1
    }
"@
        Log-Item ("Detection mechanism is : ")
        log-item ("$global:detectionscript")
}
Function Generate-IntunePackage{

        if (-not($WPFTB_IntuneFolder.text.Endswith("\")))
        {
                $WPFTB_IntuneFolder.text = $WPFTB_IntuneFolder.text + "\"
        }

        $IntunetargetPath = ($global:packagepath).Replace($global:dest,$WPFTB_IntuneFolder.text)

        if (!(test-path $IntunetargetPath))
        {
        log-item -logline ("Intune Application target folder does not exist yet, creating now.")
                New-Item -itemtype Directory -path $IntunetargetPath -Force
                
        }

        # Intune commandline
        #  IntuneWinAppUtil -c <setup_folder> -s <source_setup_file> -o <output_folder> <-q>
        log-item -logline ("Starting Intune package generation")
        #$IntunetargetPath = $WPFTB_IntuneFolder.Text
        $packagepath = $global:packagepath
        $arglist = " -c " + "`"" + $packagepath + "`"" + " -s " + "`"" + "Deploy-Application.exe" + "`"" + " -o " +"`"" + $IntunetargetPath + "`""
        log-item -logline ("Argument list : $arglist")
        $exitcode = Start-Process -FilePath "$PSScriptRoot\Intune_Win32\IntuneWinAppUtil.exe" -ArgumentList $arglist -Wait -PassThru
        log-item -logline ("exitcodee : $($exitcode.exitcode)")

        if ($exitcode.exitcode -eq 0)
        {
                $IntuneFile = Get-ChildItem -Path $IntunetargetPath -Filter ("*.intunewin")
                log-item -logline ("Generated Intune filename = $IntuneFile")
                $Newname = $wpfAppname.text + "_" + $wpfappversion.Text + ".intunewin"
                log-item -logline ("New Intune Filename = $Newname")
                Rename-Item -Path $IntunetargetPath\$IntuneFile -NewName $Newname
                
                Generate-Detection

                'Intune Install Commandline = Deploy-Application.exe -DeploymentType "Install"' | out-file -FilePath "$IntunetargetPath\Instructions.TXT"
                'Intune UnInstall Commandline = Deploy-Application.exe -DeploymentType "UnInstall"' | out-file -FilePath "$IntunetargetPath\Instructions.TXT" -Append
                'Intune Detection Method (Powershell - available in separate Detection_Method.PS1 file) = ' | out-file -FilePath "$IntunetargetPath\Instructions.TXT" -Append
                $global:intunedection | out-file -FilePath "$IntunetargetPath\Instructions.TXT" -Append
                $global:intunedection | out-file -FilePath "$IntunetargetPath\Detection_Method.PS1"

                [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                [Microsoft.VisualBasic.Interaction]::MsgBox("Intune package created with exitcode $($exitcode.exitcode). Instructions file Generated",'OKOnly,Information',"Intune Package")
             
        }
        else {
                [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                [Microsoft.VisualBasic.Interaction]::MsgBox("Intune package creation failed with exitcode $($exitcode.exitcode).",'critical',"Intune Package")
             
        }

        
}

Function Import-SCCM{
        $Appcreated = $false
        $appwithoutspaces = ""
        $appversionwithoutspaces = ""
        $fullregid = ""
        #$wpfImport_progress.Content = "Connected to SCCM Site $Sitecode, starting creation of application... Please wait"
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfImport_progress.Content = "Connected to SCCM Site $Sitecode, starting creation of application... Please wait"} ),$null, $null )
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $WPFProgress_SCCM.value = 10 } ),$null, $null )

        Generate-Detection
        
        if ($wpfappLanguage.text -ne "")
        {
                $global:fullname = $wpfappname.text + "_" + $wpfappVersion.text + "_" + $wpfappLanguage.text + "_" + $global:architecture
        }
        else {
                $global:fullname = $wpfappname.text + "_" + $wpfappVersion.text + "_" + $global:architecture
        }

        try 
        {
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfCollname.text = $global:fullname} ),$null, $null )
        New-CMApplication -Name $global:fullname -Manufacturer $wpfvendor.text -SoftwareVersion $wpfappversion.text -ErrorAction Stop
        #$wpfImport_progress.Content = "Application Created, adding SCCM Deployment Type... Please wait"
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $WPFProgress_SCCM.value = 50 } ),$null, $null )
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfImport_progress.Content = "Application Created, adding SCCM Deployment Type... Please wait"} ),$null, $null )
        log-item ("SCCM Application with name $fullname created")
        $appcreated = $true
        }
        catch
        {
            Log-Item -logline "Caught an exception:" -severity "Error"
            log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
            log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
                $wpfImport_progress.Foreground="red"  
                $wpfImport_progress.Content = "Error creating application : $($_.Exception.Message)"
                $WPFProgress_SCCM.Foreground="Red"
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $WPFProgress_SCCM.value = 100; } ),$null, $null )
        }

        if ($Appcreated)
        {
                try
                {
                Add-CMScriptDeploymentType -ApplicationName $global:fullname -ContentLocation $packagepath -DeploymentTypeName $global:fullname -InstallCommand "Deploy-Application.exe -DeploymentType `"Install`"" -UninstallCommand "Deploy-Application.exe -DeploymentType `"UnInstall`"" -Scriptlanguage powerShell -ScriptContent $global:detectionscript -InstallationBehaviorType InstallForSystemIfResourceIsDeviceOtherwiseInstallForUser -ErrorAction Stop
                log-item ("Deployment type with name $fullname created")
                #$wpfImport_progress.Content = "Done creating SCCM Application"
                
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfImport_progress.Content = "Done creating SCCM Application"} ),$null, $null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFProgress_SCCM.value = 100 } ),$null, $null )
                $wpfGenerate_Collections.Isenabled = $true
                $wpfdistribute.isenabled = $true
                }
                catch
                {
                    Log-Item -logline "Caught an exception:" -severity "Error"
                    log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
                    log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
                        $wpfImport_progress.Content = "Error creating Deployment type: $($_.Exception.Message)"
                        $WPFProgress_SCCM.Foreground="Red"
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $WPFProgress_SCCM.value = 100; } ),$null, $null )
                }

                Try 
                {
                if ($wpfCB_DP.IsChecked)
                {
                        Log-Item -logline "restoring selected dp's"
                        $storedDPs = $Preferences.Preferences.SelectedDP.DP
                        foreach ($dp in $storedDPs)
                        {
                                if ($dp.name -ne "")
                                        {
                                                $name = $dp.name
                                                Log-Item -logline "selecting $name"
                                                $wpfdp.SelectedItems.Add($dp.name)
                                        }
                        }
                }
                }
                catch
                {
                    Log-Item -logline "Failed restoring DP selection" -severity "Error"
                    Log-Item -logline "Caught an exception:" -severity "Error"
                    log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
                    log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
                    
                }

                Try 
                {
                if ($wpfCB_DPG.IsChecked)
                {
                        Log-Item -logline "restoring selected dp groups"
                        $storedDPGs = $Preferences.Preferences.SelectedDPGroup.DPG
                        
                        foreach ($dpg in $storedDPGs)
                        {
                                if ($dpg.name -ne "")
                                        {
                                                $name = $dpg.name
                                                Log-Item -logline "selecting $name"
                                                $wpfdpgroup.selecteditems.Add($dpg.name)
                                        }
                        }
                }
                }
                catch
                {
                    Log-Item -logline "Failed restoring DP Group selection" -severity "Error"
                    Log-Item -logline "Caught an exception:" -severity "Error"
                    log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
                    log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
                    
                }

               if ($wpfCB_Auto_Distribute.IsChecked)
                {
                        Distribute-Content
                }
                
                
        }
        
}

Function Distribute-Content{
        $TotalselectedDP = $wpfdp.selecteditems.count + $wpfdpgroup.selecteditems.count
        Log-item "Total number of selected Dp's and DP groups is : $TotalSelectedDP"

        If ($TotalselectedDP -gt "0")
        {
        [float]$DPBlocks = (100 / $TotalselectedDP)
        $dpprogress = 0

        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFdistribute_progress.Content = "Preparing $totalselecteddp distribution(s)"} ),$null, $null )
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFPB_Progress_distribute.Value = 0} ),$null, $null )
        foreach ($DP in $wpfdp.SelectedItems)
        {
                Try 
                {
                Start-CMContentDistribution -ApplicationName $fullname -DistributionPointName $dp -erroraction stop
                log-item -logline "Starting distribution to DP : $dp"
                $dpprogress = $dpprogress + $dpblocks
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFdistribute_progress.Content = "Starting distribution to DP : $dp"} ),$null, $null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFPB_Progress_distribute.Value = $dpprogress} ),$null, $null )
                

                if ($wpfCB_DP.IsChecked)
                {
                        $storedDP = $Preferences.Preferences.SelectedDP.DP | Where-object {$_.Name -eq $dp} | Select-Object -Property Name
                        if ($storeddp.Name -ne $DP)
			{
                                $item = Select-XML -Xml $Preferences -XPath '//DP'
                                $newnode = $item.Node.CloneNode($true)
                                $newnode.Name = $dp
                                $SelectedDP = Select-XML -Xml $Preferences -XPath '//SelectedDP'
                                $SelectedDP.Node.AppendChild($newnode)
                                $Preferences.Save("$psscriptroot\config\Prefs.xml")
                                log-item "Added $DP to favorite DP's in Preferences XML"
                        }
                        else {
                                log-item "$DP was already a favorite DP in Preferences XML"
                        }
                }
        
                }
                Catch 
                {
                        log-item -logline "Distribution to $dp failed" -severity error
                        log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
                        log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFdistribute_progress.Content = "Distribution to $dp failed"} ),$null, $null )
                }
        }
        foreach ($DPGroup in $wpfdpgroup.selecteditems)
        {
                try {
                        Start-CMContentDistribution -ApplicationName $fullname -DistributionPointGroupName $dpgroup -erroraction stop
                        log-item -logline "Starting distribution to DP group : $DPGroup"
                        $dpprogress = $dpprogress + $dpblocks
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFdistribute_progress.Content = "Starting distribution to DP group : $DPGroup"} ),$null, $null )
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFPB_Progress_distribute.Value = $dpprogress} ),$null, $null )
                        
                        if ($wpfCB_DPG.IsChecked)
                        {
                                $storedDPG = $Preferences.Preferences.SelectedDPGroup.DPG | Where-object {$_.Name -eq $DPGroup} | Select-Object -Property Name
                                if ($storedDPG.Name -ne $DPGroup)
                                {
                                        $item = Select-XML -Xml $Preferences -XPath '//DPG'
                                        $newnode = $item.Node.CloneNode($true)
                                        $newnode.Name = $DPGroup
                                        $SelectedDP = Select-XML -Xml $Preferences -XPath '//SelectedDPGroup'
                                        $SelectedDP.Node.AppendChild($newnode)
                                        $Preferences.Save("$psscriptroot\config\Prefs.xml")
                                        log-item "Added $DPGroup to favorite DP's in Preferences XML"
                                }
                                else {
                                        log-item "$DPGroup was already a favorite DP in Preferences XML"
                                }
                        }
                }
                catch {
                        log-item -logline "Distribution to $dpGroup failed" -severity error
                        log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
                        log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFdistribute_progress.Content = "Distribution to $dpGroup failed"} ),$null, $null )
                }

        }
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFdistribute_progress.Content = "All distributions have started"} ),$null, $null )
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFPB_Progress_distribute.Value = 100} ),$null, $null )
        }
        else {
                [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                [Microsoft.VisualBasic.Interaction]::MsgBox("No Distribution points or Distribution point groups are selected.",'OKOnly,Information',"Cannot start distribution")
                log-item -logline ("No Distribution points or Distribution point groups are selected") -severity Warn
        }
}

Function Generate-Collections{
        $collname_I = $global:fullname
        $collname_u = $global:fullname
        if ($wpfprefix.text -ne "")
                {
                        $Collname_i = $wpfprefix.text + "-" + $collname_i
                        $Collname_u = $wpfprefix.text + "-" + $collname_u

                        $item = Select-XML -Xml $Preferences -XPath '//Collection[1]'
                        $item.Node.Prefix = $wpfprefix.text
                        $Preferences.Save("$psscriptroot\config\Prefs.xml")

                }
        if ($wpfsuffix.text -ne "")
                {
                        $collname_i = $collname_i + "-" + $wpfsuffix.text
                        $collname_U = $collname_U + "-" + $wpfsuffix.text

                        $item = Select-XML -Xml $Preferences -XPath '//Collection[1]'
                        $item.Node.Suffix = $wpfsuffix.text
                        $Preferences.Save("$psscriptroot\config\Prefs.xml")
                }
        $collname_i = $collname_i + "-I"
        $collname_u = $collname_u + "-U"
        $coll_created = $false       
        try {
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFCollections_info.Content = "Start Creation of collections"} ),$null, $null )
                new-cmcollection -collectiontype device -limitingcollectionID $wpfTB_DefaultLimiting.Text -name $collname_I -erroraction stop
                new-cmcollection -collectiontype device -limitingcollectionID $wpfTB_DefaultLimiting.Text -name $collname_u -erroraction stop
                log-item -logline "Install collection : $collname_i created"
                log-item -logline "Uninstall collection : $collname_u created"
                $WPFCollections_info.Background = "Green"
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFCollections_info.Content = "Collection Creation has finished"} ),$null, $null )
                $coll_created = $true
        }
        catch {
                log-item -logline "failed creating collections" -severity error
                log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
                log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
                $wpfCollections_info.background = "Red"
                $WPFCollections_info.Content = "Error creating collections: $($_.Exception.Message)"
                #$Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFCollections_info.Content = "Error creating collections: $($_.Exception.Message)"} ),$null, $null )

        }

        if (($collname_I.Length -le 64) -and ($collname_u.Length -le 64))
        {     
                log-item -logline "AD group names are less than 65 characters, allowing AD group creation if needed"
                $ADGrouplengthOk = $True

        }
        else {
                $ADGrouplengthOk = $false
        }

        if ($wpfCB_Auto_ADGroup.Ischecked -and $ADGrouplengthOk)
        {
                $ADGroupscreated = $false
                Try {
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFCollections_info.Content = "Start Creation of AD Groups"} ),$null, $null )
                        New-ADGroup -Path $wpfTB_LDAP.Text -Name $collname_I -GroupScope DomainLocal  -GroupCategory Security -erroraction stop
                        New-ADGroup -Path $wpfTB_LDAP.Text -Name $collname_U -GroupScope DomainLocal  -GroupCategory Security -erroraction stop
                        log-item -logline "AD group $collname_i created."
                        log-item -logline "AD Group $collname_u created."
                        $WPFCollections_info.Background = "Green"
                        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFCollections_info.Content = "Creation of AD groups has finished"} ),$null, $null )
                        $ADGroupscreated = $True
                }
                catch {
                        log-item -logline "failed creating AD Groups" -severity error
                        log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
                        log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
                        $wpfCollections_info.background = "Red"
                        $WPFCollections_info.Content = "Error creating AD groups: $($_.Exception.Message)"
                        #$Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFCollections_info.Content = "Error creating AD groups: $($_.Exception.Message)"} ),$null, $null )
                }
        }
        elseif ($ADGrouplengthOk -eq $false) {
                log-item -logline "AD group names are more than 65 characters, Can't create AD groups. $($collname_I.Length) characters counted" -severity Error
                [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                [Microsoft.VisualBasic.Interaction]::MsgBox("AD group names are more than 65 characters, Can't create AD groups",'critical',"AD groups")
        }
        
        if ($coll_created -and $ADGroupscreated)
                {
                        try {
                                Log-item -logline "extracting domain details from OU path provided"
                                $domaindetails = $wpfTB_LDAP.Text
                                $domaindetails = $domaindetails.split(",")
                                $DN = ""
                                foreach ($item in $domaindetails)
                                        {
                                                if ($item.startswith("DC"))
                                                {
                                                        $DN = $DN + $item + ","
                                                }
                                        }
                                $DN = $DN.Trimend(",")
                                Log-Item -logline "Extracted Distringuished name is $DN"
                                $domain = Get-AdDomain -Identity $DN
                                $DomainName = $Domain.NetBiosName
                                Log-item -logline ("NetBios name is $DomainName")
                                $Query_I = 'Select * from SMS_R_System where SMS_R_System.SystemGroupName = "' + $DomainName +'\\' + $collname_i +'"'
                                $Query_U = 'Select * from SMS_R_System where SMS_R_System.SystemGroupName = "' + $DomainName +'\\' + $collname_U +'"'
                                log-item -logline "Domain name = $DomainName"
                                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFCollections_info.Content = "Adding AD groups as members of the collection"} ),$null, $null )
                                Add-CMDeviceCollectionQueryMembershipRule -CollectionName $collname_i -QueryExpression $Query_I -RuleName "AD group $collname_i" -ErrorAction Stop
                                Add-CMDeviceCollectionQueryMembershipRule -CollectionName $collname_u -QueryExpression $Query_U -RuleName "AD group $collname_u" -ErrorAction Stop
                                log-item -logline "AD group $collname_i added to collection  $collname_i"
                                log-item -logline "AD group $collname_u added to collection  $collname_u"
                                $WPFCollections_info.Background = "Green"
                                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFCollections_info.Content = "AD groups are now members of the collections"} ),$null, $null )
                        }
                        Catch {
                                log-item -logline "failed creating Collection membership rules for AD groups" -severity error
                                log-item -logline "Exception Type: $($_.Exception.GetType().FullName)" -severity "Error"
                                log-item -logline "Exception Message: $($_.Exception.Message)" -severity "Error"
                                $wpfCollections_info.background = "Red"
                                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFCollections_info.Content = "failed adding AD Groups as collection members"} ),$null, $null )


                        }
                }
}

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{log-item -logline "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed." -severity "Error"}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -scope global}
 
# Function Get-FormVariables{
# if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
# write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
# get-variable WPF*
# }
#  
# Get-FormVariables

#Fill form with data retrieved from preferences file
Prefill

[float]$blocks = (100 / $inputApptoolkit.Count)

[int]$progress = 0
$wpfprogress.Value = $progress
$WPFProgress_SCCM.value = 0

$WPFGenerate.add_Click({

Set-Location $global:Startlocation #set location out of SCCM environment
#$WPFPackage_progress.Content = "Valdating input ..."
$Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFPackage_progress.Content = "Valdating input ..."} ),$null, $null )

[int]$countpath = 0
        if(test-path $wpfInputSource.text)
        {
                $wpfInputSource.background = "green"
                $countpath += 1
                $writeoutput = $wpfInputSource.text
                Log-Item("$writeoutput is accessible and will be used for the source-binaries")
        }
        else 
        {
                $wpfInputSource.background = "red"
                $writeoutput = $wpfInputSource.text
                Log-Item -logline "$writeoutput is not accessible, script cannot continue" -severity "Error"
        }

        if (test-path $wpfInputDestination.text)
        {
                $wpfInputDestination.background = "green"
                $countpath += 1
                $writeoutput = $wpfInputdestination.text
                Log-Item("$writeoutput is accessible and will be used to copy the full App deployment package to")
        }
        else
        {
                $wpfInputDestination.background = "red"
                $writeoutput = $wpfInputdestination.text
                Log-Item -logline "$writeoutput is not accessible, script cannot continue" -severity "error"
        }

        if ($wpfvendor.text -eq "")
        {
                $wpfvendor.background = "red"
                Log-Item -Logline "Need a Vendor to continue" -Severity "Warn"
        }
        else {
                $wpfvendor.background = "green"
                $countpath += 1
        }

        if ($wpfappname.text -eq "")
        {
                $wpfappname.background = "red"
                Log-Item -Logline "Need an Application Name to continue" -Severity "Warn"
        }
        else {
                $wpfappname.background = "green"
                $countpath += 1
        }

        if ($wpfappversion.text -eq "")
        {
                $wpfappversion.background = "red"
                Log-Item -Logline "Need an Application Version to continue" -Severity "Warn"
        }
        else {
                $wpfappversion.background = "green"
                $countpath += 1
        }

        [Regex]$regidYY = "[1-2]\d{3}"
        [Regex]$regidMM = "[0-1][0-9]"
        [Regex]$regidDomain = "\w{2,}.\w+"
        $currentyear = (Get-Date).year
        
        
        if (!($WPFRegIDYY.text -in 1985..$currentyear))
        {
                $WPFRegIDYY.background = "red"
                Log-Item -Logline "SoftwareID-Tag : Year when domain-name was registered is missing or incorrect (must be between 1985 and $currentyear). Info can be found on www.whois.net" -Severity "Warn"
        }
        else {
                $WPFRegIDYY.background = "green"
                $countpath += 1
        }

        if (!($wpfregIDMM.text -in 1..12 -and $wpfregidMM.text -match $regidMM))
        {
                $wpfregidMM.background = "red"
                Log-Item -Logline "SoftwareID-Tag : Month when domain-name was registered is missing or incorrect (must be in 2-digit format, eg January = 01). Info can be found on www.whois.net" -Severity "Warn"
        }
        else {
                $wpfregidMM.background = "green"
                $countpath += 1
        }

        if (!($wpfregidDomain.text.toupper() -match $regidDomain))
        {
                $wpfregidDomain.background = "red"
                Log-Item -Logline "SoftwareID-Tag : Domain-name is missing." -Severity "Warn"
        }
        else {
                $wpfregidDomain.background = "green"
                $countpath += 1
        }

        if ($WPFLicense.SelectedIndex -eq "-1")
        {
                $wpfLicenseText.background = "red"
                Log-Item -Logline "Select Yes/No for License Requirements" -Severity "Warn"
        }
        else {
                $wpfLicenseText.background = "green"
                $countpath += 1
        }

if ($countpath -eq 9)
{
        
        Check_and_Save_xml -name $wpfVendor.text -year $WPFRegIDYY.text -month $wpfregIDMM.text -domain $wpfregidDomain.text
        Update_Prefs -SP $wpfInputSource.text -DP $wpfInputDestination.text -EM $wpfTB_EndMessage.Text

        $FullRegid = "RegID." + $wpfregidYY.text + "-" + $wpfregidMM.text + "." + $wpfregiddomain.text
        #$WPFPackage_progress.Content = "Copying toolkit files to destination folder ... (This can take a while)"
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$WPFPackage_progress.Content = "Copying toolkit files to destination folder ... (This can take a while)"} ),$null, $null )
        
        Compose-Package -source $wpfinputsource.text -destination $wpfInputDestination.text -appvendor $wpfvendor.text -appname $wpfappname.text -appversion $wpfappversion.text -appbitness $wpfApparchitecture.SelectedIndex

        $output = $global:packagepath + "\Deploy-Application.ps1"
       
        #$wpfpackage_progress.Content = "Generating Deploy-Application.PS1 file ..."
        $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfpackage_progress.Content = "Generating Deploy-Application.PS1 file ..."} ),$null, $null )
        $writeoutput = "### Generated FILE from GUI with GUIToolkit version $gui_version ###"
        $writeoutput | out-file $output
        for ($i=0; $i -lt  $inputApptoolkit.Count; $i++)
                {
                $defaultline = $false
                $Linefrominput =  $inputApptoolkit[$i].tostring().trim()
                Switch ($Linefrominput)
                        {
                        "<GUIAppVendor>"
                                {
                                        $writeoutput = '	[string]$appVendor = ''' + $wpfvendor.text + "'"
                                }
                        "<GUIAppVendor_WS>"
                                {
                                        $writeoutput = '	[string]$appVendor_WithSpace = ''' + $wpfvendor.text + "'"
                                }   
                        "<GUIAppName>"
                                {
                                        $writeoutput = '	[string]$appName = ''' + $wpfAppname.text + "'"
                                }
                        "<GUIAppName_WS>"
                                {
                                        $writeoutput = '	[string]$appName_WithSpace = ''' + $wpfAppname.text + "'"
                                }
                        "<GUIAppVersion>" 
                                {
                                        $writeoutput = '	[string]$appVersion = ''' + $wpfAppversion.text + "'"
                                }
                        "<GUIAppArch>"
                                {Switch ($wpfApparchitecture.SelectedIndex)
                                {
                                "0"
                                        {$writeoutput = '	[string]$appArch = ''X86'''
                                        }
                                "1"
                                        {$writeoutput = '	[string]$appArch = ''X64'''
                                        }
                                }
                                }
                        "<GUIAppLang>"
                                {$writeoutput = '	[string]$appLang = ''' + $wpfapplanguage.text + "'"
                                }
                        "<GUIAppRevision>"
                                {$writeoutput = '	[string]$appRevision = ''' + $wpfapprevision.text + "'"
                                }
                        "<GUIRegID>"
                                {$writeoutput = '	[string]$appVendorSWIDTAGRegID = ''' + $FullRegid + "'"
                                }
                        "<GUIScriptVersion>"
                        {$Writeoutput = '	[string]$appScriptVersion = ''' + '1.0.0' + "'"
                        }
                        "<GUIAppCreator>"
                        {$username = [Security.Principal.WindowsIdentity]::GetCurrent().Name
                        $Writeoutput = '	[string]$appScriptAuthor = ''' + $username + "'"
                        }
                        "<GUIAppdate>"
                                {$date = get-date -format dd/MM/yyyy 
                                $Writeoutput = '	[string]$appScriptDate = ''' + $date + "'"
                                }
                        "<GUIInstallCmdline>"    
                                {   Switch ($WPFInstallType.SelectedIndex)
                                {
                                "0" 
                                        {
                                                if ($($WPFInstallPath_Parameters.Text) -ne "")
                                                {
                                                        log-item -logline ("MSI parameter included,adding it to commandline")
                                                        $writeoutput = '		Execute-MSI -Action ''Install'' -path ''' + $wpfinstallpath.Text + "' -Parameters '" + $WPFInstallPath_Parameters.Text + "'"
                                                }
                                                else 
                                                {
                                                        log-item -logline ("NO MSI parameter included, skipping part of commandline")
                                                        $writeoutput = '		Execute-MSI -Action ''Install'' -path ''' + $wpfinstallpath.Text + "'"
                                                }
                                       }
                                "1" 
                                        {
                                                if ($($WPFInstallPath_Parameters.Text) -ne "")
                                                {
                                                        log-item -logline ("Script parameters included,adding it to commandline")
                                                        $writeoutput = '		Execute-Process -path ''' + $wpfinstallpath.Text + "' -Parameters '" + $WPFInstallPath_Parameters.Text + "'"
                                                }
                                                else 
                                                {
                                                        log-item -logline ("Script parameters Not included,skipping part of commandline")
                                                        $writeoutput = '		Execute-Process -path ''' + $wpfinstallpath.Text + "'"   
                                                }
                                         }
                                } 
                                }
                        "<GUIUnInstallCmdline>"
                                { Switch ($WPFUnInstallType.SelectedIndex)
                                {
                                "0" 
                                        {
                                        $writeoutput = '		Remove-MSIApplications -Name ''' + $WPFUninstallCmd.Text + "'"
                                        }
                                "1"
                                        {
                                        $writeoutput = '		Execute-Process -path ''' + $WPFUninstallCmd.Text + "'"
                                        }
                                "2"
                                        {
                                        $writeoutput = '		Execute-MSI -Action ''' + 'Uninstall'  + "'" + ' -Path ''' + $WPFUninstallCmd.Text + "'"
                                        }
                                }
                                } 
                        "<GUIWriteIDTag>"
                                { 
                                Switch ($WPFLicense.SelectedIndex)
                                        {
                                        "1" { 
                                        $writeoutput =
@"
		## Writes unique file in ProgramData folder
		Write-Log -Message "Create and Write SWIDTag in ProgramData Folder" -Source 'Install' -LogType 'CMTrace'
		Write-SWIDTag -CreatorRegid `$appVendorSWIDTAGRegID -EntitlementRequired `$FALSE
		if ((get-CimClass -ClassName "sms_client" -Namespace root\ccm -ErrorAction SilentlyContinue).count -gt 0)
                {Invoke-SCCMTask 'HardwareInventory'}
                else
                {
                        Write-Log -Message "No SCCM Client WMI namespace found, not triggering hardware inventory" -Source 'Install' -LogType 'CMTrace'
                } 
"@

                                        }
                                "0"
                                        {
                                        $writeoutput = 
@"
		## Writes unique file in ProgramData folder
		Write-Log -Message "Create and Write SWIDTag in ProgramData Folder" -Source 'Install' -LogType 'CMTrace'
		Write-SWIDTag -CreatorRegid `$appVendorSWIDTAGRegID -EntitlementRequired `$TRUE
		if ((get-CimClass -ClassName "sms_client" -Namespace root\ccm -ErrorAction SilentlyContinue).count -gt 0)
                {Invoke-SCCMTask 'HardwareInventory'}
                else
                {
                        Write-Log -Message "No SCCM Client WMI namespace found, not triggering hardware inventory" -Source 'Install' -LogType 'CMTrace'
                } 
"@
                                        }
                                
                                }
                        }
                        "<GuiRemoveIDTag>"
                                {
                                        $writeoutput = 
@"
                ## Removes unique file from ProgramData folder
                `$appversionwithoutspaces = `$AppVersion -replace " ",""
                `$Appwithoutspaces = `$appName_WithSpace -replace " ", "_"
		Write-Log -Message "Remove SWIDTAG File" -Source 'UnInstall' -LogType 'CMTrace'
                `$tag = `$env:ProgramData + '\' + `$appVendorSWIDTAGRegID + '\' + `$appVendorSWIDTAGRegID + '_' + `$Appwithoutspaces + '_' + `$appversionwithoutspaces + '_' + `$AppArch +'.swidtag'
		Remove-File -Path `$Tag
		Invoke-SCCMTask 'HardwareInventory'
"@

                                }
                        "<GuiEndMessage>"
                        {
                                If ($wpfTB_EndMessage.Text -eq "")
                                {
                                $Writeoutput = "        # Show-InstallationPrompt -Message 'You can customize text to appear at the end of an install or remove it completely for unattended installations.' -ButtonRightText 'OK' -Icon Information -NoWait"
                                }
                                Else
                                {
                                        $Writeoutput = "        Show-InstallationPrompt -Message '" + $wpfTB_EndMessage.Text + "' -ButtonRightText 'OK' -Icon Information -NoWait"
                                }
                        }
                        default 
                                {
                                        $writeoutput = $inputApptoolkit[$i]
                                        $defaultline = $true       
                                }
                        }

                $writeoutput | out-file $output -append
                if ($defaultline -eq $false)
                {log-Item("Adding $writeoutput on line $i")}
                $progress += $blocks
                
                #$wpfprogress.Dispatcher.Invoke("Render", ([Action]{$wpfprogress.Value = $progress}),"Normal",$null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $wpfprogress.Value = $progress; } ),$null, $null )
              
               
                
                }
                $wpfprogress.Value = 100
                $WPFSCCM.isenabled = $true

                if ($wpfCB_Intunepackage.IsChecked)
                {
                        Generate-IntunePackage
                }

                if ($wpfCB_Auto_Sccm.IsChecked)
                        {
                                $WPFSCCM.IsSelected = $True
                                Connect-SCCM
                        }
                
                #$wpfsccm.background = "green"
                #$wpfpackage_progress.Content = "Finished creating package on destination location"
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfpackage_progress.Content = "Finished creating package on destination location"} ),$null, $null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{$wpfsccm.background = "green" } ),$null, $null )
                $Form.Dispatcher.Invoke( "Render", ([Action``2[[System.Object,mscorlib],[System.Object, mscorlib]]]{ $wpfprogress.Value = $progress } ),$null, $null )
}               
})

$WPFConnect_SCCM.add_click(
        {
                Connect-SCCM
        }
)

$WPFImport_SCCM.add_Click(
{
       Import-SCCM
})

$wpfdistribute.add_click(
{
        Distribute-Content
})

$wpfgenerate_collections.add_click({

       Generate-Collections
}) 

$wpfinstalltype.add_Selectionchanged(
{
Switch ($WPFInstallType.SelectedIndex)
    {
    "0" 
            {
            $wpfinstalltext1.text = "MSI Filename"
            $WPFInstallText2.text = "MSI Parameters"
            }
    "1" 
            {
            $wpfinstalltext1.text = "Script executable"
            $WPFInstallText2.text = "Script Parameters" 
            }
    } 
})

$wpfUnInstallType.add_Selectionchanged(
        {
                Switch ($WPFUnInstallType.SelectedIndex)
                {
                        "0"
                        {
                                $WPFUninstallTxt.text = "Application name in Add/Remove programs to uninstall"
                                $WPFUninstallCmd.tooltip = "if you put in 'Adobe', it removes all versions of software that match the name 'Adobe'"
                        }
                        "1"
                        {
                                $WPFUninstallTxt.text = "Full Uninstall Command-Line"
                                $WPFUninstallCmd.tooltip = "Please provide the full uninstall Command-Line" 
                        }
                        "2"
                        {
                                $WPFUninstallTxt.text = "MSI Product Code"
                                $WPFUninstallCmd.tooltip = "Please provide the MSI Product Code" 
                        }
                }

        }
)

$wpfBTN_Lookup.add_click({
        $url = $wpfVendorLookup.text
        $url = $url.toupper()
        
        if ($url.substring(0,3) -eq "WWW")
        {
                $url = $url.trimstart("W.")
                lookup-url $url
        }
        else 
        {
                lookup-url $url
        }
})

$wpfVendorLookup.add_gotfocus({
        if($wpfVendorLookup.Tag -ne 'cleared') 
        { # clear the text box
                $wpfVendorLookup.Text = ''
                $wpfVendorLookup.Tag = 'cleared' # use any text you want to indicate that its cleared
        }
})

$wpfinputSource.add_gotfocus({

        if($wpfinputSource.Tag -ne 'cleared') 
        { # clear the text box
        $wpfinputSource.Text = ''
        $wpfinputSource.Tag = 'cleared' # use any text you want to indicate that its cleared
        }
})


$wpfinputdestination.add_gotfocus({

        if($wpfinputdestination.Tag -ne 'cleared') 
        { # clear the text box
        $wpfinputdestination.Text = ''
        $wpfinputdestination.Tag = 'cleared' # use any text you want to indicate that its cleared
        }
})

$wpfregidYY.add_gotfocus({
        if($wpfregidYY.Tag -ne 'cleared') 
        { # clear the text box
        $wpfregidYY.Text = ''
        $wpfregidYY.Tag = 'cleared' # use any text you want to indicate that its cleared
        }
})

$wpfregidMM.add_gotfocus({
        if($wpfregidMM.Tag -ne 'cleared') 
        { # clear the text box
        $wpfregidMM.Text = ''
        $wpfregidMM.Tag = 'cleared' # use any text you want to indicate that its cleared
        }
})

$wpfregidDomain.add_gotfocus({
        if($wpfregidDomain.Tag -ne 'cleared') 
        { # clear the text box
        $wpfregidDomain.Text = ''
        $wpfregidDomain.Tag = 'cleared' # use any text you want to indicate that its cleared
        }
})

$wpfCB_Intunepackage.Add_Click({
        if ($wpfCB_Intunepackage.IsChecked)
        {
                $WPFTB_IntuneFolder.IsEnabled = $true
        }
        else {
                $WPFTB_IntuneFolder.IsEnabled = $false
        }
})

$wpfBTN_SaveSettings.add_click({
        
        $SaveOKIntune = $False
        $SaveOKAD = $False

        if ($wpfCB_SourcePath.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_SP = "True"
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_SP = "False"
        }
        if ($wpfCB_DestinationPath.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_DP = "True"
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_DP = "False"
        }   
        if ($wpfCB_SiteCode.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_SiteCode = "True"
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_SiteCode = "False"
        }   
        if ($wpfCB_DP.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_DistPoint = "True"
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_DistPoint = "False"
        }   
        if ($wpfCB_DPG.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_DistGroup = "True"
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.Save_DistGroup = "False"
        }   

        if ($wpfCB_Auto_Sccm.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.Autoconnect = "True"
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.Autoconnect = "False"
        } 
        if ($wpfCB_Auto_Import.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.AutoImport = "True"
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.AutoImport = "False"
        } 
        if ($wpfCB_Auto_Distribute.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.AutoDistribute = "True"
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.AutoDistribute = "False"
        } 
        if ($wpfCB_Auto_Collection.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.AutoCollection = "True"
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.AutoCollection = "False"
        } 
        if ($wpfCB_Auto_ADGroup.IsChecked){
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.AutoADGroup = "True"
               # $wpfCB_CreateAD.IsChecked = $True
        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                $item.Node.AutoADGroup = "False"
               # $wpfCB_CreateAD.IsChecked = $false
        } 
        
        #add node for IntunePackage to XML if it does not exist yet
        $item = Select-XML -Xml $Preferences -XPath '//Settings'
        if ($item.Node.IntunePackage -eq $null)
        {
                log-item -logline ("Adding node for IntunePackage")
                $xmlElt = $Preferences.CreateElement("IntunePackage")
                $Preferences.GetElementsByTagName("Settings")[0].AppendChild($xmlElt)
           
        }
        else {
                log-item -logline ("Node for IntunePackage already existed.")
        }

        #add node for Intune Appfolder to XML if it does not exist yet
        $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs'
        if ($item.Node.IntuneAppfolder -eq $null)
        {
                log-item -logline ("Adding node for Intune Appfolder")
                $xmlElt = $Preferences.CreateElement("IntuneAppfolder")
                $Preferences.GetElementsByTagName("SCCM_Prefs")[0].AppendChild($xmlElt)
                   
        }
        else {
                log-item -logline ("Node for Intune Appfolder already existed.")
        }


        $intunepathok = $false
        if ($wpfCB_Intunepackage.IsChecked)
        {
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.IntunePackage = "True"
                if (Test-Path $WPFTB_IntuneFolder.Text) 
                {
                        $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
                        $item.Node.IntuneAppfolder = $WPFTB_IntuneFolder.Text
                        log-item "Intune target path exists, continuing"
                        $intunepathok = $true
                        $SaveOKIntune = $True
                        $WPFTB_IntuneFolder.Background = "green"
                }
                else {
                        log-item "Intune target path does not exist, unable to save preferences" -severity error
                        $intunepathok = $false
                        $SaveOKIntune = $False
                        $WPFTB_IntuneFolder.Background = "red"
                }
                

        }
        else {
                $item = Select-XML -Xml $Preferences -XPath '//Settings[1]'
                $item.Node.IntunePackage = "False"
                $SaveOKIntune = $True
        }

        $item = Select-XML -Xml $Preferences -XPath '//SCCM_Prefs[1]'
        $item.Node.LimitingCollection = $wpfTB_DefaultLimiting.Text
        

        if ($wpfCB_Auto_ADGroup.IsChecked) 
        {
                $exists = $false
                log-item -logline "Checking if OU : $($wpfTB_LDAP.Text) exists."
                $exists = $(try { Get-ADOrganizationalUnit -Identity $wpfTB_LDAP.Text -ErrorAction Ignore } catch{}) -ne $null
                if ($exists)
                {
                        log-item -logline "OU Exists, continuing"
                        $item.Node.LDAP = $wpfTB_LDAP.Text
                        $wpfTB_LDAP.Background = "green"
                        $SaveOKAD = $True
                }
                else {
                        log-item -logline "OU does not Exist, not saving details" -severity error
                        $wpfTB_LDAP.Background = "red"
                }
                
        }
        else {
                $SaveOKAD = $True
        }

         #$item.Node.IntunePath = $wpfTB_IntuneFolder.Text
        if (($SaveOKIntune) -and ($SaveOKAD))
        {
                $Preferences.Save("$psscriptroot\config\Prefs.xml")
                [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                [Microsoft.VisualBasic.Interaction]::MsgBox("Settings Saved",'OKOnly,Information',"Settings Update")
                log-item "Settings saved"
        }
        else {
                [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
                [Microsoft.VisualBasic.Interaction]::MsgBox("Unable to save settings",'critical',"Settings Update")
                log-item "Unable to save settings"
        }

})

$WPFBTN_FilePicker.add_click({
        File-Picker
})

$wpfVendor.add_Lostfocus({
        Fill-swid
})

$form.add_closing({
        Set-Location $global:Startlocation
        log-item -logline "Exiting application"
})
$form.title = "Powershell App deployment Tookit - GUI - Version $Gui_version"
$Form.ShowDialog() | out-null

#$wpfVendor | Get-Member -MemberType Event
#$wpfVendor | Get-member Add* -MemberType Method -force
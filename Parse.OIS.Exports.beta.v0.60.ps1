<#	
	.NOTES
	===========================================================================
	OIS Export Parser - HTML Document Generator v0.60
	===========================================================================
	 Created by:   	Michael Adams <michael_adams@outlook.com>
	 Organization: 	http://willcode4foodblog.wordpress.com
	 Created on:   	8/18/2014 7:01 PM
	 Last Updated:	10/26/2014 9:00 PM 
	 
	 Filename:     	Parse.OIS.Exports.beta.v0.60.ps1
	====================
	Change Log:
	====================
	Version:		0.60 - Rewrote Table of Contents to output to index.html vs. "activitylist.html"
						 - Fixed CSS Menu Bugs
						 - Included Master Runbook in export 
						 - Toned down (commented out) some unnecessarily verbose logging 
						 - 
						 - Corrected companylogo image issue
						 - Introduced ShowNullorEmptyProperties Configuration Item (General Section)
	--------------------	 
	 Version:		0.50 - Rewrote code to properly process multiple runbooks within a folder
						 - Introduced Left Side CSS Menu Problem that needs to be fixed.
						 - Introduced Link CSS Menu bug where links aren't properly identified
	--------------------	 
	 Version:		0.40 - Added Drop Down for "Execution Data" (Available Options on an Activity)
						 - Removed "areaLinkSource" green mappings on the Index.html files
	--------------------	 
	 Version:		0.35 - Beta Code - Added Human-Readable: 
						 - GUIDs 
						 - Link Conditions
						 - Returned Data
						 - Custom Start Parameters
	--------------------						 
	 Version:		0.30 - Beta Code - Fixed GUID Replacement Functionality
	--------------------	 
	 Version:		0.25 - Beta Code - Fixed Links to proper url on left menu
	-------------------- 
	 Version:		0.20 - Beta Code - Ready to Demo
	--------------------	 
	 Version:		0.10 - Beta Code - DO NOT RUN OUTSIDE OF SPECIALIZED ENVIRONMENT

	 
	Outstanding Issues:
	--------------------
		- Requires manual creation of icons / maps
		- No UI

	===========================================================================
	.DESCRIPTION
		This script is intended to help an Orchestrator Administrator in creating understandable
		documentation for a runbook.  This script will take a runbook export (OIS) file,
		and by utilizing that information - as well as a few queries to the SCORCH database - will
		generate a navigatable set of HTML files, that use CSS and JQUERY to provide an interactive
		"Map" of the runbook and the corrosponding objects/activities in that runbook. 
		
		When a user clicks on the activity icon in one of the "Index.html" files, they will
		be sent to a web page created specifically for that activity.  Each activity will be highlighted
		on an (embedded into the html file) image of the current runbook map, and will show a user the following:
		
			- A CSS drop down menu containing links to each "Link" Activity connected to/from the current Activity
			- Properties for each activity (Name, ObjectType, Timeout values, etc.)
			- Any configured data from the user in a format that is readable
			- Format Script Code from .NET activities with CSS to allow easy reading with comments "highlighted"
			- Format SQL Code in the same manner as .NET Script Activities
			- Any Custom Start Parameters from an Initialize Data Activity
			- Any Data sent to the databus from a "Return Data" activity
			- Any Published variables from Script Activities.
			- Object GUIDs converted to human-readable text
			- Link Conditions converted to human-readable "psuedo functions"
			
	.REQUIREMENTS
		The user of this script must be able to complete tasks that may include:
			- Using the built in windows "Snip" utility to take shots of Activity Icons from within the SCORCH Console
			- Using the built in windows "Snip" utility to take shots of Runbook Workflows from within the SCORCH Console
			- Read access to the Orchestrator Database (may be remote if integrated auth is true)
			- Exporting runbook information from SCORCH from within the SCORCH Designer Console
		To change themes (colors of generated files)
			- A moderate understanding of CSS Files and using "classes"
		
#>
# Note: if the above line is commented out the line below it will use the dynamic way to obtain the path
if($ScriptDirectory -eq $null) {
	$script:ScriptDirectory = split-path -parent $MyInvocation.MyCommand.Definition;
}

#region Load Configuration File
# This function needs to come first so we can load the global variables below
# based on the configuration file information
Function Read-ConfigFile($configFile) {	
<#
	.SYNOPSIS
		Function: Read-ConfigFile([string]configFile)

	.DESCRIPTION
		Reads configuration files that are formatted like an INI
		
		Such as:
			
			[General]
			MyPreference1=1
			MyPreference2=0
			
			[Logs]
			LogFilePath=C:\users\me\desktop\logs
			VerboseLogging=1
					
	.PARAMETER  configFile
		[string] String path to file with INI type format

	.EXAMPLE
		PS C:\> $mySettings = Read-ConfigFile "C:\scripts\myscript\myscript.config"
		PS C:\> Load-ConfigurationSettings $mySettings # Load our global settings into the script from the config file

	.LINK
		Source:  http://blogs.technet.com/b/heyscriptingguy/archive/2011/08/20/use-powershell-to-work-with-any-ini-file.aspx
#>
	$config = @{}
	switch -regex -file $configFile
	{
		"^\[(.+)\]" # Section
		{
			$section = $matches[1]
			$config[$section] = @{}
			$CommentCount = 0
		}
		"^(;.*)$" # Comment
		{
			$value = $matches[1]
			$CommentCount = $CommentCount + 1
			$name = "Comment" + $CommentCount
			$config[$section][$name] = $value
		} 
		"(.+?)\s*=(.*)" # Key
		{
			$name,$value = $matches[1..2]
			$config[$section][$name] = $value
		}
	}
	
	return $config
}# End Function Read-ConfigFile

#endregion

#region Global Variables

# Here we declare all paths to files and folders that have resources that are used in the script

# First load all configuration data from the config file we have in the script directory
$script:config = Read-ConfigFile "$($ScriptDirectory)\Parse.OIS.Exports.config"

# Variables that we will need throughout
# Misc Global variables we will be using with some of the functions
[string]$script:SubRunbookName
[string]$script:fileName

$script:objGlobalRunbookVariables;
$script:thisRunbook = $null
$script:sourceHTML = $null;
$script:targetHTML = $null;

$script:SCORCHDBServer = $config["SQLConnectionInfo"]["SCORCHDBServer"]
$script:GUIDRegex = ("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$") 

# Awesome Scriptblock I found on stackOverflow that lets me sort the file names using the digit(s) in the file name as an index number
$ToNatural = { [regex]::Replace($_, '\d+', { $args[0].Value.PadLeft(20) }) }

# Create the SQL Connection String depending on the config file settings
if($config["SQLConnectionInfo"]["INTEGRATEDAUTH"] -eq 0) {
	
	$sqlUserID = $config["SQLConnectionInfo"]["SQLACCOUNT"]
	$sqlIP = $config["SQLConnectionInfo"]["DBIPADDRESS"]
	$sqlPW = $config["SQLConnectionInfo"]["SQLPASSWORD"]
	$sqlPort= $config["SQLConnectionInfo"]["PORT"]
	$initialCatalog = $config["SQLConnectionInfo"]["DBNAME"]

	# Note: 
	# If you specify a port number other than 1433 when you are trying to connect to an instance of SQL Server 
	# and using a protocol other than TCP/IP, the Open method fails. To specify a port number other than 1433, 
	# include "server=machinename,port number" in the connection string, and use the TCP/IP protocol.
	# Source: http://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqlconnection.open(v=vs.110).aspx
	
	$script:SQLConnectionString = "Server=$($SCORCHDBServer),$($sqlPort);Network=DBMSSOCN;Initial Catalog='$($initialCatalog)';uid=$($sqlUserID);pwd=$($sqlPW);"
} else {
	$script:SQLConnectionString = "Server=$($SCORCHDBServer),$($sqlPort);Network=DBMSSOCN;Initial Catalog='$($initialCatalog)';Integrated Security=True;"
}

# Load the rest of the configuration settings so all the script pieces have access to it
$script:path = $config["General"]["OISExportFile"];

$script:ShowNullorEmptyProps = $config["General"]["ShowNullorEmptyProperties"];
$script:ReportExportPath = $config["General"]["ReportExportPath"]
$script:PictureRepository = $config["General"]["PictureRepository"]
$script:strCompanyLogo = $config["General"]["CompanyLogo"]
$script:templateReport = $config["Templates"]["templateReport"]
$script:PoshInfoTemplate = $config["Templates"]["PoshInfoTemplate"]
$script:PoshPublishedDataTemplate = $config["Templates"]["PoshPublishedDataTemplate"]
$script:LinkActivityTemplate = $config["Templates"]["LinkActivityTemplate"]
$script:SQLInfoTemplate = $config["Templates"]["SQLInfoTemplate"]
$script:GeneralActivityTemplate = $config["Templates"]["GeneralActivityTemplate"]
$script:TableOfContentsHeader = $config["Templates"]["TableOfContentsHeader"]
$script:TableOfContentsFooter = $config["Templates"]["TableOfContentsFooter"]
$script:CSSMenuTemplate = $config["Templates"]["CSSMenuTemplate"]
$script:JQueryMenuHighlightScript = $config["Templates"]["JQueryMenuHighlightScript"]
$script:StyleSheetPath = $config["CSS"]["StyleSheetPath"]

#endregion

Write-Host "`t`t===========================================================================" -ForegroundColor White
Write-Host "`t`t`t`tOIS Export Parser - HTML Document Generator v0.60" -ForegroundColor Yellow
Write-Host "`t`t===========================================================================" -ForegroundColor White
Write-Host "`t`t`tCreated by:" -ForegroundColor White -NoNewLine
Write-Host "`t`t`tMichael Adams" -ForegroundColor Yellow
Write-Host "`t`t`tOrganization:" -ForegroundColor White -NoNewLine
Write-Host "`t`t`thttp://willcode4foodblog.wordpress.com" -ForegroundColor Yellow
Write-Host "`t`t`tCreated on:" -ForegroundColor White -NoNewLine  	
Write-Host "`t`t`t08/18/2014 7:01 PM" -ForegroundColor Yellow
Write-Host "`t`t`tLast Updated:" -ForegroundColor White -NoNewLine  		
Write-Host "`t`t`t10/26/2014 9:00 PM" -ForegroundColor Yellow
Write-Host "`t`t===========================================================================" -ForegroundColor White

#region Functions for Script
Function Read-ConfigFile($configFile) {	
<#
	.SYNOPSIS
		Function: Read-ConfigFile([string]configFile)

	.DESCRIPTION
		Reads configuration files that are formatted like an INI
		
		Such as:
			
			[General]
			MyPreference1=1
			MyPreference2=0
			
			[Logs]
			LogFilePath=C:\users\me\desktop\logs
			VerboseLogging=1
					
	.PARAMETER  configFile
		[string] String path to file with INI type format

	.EXAMPLE
		PS C:\> $mySettings = Read-ConfigFile "C:\scripts\myscript\myscript.config"
		PS C:\> Load-ConfigurationSettings $mySettings # Load our global settings into the script from the config file

	.LINK
		Source:  http://blogs.technet.com/b/heyscriptingguy/archive/2011/08/20/use-powershell-to-work-with-any-ini-file.aspx
#>
	$config = @{}
	switch -regex -file $configFile
	{
		"^\[(.+)\]" # Section
		{
			$section = $matches[1]
			$config[$section] = @{}
			$CommentCount = 0
		}
		"^(;.*)$" # Comment
		{
			$value = $matches[1]
			$CommentCount = $CommentCount + 1
			$name = "Comment" + $CommentCount
			$config[$section][$name] = $value
		} 
		"(.+?)\s*=(.*)" # Key
		{
			$name,$value = $matches[1..2]
			$config[$section][$name] = $value
		}
	}
	
	return $config
}# End Function Read-ConfigFile

function Generate-RunbookMenuSelections($oisExportXML, $ReportExportPath)
{
	# Find every runbook in the export:
	$strLeftMenuHTML = $null;
	$arrRunbookRootFolders = @();
		
	# Get the root folder that contains all of these runbooks:
	$rootRunbookDirectory = $oisExportXML.ExportData.Policies.Folder.Name.Replace(' ',$null);
	$arrRunbookRootFolders += $oisExportXML.ExportData.Policies.Folder
	$arrRunbookRootFolders += $oisExportXML.ExportData.Policies.Folder.Folder
	
	Log-Message "[ Debug-MainScript-ToCLeftMenu ]: Beginning Process on $($arrRunbookRootFolders.Count) Runbook Folders..."
	# This UL tag defines a runbook (header) (example: '1.00 - Staging')
	$strLeftMenuHTML += "<ul>"
	
	
	$objAllActivities = @();	

	
	# for each folder found:
	$arrRunbookRootFolders | % {	
		$objRunbooks = @();
		$objRunbooks += $_.Policy
		
		# for each runbook found
		$objRunbooks | % {
		
			$thisRootRunbookPolicy = $_
			$thisRootRunbookName = $thisRootRunbookPolicy.Name.'#text';
			$thisRootRunbookFolderName = $thisRootRunbookName.Replace(' ', $null)
			
			Log-Message "[ Debug-MainScript-ToCLeftMenu ]: Beginning to build list of activities for menu"
				
			if($num -eq 0) {
				$strLeftMenuHTML += "<li class='Active'><a href='..\$($thisRootRunbookFolderName)\Index.html'><span>$($thisRootRunbookName)</span></a>"
			} else {
				$strLeftMenuHTML += "<li class='has-sub'><a href='..\$($thisRootRunbookFolderName)\Index.html'><span>$($thisRootRunbookName)</span></a>"
			}
			
			if ($thisRootRunbookPolicy.Object.Count -gt 0)
			{
				# Begin building our Left side menu
				$arrFilesInDirectory = @();
				$objAllActivities += $thisRootRunbookPolicy.Object
				Log-Message "[ Debug-MainScript-ToCLeftMenu ]: Found $($objAllActivities.Count) Activities for this Runbook: '$($thisRootRunbookName)'"

				
					Log-Message "[ Debug-MainScript-ToCLeftMenu ]: Searching: $($thisRootRunbookFolderName) for files"
					# Start building our array of files/actvities for this object
					$arrFilesInDirectory += (Get-ChildItem "$($ReportExportPath)$($thisRootRunbookFolderName)\" | ? { $_.Name -ne "Index.html" } | Sort-Object $ToNatural)
					
					$strLeftMenuHTML += "<ul>"	
					
					if ($arrFilesInDirectory.Count -gt 0)
					{
						# For each file (activity / url link)
						for ($f=0; $f -le ($arrFilesInDirectory.Count - 1); $f++)
						{
							# Get the current File path
							$thisMenuActivityLink = $arrFilesInDirectory[$f];
		
							if($thisMenuActivityLink -eq $null) { 
								Log-Message "[ DEBUG TRAP WARNING]: thisMenuActivitylink is not returning data (is currently null)"
								Log-Message "[ DEBUG TRAP ]: thisMenuActivityLink.Name = $($thisMenuActivityLink.Name)"
								Log-Message "[ DEBUG TRAP ]: thisMenuActivityLink = $($thisMenuActivityLink)"
								Log-Message "[ DEBUG TRAP ]: thisMenuActivityLink.Name = $($thisMenuActivityLink.Name)"
								Log-Message "[ DEBUG TRAP ]: thisMenuActivityLink.FullName = $($thisMenuActivityLink.FullName)"
								Log-Message "[ DEBUG TRAP WARNING]: End Trap"
							} else {
						
								if ($f -eq ($arrFilesInDirectory.Count - 1))
								{
									$strLeftMenuHTML += "<li class='last'><a href='..\$($thisRootRunbookFolderName)\$($thisMenuActivityLink.Name)'><span>" + "$($thisMenuActivityLink.Name.Replace('.html',$null))" + "</span></a></li>";
								}
								else
								{
									$strLeftMenuHTML += "<li><a href='..\$($thisRootRunbookFolderName)\$($thisMenuActivityLink.Name)'><span>" + "$($thisMenuActivityLink.Name.Replace('.html',$null))" + "</span></a></li>";
								}
							
							}
						}

					}
					$strLeftMenuHTML += "</ul>"
					# The LI tag will end the current runbook (header) menu with it's activies (submenus) and let us move to the next runbook.	
					$strLeftMenuHTML += "</li>"
					Log-Message "[ Debug-MainScript-ToCLeftMenu ]: Processed Menu Item for: $($thisMenuActivityLink.Name)"
			}
		}
	}
	
	# This closes out the entire left side menu - so at this point we should be done.
	$strLeftMenuHTML += "</ul>"	
	Log-Message "[ Debug-MainScript-ToCLeftMenu ]: Finished processing left side menu"
	
	Log-Message "[ TESTING MENU ]: This is the HTML Built from the Generate-RunbookMenuSelections function:"
	Log-Message "$($strLeftMenuHTML)"
	Log-Message "[ TESTING MENU ]: End of testing logging"
	return $strLeftMenuHTML;
	
}

function Generate-RunbookCoords ($xPath, $yPath, $id)
{
	# my hack of a function here will use some static values to "offset" my images against the maps
	#
	if ($xPath -eq 2)
	{
		# Add 30 - x of 2 means far left, and move 30 to center the circle
		$xPath = [int]$xPath + 32
		$templateMapHTML = "`r<!-- Add 32px to X in order to offset circle highlight -->`r"
		
	}
	else
	{
		# So far, other than the first item (initialize data with an offset of 30, the rest
		# have worked by only incrementing them by 10
		$xPath = [int]$xPath + 35
		$templateMapHTML = "`r<!-- Add 35px to X in order to offset circle highlight -->`r"
		
	}
	
	#Y Coords = Y property + 100, and use the following markers:
	# 102 = Second Block from Top / 327 = 5th Block from Top
	$yPath = [int]$yPath + 30
	
	
	$templateMapHTML += "<area id=`"area$($id)`" href=`"$($id).html`" shape=`"circle`" coords=`"$($xPath),$($yPath),30`" alt=`"$($id)`" data-maphilight='{`"stroke`":false,`"fillColor`":`"ff0000`",`"fillOpacity`":0.6}' title=`"$($id)`">`n"
	
	return $templateMapHTML
	
}

function Generate-LinkCoords ($source, $target, $sourceFile, $targetFile)
{
	# ---------------------------------------------------------------------------
	# I hate math, so this is what you get.  First we get the DIFFERENCE
	# between the x and y coords of the source and target, respectivly.
	#
	# Once the differences are calculated, we will compare for the following
	# three conditions, which will produce separate coords depending on the
	# situation.
	# ---------------------------------------------------------------------------
	# Step 1: Calculate Differences
	# -
	[string]$templateMapHTML
	
	# Declare return coords
	[int]$xSrc = [int]$source.PositionX.'#text' + 32
	[int]$xTarget = [int]$target.PositionX.'#text' + 32
	[int]$ySrc = [int]$source.PositionY.'#text' + 35
	[int]$yTarget = [int]$target.PositionY.'#text' + 35
	
	if ($sourceFile -ne $null -and $targetFile -ne $null)
	{
		Log-Message "[Generate-LinkCoords]: Found file $($srcURL)...updating area map with an href tag to point to that file..."
		$templateMapHTML += "`n<area id=`"areaLinkSource`" href=`"$($sourceFile)`" shape=`"circle`" coords=`"$($xSrc),$($ySrc),30`" alt=`"Link Source`" data-maphilight='{`"stroke`":false,`"fillColor`":`"04B45F`",`"fillOpacity`":0.6}' title=`"Link Source`">"
		
		Log-Message "[Generate-LinkCoords]: Found file $($targetFile)...updating area map with an href tag to point to that file..."
		$templateMapHTML += "`n<area id=`"areaLinkTarget`" href=`"$($targetFile)`" shape=`"circle`" coords=`"$($xTarget),$($yTarget),30`" alt=`"Link Target`" data-maphilight='{`"stroke`":false,`"fillColor`":`"ff0000`",`"fillOpacity`":0.6}' title=`"Link Target`">"
	}
	else
	{
		#Log-Message "[Generate-LinkCoords]: No URL Passed - no url link for the coordinates of the source object"
		$templateMapHTML += "`n<area id=`"areaLinkSource`" shape=`"circle`" coords=`"$($xSrc),$($ySrc),30`" alt=`"Link Source`" data-maphilight='{`"stroke`":false,`"fillColor`":`"04B45F`",`"fillOpacity`":0.6}' title=`"Link Source`">"
		
		#Log-Message "[Generate-LinkCoords]: No URL Passed - no url link for the coordinates of the source object"
		$templateMapHTML += "`n<area id=`"areaLinkTarget`" shape=`"circle`" coords=`"$($xSrc),$($ySrc),30`" alt=`"Link Source`" data-maphilight='{`"stroke`":false,`"fillColor`":`"04B45F`",`"fillOpacity`":0.6}' title=`"Link Source`">"
			
	}

	return $templateMapHTML
}

#region Log-Message Function Info
####################################################################################
# *******************************************************************
# Function: Log-Message
# *******************************************************************
#	Author:			Michael Adams 
#					http://willcode4foodblog.wordpress.com
#	Purpose:
#					Take any string input and output it to a file path.
#					Nice things this does:
#						- Automatically Create a date stamped log file
#						- Automatically Timestamp each log entry
#						- Automatically output any error data at the time of the log
#
#	Paramaters:
#		-logMsg 	String to log to
#		-file		String of file path to log to
#					NOTE: You can use the default $logFile variable
#					when loading the GeneralFunctionLibary.ps1 script
#
#	Example Usage:
#
#		$myString = "This is a test";
#		Log-Message -logMsg $myString -file $logFile;
#
#	NOTE: Remember $logFile was already loaded when we loaded the function library,
#		  so you just have to provide the message to log and it will handle the rest!
#
# *******************************************************************
#endregion
Function Log-Message {
param (
	[Parameter(Mandatory = $true)]
	[string]$logMsg,
	[string]$file = "$($ScriptDirectory)\log\OISExport.log"
)

	# If a file doesnt get passed by accident, default to general log
	if ([String]::IsNullOrEmpty($file))
	{
		if (!(Test-Path "$($ScriptDirectory)\log"))
		{
			New-Item -ItemType Directory -Path "$($ScriptDirectory)\log" -Force -Confirm:$false
		}
	}

	# If the err variable is not empty
	if ($logMsg -ne $null)
	{
		# Log the message, and prepend a timestamp so we can refer to that later if needed
		$msgString = "[ $((Get-Date).DateTime) ]: $($logMsg)"
		
		$msgString | Out-File "$($file)" -Append -NoClobber;
		
		if ($Error -ne $null)
		{
			
			if ($Error[0].CategoryInfo.Activity -ne "Get-WMIObject")
			{
				
				# Check for additional Data (Only if there is data in the default error variable)
				if ($Error[0].InvocationInfo -ne $null)
				{
					$Error[0] | Select FullyQualifiedErrorId | FL | Out-File "$($file)" -Append -NoClobber;
					$errString = "[ $((Get-Date).DateTime) ] Additional Error Information:"
					$errString | Out-File "$($file)" -Append -NoClobber;
					$Error[0].InvocationInfo | Select ScriptLineNumber, Line, ScriptName | FL | Out-File "$($file)" -Append -NoClobber;
				}
				
			}
		}
		
		# Reset the error variable so we dont log the same crap over and over
		$Error.Clear();	
	}
	else
	{
		# Gotta do something, lets let the person know whats up
		$errMsg = 'No Message To Log!'
		Write-Error $errMsg
	}	
	
}
# End Function Log-Message

function Format-PoSHComments($testString)
{
	
	# Find all of the comments (by regexing a search for lines starting with a hashtag
	$results = $testString | Select-String '\B# (\w*[A-Za-z_]+\w*.*)' -AllMatches
	
	# ForEach String Found, Format it accordingly with my special CSS tags for PoshComments
	if ($results.Matches.Count -gt 0)
	{
		$results.Matches | % {
			
			# Get the current comment
			$Comment = $_.Value
			
			# Take the string and add our html/css class
			$thisComment = '<span class="PowerShellComment">' + $Comment + '</span>'
			
			# update the text with the new version of html
			$testString = $testString.Replace($Comment, $thisComment)
			
		}
	}
	return $testString;
}

function Format-SQLComments($testString)
{
	
	# Find all of the comments (by regexing a search for lines starting with a hashtag
	$results = $testString | Select-String '(--.*)|(((/\*)+?[\w\W]+?(\*/)+))' -AllMatches
	
	# ForEach String Found, Format it accordingly with my special CSS tags for PoshComments
	if ($results -ne $null -and $results.Matches.Count -gt 0)
	{
		$results.Matches | % {
			
			# Get the current comment
			$Comment = $_.Value
			
			# Take the string and add our html/css class
			$thisComment = '<span class="PowerShellComment">' + $Comment + '</span>'
			
			# update the text with the new version of html
			$testString = $testString.Replace($Comment, $thisComment)
			
		}
	}
	return $testString;
}

function Generate-RunbookActivityList($listActivities)
{
	
	# for each object
	$listActivities | % {
		
		$_.Name.'#text'
		
	}
	
}

function Convert-ToBase64Pic
{
	Param ([string]$path)
	
	# Set the variable to null - which will be returned if there are no results
	$result = $null

	# If the file specified actually exists
	if (Test-Path $path)
	{
		# convert the file to a base64 string
		$result = [convert]::ToBase64String((get-content $path -encoding byte))
	}
	
	$result
}

function Organize-RunbookFolders ($rootPath)
{
	# Check to see if the folder already exists (shouldn't unless it's a re-run)
	if (!(Test-Path $rootPath))
	{
		try
		{
			# Create the folder because it didn't exist and don't prompt for confirmation
			New-Item -ItemType Directory -Path $rootPath -Confirm:$false -Force
			return $true
			
		}
		Catch [Exception] {
			
			Write-Host "Error Creating Runbooks: $($Error[0])"
			return $false
		}
	}
}

function Build-TableOfContents($arrAllActivities, $ActivityTOCFile, $refImage, $runbook)
{
	# Import first half of ToC Template
	$htmlHeader = (Get-Content $TableOfContentsHeader) -join "`n"
	
	$thisRunbook = $runbook
	$SubRunbookName = $thisRunbook.Name.'#text'
	Write-Host "Preparing Report for Table of Contents for Runbook $SubRunbookName ..."
	
	# Get the javascript from our local source for the menu highlighting jquery
	# The join piece that enables the text to keep the line breaks was found at: http://stackOverflow.com/questions/15041857/powershell-keep-text-formatting-when-reading-in-a-file
	$jQueryCode = (Get-Content $JQueryMenuHighlightScript) -join "`n"
	
	# Convert our reference pic to a string so we can hardcode it into the report
	$thisReferenceImage = $PictureRepository + $SubRunbookName.Replace(' ', $null) + ".png"
	$refImage = Convert-ToBase64Pic -path $thisReferenceImage
	
	$htmlHeader = $htmlHeader.Replace('%_REFERENCE_RUNBOOK_IMAGE_%', $refImage)
	$htmlHeader = $htmlHeader.Replace('%_RUNBOOK_%', $SubRunbookName)
	
	# Embed the JQuery Code into the file.  This may make the file a tad bigger - but you do not get 
	# The security warnings our Business Web Browser Policies generate.
	# Additionally, no one should ever be trying to directly modify the report files, 
	# since the HTML was all generated from the script, the script, process, and templates
	# are what need to be reviewed for any modifications that would result in modified HTML output
	#-----------------------------------------------------------------------------------------------
	# The only reason you should directly modify a generated report is to validate code changes you 
	# wish to reflect across all subsequent reports generated from the script
	$htmlHeader = $htmlHeader.Replace('%_PLACEHOLDER_JQUERY_JAVASCRIPT_%',$jQueryCode)
	
	
	# Spit out the header HTML to the TOC file
	$htmlHeader | Out-File $ActivityTOCFile -Append -Force -Confirm:$false -Encoding UTF8
	
	
	# Initialize a counter to use in the loop below
	$menuCounter = 1
	$activityCount = $arrAllActivities.Count
	$activityCounter = 0
	
	$arrJQueryFunctions = @() # Empty array that we will build up in the loop below
	$arrJQueryAreaMaps = @() # Empty Array that will build up in the loop below
	
	# for each activity found Write the TOC File
	$arrAllActivities | select Name, PositionY, PositionX, UniqueID, ParentID, SourceObject, TargetObject, ObjectTypeName | % {
		
		$activityCounter++
		$thisActivity = $_
		
		# Declare empty variables for our interesting data (Link source and targets)
		$objSourceObject = $null
		$objTargetObject = $null		
		
		# Instead of retyping wierd code - lets assign the name to a variable
		$strActivityName = $thisActivity.Name.'#text'
		$strLinkID = $strActivityName.Replace(' ', $null)
		$strHTMLLinkText = "$($activityCounter)-$($strLinkID)"
		
		# If the X and Y coords for this activity are NOT null, then do the coords script
		if ($thisActivity.PositionY.'#text' -ne $null -and $thisActivity.PositionX.'#text' -ne $null -and $thisActivity.ObjectTypeName.'#text' -ne 'Link')
		{
			$intYPosition = $_.PositionY.'#text'
			$intXPosition = $_.PositionX.'#text'
			
			# Generate HTML for an area (coordinates for the red dot as an image map)
			$strJQueryAreaMapTemplate = Generate-RunbookCoords -xPath $intXPosition -yPath $intYPosition -id $strHTMLLinkText
		}

		
		# Add the map to the array so we can update this data in our HTML at the end
		$arrJQueryAreaMaps += $strJQueryAreaMapTemplate
		
	}
	
	# For each Function found - add it to the table of contents file
	# Update the File with new information - appending another placeholder at the end for the next var
	$tmpTOCFile = Get-Content $ActivityTOCFile
	
	Write-Host "Updating Area Maps...." -NoNewLine
	
	# For each block of code in the Area Maps array
	for ($i = 0; $i -lt $arrJQueryAreaMaps.Count; $i++)
	{
		
		# Current Map Text
		$thisAreaHTML = $arrJQueryAreaMaps[$i]
		
		# Temporary declare the variable we will return
		[string]$strResult = "NULL"
		
		# If this item is NOT the last one
		if ($i -ne ($arrJQueryAreaMaps.Count - 1))
		{
			
			# Build New String and append ANOTHER placeholder so the NEXT string can be processed
			$strResult = $thisAreaHTML + "`r`n" + '%_PLACEHOLDER_ACTIVITY_COORDS_%'
			
		}
		else
		{
			# Build New String and LEAVE OUT placeholder - this is the last item in the array
			$strResult = $thisAreaHTML
		}
		
		# Update the temporary variable for the report to replace the map placeholder with the code we just "built"
		$tmpTOCFile = $tmpTOCFile.Replace('%_PLACEHOLDER_ACTIVITY_COORDS_%', $strResult)
	}
	
	Write-Host "Outputting custom Table Of Contents file - just missing the footer..."
	$tmpTOCFile | Out-File $ActivityTOCFile -Force -Confirm:$false -Encoding UTF8
	
	Write-Host "Updating Footer HTML..." -NoNewLine
	
	# Import footer template code (probably doesn't need this - but too bad its the way I did it)
	$htmlFooter = Get-Content $TableOfContentsFooter
	
	# Spit out the footer HTML to the TOC file
	$htmlFooter | Out-File $ActivityTOCFile -Append -Force -Confirm:$false -Encoding UTF8

}

function Process-PowerShellReport ($object, $counter, $icon, $runbook)
{
	$x = $counter
	
	$thisRunbook = $runbook
	Write-Host "Preparing Report for .NET Script Activity: $($object.Name.'#text') ..."
	$SubRunbookName = $thisRunbook.Name.'#text'
	
	# Import the template contents to work with later
	$tmpHTMLReport = (Get-Content $templateReport) -join "`n"
	
	# Get the javascript from our local source for the menu highlighting jquery
	# The join piece that enables the text to keep the line breaks was found at: http://stackOverflow.com/questions/15041857/powershell-keep-text-formatting-when-reading-in-a-file
	$jQueryCode = (Get-Content $JQueryMenuHighlightScript) -join "`n"

	# Embed the JQuery Code into the file.  This may make the file a tad bigger - but you do not get 
	# The security warnings our Business Web Browser Policies generate.
	# Additionally, no one should ever be trying to directly modify the report files, 
	# since the HTML was all generated from the script, the script, process, and templates
	# are what need to be reviewed for any modifications that would result in modified HTML output
	#-----------------------------------------------------------------------------------------------
	# The only reason you should directly modify a generated report is to validate code changes you 
	# wish to reflect across all subsequent reports generated from the script
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_JQUERY_JAVASCRIPT_%',$jQueryCode)
		
	# Convert our reference pic to a string so we can hardcode it into the report
	$thisReferenceImage = $PictureRepository + $SubRunbookName.Replace(' ', $null) + ".png"
	$picRunbookReference = Convert-ToBase64Pic -path $thisReferenceImage
	
	
	# Build a new object so we can see whats up
	$objActivity = Generate-NewSCOrchObject -inputObject $object
	
	# Update the data in the temporary report text to reflect the actual instance we are working with
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_RUNBOOK_%', $SubRunbookName)
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_ACTIVITY_%', $objActivity.Name)
	
	# Remove Spaces so they are not in a file name (keeping length to a minimum)
	$activityFileName = $($objActivity.Name.Replace(' ', $null))
	
	# Build a file name with the current Runbook and Activity Name in it
	$PowerShellScriptsFile = $fileName.Replace(".html", "\$($x)-$($activityFileName).html")
	
	# Create txt files for which we will input Powershell and SQL Stuff Respectively
	New-Item -ItemType File -Path $PowerShellScriptsFile -Force -Confirm:$false
	
	# Import templates to build our report
	$tmpPoSHActivityHTML = (Get-Content $PoshInfoTemplate) -join "`n"
	
	# Add the company logo to the upper left corner
	$base64companyLogo = Convert-ToBase64Pic -path "$($script:strCompanyLogo)"
	$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_SRC_IMG_COMPANYLOGO_%', $base64companyLogo)
	
	# Now add (embed) the image of the individual activity
	$base64ActivityIcon = Convert-ToBase64Pic -path "$($icon)"
	$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_ICON_ACTIVITY_%', $base64ActivityIcon)
	# Update the date of this document being generated:
	$tmpPoSHActivityHTML = $tmpPosHActivityHTML.Replace('%_TODAYS_DATE_%', "$(Get-Date) by $([Environment]::UserName)")
	# And the coup de gras - lets add the reference image for storage in the file itself
	$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_PLACEHOLDER_RUNBOOK_IMAGE_%', $picRunbookReference)
	
	# Update the Activity Information by Cheating (Replace)
	$strrunBookTitle = $PowerShellScriptsFile.Replace($ReportExportPath, $null)
	$arrRunbookTitle = $strrunBookTitle.Split('-')
	$runBookTitle = $arrRunbookTitle[0] + "-" + $arrRunbookTitle[1]
	
	# Update report HTML with current variables
	$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_RUNBOOK_%', $runBookTitle)
	$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_ACTIVITY_%', $objActivity.Name.ToString())
	$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_UNIQUEID_%', $objActivity.UniqueID.ToString())
	# if there's no description we need to make sure to have something
	if ($object.Description.'#text' -ne $null)
	{
		$tmpPoSHActivityHTML = $tmpPoshActivityHTML.Replace('%_DESCRIPTION_%', "<b>" + $object.Description.'#text' + "</b>")
	}
	else
	{
		# Replace the description with some html code and make it red to ensure people see this needs to be updated
		$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_DESCRIPTION_%', "<span style=`"color:red;`">(empty / null)</span>")
	}
	
	# Script Properties
	$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_OBJECTTYPENAME_%', $objActivity.ObjectTypeName.ToString())
	
	if($objActivity.ScriptType -ne $null) {
		$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_SCRIPTTYPE_%', $objActivity.ScriptType)
	} else {
		$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_SCRIPTTYPE_%',"(null)")
	}
	# Get the script text and then update the report HTML appropriately
	$formattedScriptBody = $objActivity.ScriptBody
	$tmpPoSHActivityHTML = $tmpPoSHActivityHTML.Replace('%_SCRIPTBODY_%', $formattedScriptBody)
	
	# Update the Activity Information in your temporary report
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_ACTIVITY_INFORMATION_%', $tmpPoSHActivityHTML)
	
	if ($objActivity.PublishedData -ne $null)
	{
		$tmpPoSHPublishedDataHTML = $objActivity.PublishedData
		
		# Update and format the published data from this activity
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('<IsCollection>', '<td><b>')
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('</IsCollection>', '</b></td>')
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('<Entry>', '<tr>')
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('</Entry>', '</tr>')
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('<Name>', '<td>')
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('</Name>', '</td>')
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('<Type>', '<td><b>')
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('</Type>', '</b></td>')
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('<Variable>', '<td>')
		$tmpPoSHPublishedDataHTML = $tmpPoSHPublishedDataHTML.Replace('</Variable>', '</td>')
		# Update the temp report with the published data
		$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_PUBLISHED_DATA_%', $tmpPoSHPublishedDataHTML)
	}
	else
	{
		# This is the HTML to use if there is not any "Published Data" for the Activity
		$replacementHTML = @'
	<tr colspan=4>
		<td>No Published Data For This Activity</td>
	</tr>
'@
		$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_PUBLISHED_DATA_%', $replacementHTML)
	}
	
	# Before sending off the data, let's format the comments real quick
	$tmpHTMLReport = Format-PoSHComments -testString $tmpHTMLReport
	
	# Replace the formatting for subscription data
	$tmpHTMLReport = $tmpHTMLReport.Replace('\`d.T.~Ed/', $null)
	$tmpHTMLReport = $tmpHTMLReport.Replace('\`d.T.~Vb/', $null)
	
	# Instead of retyping wierd code - lets assign the name to a variable
	$strActivityName = $object.Name.'#text'
	$strLinkID = $strActivityName.Replace(' ', $null)
	$strHTMLLinkText = "$($counter)-$($strLinkID)"
	
	# Now get the coordinates of the activity icon and generate an "area" for our map to be highlighted for this page
	# The JQuery setting for this instance of the maphilight plugin is "alwaysOn:true" so we don't need to mouseover the image
	$intYPosition = $object.PositionY.'#text'
	$intXPosition = $object.PositionX.'#text'
	
	# Generate HTML for an area (coordinates for the red dot as an image map)
	$strJQueryAreaMapTemplate = Generate-RunbookCoords -xPath $intXPosition -yPath $intYPosition -id $strHTMLLinkText
	
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_ACTIVITY_COORDS_%', $strJQueryAreaMapTemplate)
	
	# Obtain the template we have for our CSS Menu that will show links (associated with this object/Activity) to the user
	$templateCSSMenu = Get-Content "$($CSSMenuTemplate)"
	
	# Now update the actual HTML template with our CSS Menu (that also was generated from a template)
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_CSS_LINK_MENU_%', $templateCSSMenu)
	
	
	# Last but not least: Update the CSS with our style sheet code (this will embed it in the doc so we don't need another location to store the css file)
	$stylesheetData = Get-Content $StyleSheetPath
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_LAYOUT_CSS_%', $stylesheetData)
	
	# Send our formatted/dolled up string to the file waiting for it, creating our report
	$tmpHTMLReport | Out-File -FilePath $PowerShellScriptsFile -Append -Encoding UTF8
}

function Process-DBQueryReport ($object, $counter, $icon,$runbook)
{
	$x = $counter
	$thisRunbook = $runbook
	Write-Host "Preparing Report for DB Query Activity: $($object.Name.'#text') ..."
	$SubRunbookName = $thisRunbook.Name.'#text'
	# Import the template contents to work with later
	$tmpHTMLReport = Get-Content $templateReport
	
	
	# Get the javascript from our local source for the menu highlighting jquery
	# The join piece that enables the text to keep the line breaks was found at: http://stackOverflow.com/questions/15041857/powershell-keep-text-formatting-when-reading-in-a-file
	$jQueryCode = (Get-Content $JQueryMenuHighlightScript) -join "`n"

	# Embed the JQuery Code into the file.  This may make the file a tad bigger - but you do not get 
	# The security warnings our Business Web Browser Policies generate.
	# Additionally, no one should ever be trying to directly modify the report files, 
	# since the HTML was all generated from the script, the script, process, and templates
	# are what need to be reviewed for any modifications that would result in modified HTML output
	#-----------------------------------------------------------------------------------------------
	# The only reason you should directly modify a generated report is to validate code changes you 
	# wish to reflect across all subsequent reports generated from the script
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_JQUERY_JAVASCRIPT_%',$jQueryCode)	
	
	# Convert our reference pic to a string so we can hardcode it into the report
	$thisReferenceImage = $PictureRepository + $SubRunbookName.Replace(' ', $null) + ".png"
	$picRunbookReference = Convert-ToBase64Pic -path $thisReferenceImage
	
	
	# Build a new object so we can see whats up
	$objActivity = Generate-NewSCOrchObject -inputObject $object
	
	# Update the data in the temporary report text to reflect the actual instance we are working with
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_RUNBOOK_%', $SubRunbookName)
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_ACTIVITY_%', $objActivity.Name)
	
	# Remove Spaces so they are not in a file name (keeping length to a minimum)
	$activityFileName = $($objActivity.Name.Replace(' ', $null))
	
	# Build a file name with the current Runbook and Activity Name in it
	$SQLQueryFile = $fileName.Replace(".html", "\$($x)-$($activityFileName).html")
	#$SQLScriptsFile = $ReportExportPath + $fileName.Replace(".html", "-SQLScripts.html")
	
	# Create txt files for which we will input Powershell and SQL Stuff Respectively
	New-Item -ItemType File -Path $SQLQueryFile -Force -Confirm:$false
	
	
	# Import templates to build our report
	$tmpSQLActivityHTML = (Get-Content $SQLInfoTemplate) -join "`n"
	
	# Add the company logo to the upper left corner
	$base64companyLogo = Convert-ToBase64Pic -path "$($script:strCompanyLogo)"
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_SRC_IMG_COMPANYLOGO_%', $base64companyLogo)
	
	# Now add (embed) the image of the individual activity
	$base64ActivityIcon = Convert-ToBase64Pic -path "$($icon)"
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_ICON_ACTIVITY_%', $base64ActivityIcon)
	
	# And the coup de gras - lets add the reference image for storage in the file itself
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_PLACEHOLDER_RUNBOOK_IMAGE_%', $picRunbookReference)
	
	# Update the Activity Information by Cheating (Replace)
	$strrunBookTitle = $SQLQueryFile.Replace($ReportExportPath, $null)
	$arrRunbookTitle = $strrunBookTitle.Split('-')
	$runBookTitle = $arrRunbookTitle[0] + "-" + $arrRunbookTitle[1]
	
	# Update report HTML with current variables
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_RUNBOOK_%', $runBookTitle)
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_ACTIVITY_%', $objActivity.Name.ToString())
	
	# if there's no description we need to make sure to have something
	
	# if there's no description we need to make sure to have something
	if ($object.Description.'#text' -ne $null)
	{
		$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_DESCRIPTION_%', "<b>" + $object.Description.'#text' + "</b>")
	}
	else
	{
		# Replace the description with some html code and make it red to ensure people see this needs to be updated
		$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_DESCRIPTION_%', "<span style=`"color:red;`">(empty / null)</span>")
	}
	
	# Script Properties
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_OBJECTTYPENAME_%', $objActivity.ObjectTypeName.ToString())
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_FLATTEN_%', $objActivity.Flatten.ToString())
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_FLAT_USE_LINE_BREAK_%', $objActivity.FlatUseLineBreak.ToString())
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_FLAT_USE_CUSTOM_SEP_%', $objActivity.FlatUseCustomSep.ToString())
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_UNIQUEID_%', $objActivity.UniqueID.ToString())
	
	if ($objActivity.FlatCustomSep -ne $null)
	{
		$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_FLAT_CUSTOM_SEP_%', $objActivity.FlatCustomSep.ToString())
	}
	else
	{
		$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_FLAT_CUSTOM_SEP_%', "N/A")
	}
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_SERVER_%', $objActivity.ServerName.ToString())
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_DATABASE_TYPE_%', $objActivity.DatabaseType.ToString())
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_INITIAL_CATALOG_%', $objActivity.InitialCatalog.ToString())
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_DB_AUTHENTICATION_%', $objActivity.DatabaseAuthentication.ToString())
	
	if ($objActivity.UserName -ne $null)
	{
		# If the value is not null, use the value
		$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_USERNAME_%', $objActivity.UserName.ToString())
	}
	else
	{
		# Otherwise let them know its being run under Windows Auth (Svc Account or current user)
		$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_USERNAME_%', "Runbook Service Account Credentials")
	}
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_CONNECTION_TIMEOUT_%', $objActivity.ConnectionTimeout.ToString())
	
	# Update the date of this document being generated:
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_TODAYS_DATE_%', "$(Get-Date) by $([Environment]::UserName)")
	
	$objActivity.Query = Format-SQLComments -testString $objActivity.Query
	
	# Get the script text and then update the report HTML appropriately
	$tmpSQLActivityHTML = $tmpSQLActivityHTML.Replace('%_QUERY_%', $objActivity.Query)
	
	# Format the comments from the SQL Query
	
	# Update the Activity Information in your temporary report
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_ACTIVITY_INFORMATION_%', $tmpSQLActivityHTML)
	
	
	$replacementHTML = @'
	<tr colspan=4>
		<td>No Published Data For This Activity - Database queries return row(s) data</td>
	</tr>
'@
	
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_PUBLISHED_DATA_%', $replacementHTML)
	
	# Before sending off the data, Clean up the data a bit
	# --------------------------------------------------------------------------------------------
	# Replace the formatting for subscription data
	$tmpHTMLReport = $tmpHTMLReport.Replace('\`d.T.~Ed/', $null)
	$tmpHTMLReport = $tmpHTMLReport.Replace('\`d.T.~Vb/', $null)
	$tmpHTMLReport = $tmpHTMLReport.Replace('\`d.T.~Br/', $null)
	
	# Instead of retyping wierd code - lets assign the name to a variable
	$strActivityName = $object.Name.'#text'
	$strLinkID = $strActivityName.Replace(' ', $null)
	$strHTMLLinkText = "$($counter)-$($strLinkID)"
	
	# Now get the coordinates of the activity icon and generate an "area" for our map to be highlighted for this page
	# The JQuery setting for this instance of the maphilight plugin is "alwaysOn:true" so we don't need to mouseover the image
	$intYPosition = $object.PositionY.'#text'
	$intXPosition = $object.PositionX.'#text'
	
	# Generate HTML for an area (coordinates for the red dot as an image map)
	$strJQueryAreaMapTemplate = Generate-RunbookCoords -xPath $intXPosition -yPath $intYPosition -id $strHTMLLinkText
	
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_ACTIVITY_COORDS_%', $strJQueryAreaMapTemplate)
	
	# Last but not least: Update the CSS with our style sheet code (this will embed it in the doc so we don't need another location to store the css file)
	$stylesheetData = Get-Content $StyleSheetPath
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_LAYOUT_CSS', $stylesheetData)
	
	# Obtain the template we have for our CSS Menu that will show links (associated with this object/Activity) to the user
	$templateCSSMenu = Get-Content "$($CSSMenuTemplate)"
	
	# Now update the actual HTML template with our CSS Menu (that also was generated from a template)
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_CSS_LINK_MENU_%', $templateCSSMenu)
	
	# Send our formatted/dolled up string to the file waiting for it, creating our report
	$tmpHTMLReport | Out-File -FilePath $SQLQueryFile -Append -Encoding UTF8
	
}

function Generate-NewSCOrchObject ($inputObject, $runbook)
{
	$objActivity = $null
	$results = $null
	$thisRunbook = $runbook
	# Build a new object so we can see whats up
	# This also acts as a "filter" -as tons of unnecessary
	$objActivity = New-Object -TypeName PSObject
	
	# Mostly debugging text to see where we are in the process and that each piece of data came across right
	Write-Host "Object is of type: $($inputObject.ObjectTypeName.'#text')...adding interesting data" -ForegroundColor White
	Write-Host ""
	
	# Find the object type name and either throw it through the default process of spitting out every property
	# or add additional conditions for activities where you want to "target" only certain properties for brevity's sake
	switch ($inputObject.ObjectTypeName.'#text')
	{
		'Run .Net Script' {
			Write-Host "Object is of type: $($inputObject.ObjectTypeName.'#text')...adding interesting data" -ForegroundColor White
			Write-Host ""
			# Add the interesting properties for a .NET Script
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Name -Value $inputObject.Name.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Description -Value $inputObject.Description.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name ObjectTypeName -Value $inputObject.ObjectTypeName.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name ScriptType -Value $inputObject.ScriptType.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name ScriptBody -Value $inputObject.ScriptBody.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name PublishedData -Value $inputObject.PublishedData.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name UniqueID -Value $inputObject.UniqueID.'#text'
			Write-Host "Interesting Data for objecttype: $($inputObject.ObjectTypeName.'#text') has been added" -ForegroundColor White
			Write-Host ""
		}
		'Query Database' {
			# Add the interesting properties for a Query Database Activity
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Name -Value $inputObject.Name.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Description -Value $inputObject.Description.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name ObjectTypeName -Value $inputObject.ObjectTypeName.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Flatten -Value $inputObject.Flatten.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name FlatUseLineBreak -Value $inputObject.FlatUseLineBreak.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name FlatUseCustomSep -Value $inputObject.FlatUseCustomSep.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name FlatCustomSep -Value $inputObject.FlatCustomSep.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Query -Value $inputObject.Query.'#text'.ToString()
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name ServerName -Value $inputObject.ServerName.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name DatabaseType -Value $inputObject.DatabaseType.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name InitialCatalog -Value $inputObject.InitialCatalog.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name DatabaseAuthentication -Value $inputObject.DatabaseAuthentication.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name UserName -Value $inputObject.UserName.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name ConnectionTimeout -Value $inputObject.ConnectionTimeout.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name UniqueID -Value $inputObject.UniqueID.'#text'
			Write-Host "Interesting Data for objecttype: $($inputObject.ObjectTypeName.'#text') has been added" -ForegroundColor White
			Write-Host ""
		}
		'Link' {
			# Add the interesting properties for a Link Activity
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Name -Value $inputObject.Name.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Description -Value $inputObject.Description.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name ObjectTypeName -Value $inputObject.ObjectTypeName.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Enabled -Value $inputObject.Enabled.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name CreationTime -Value $inputObject.CreationTime.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name CreatedBy -Value $inputObject.CreatedBy.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name LastModified -Value $inputObject.LastModified.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name LastModifiedBy -Value $inputObject.LastModifiedBy.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Flatten -Value $inputObject.Flatten.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name FlatUseLineBreak -Value $inputObject.FlatUseLineBreak.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name FlatUseCustomSep -Value $inputObject.FlatUseCustomSep.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name FlatCustomSep -Value $inputObject.FlatCustomSep.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Color -Value $inputObject.Color.'#text'
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name Triggers -Value $inputObject.TRIGGERS
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name SourceObject -Value (Get-ObjectByID -objectID $inputObject.SourceObject.'#text' -runbookObject $thisRunbook)
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name TargetObject -Value (Get-ObjectByID -objectID $inputObject.TargetObject.'#text' -runbookObject $thisRunbook)
			Add-Member -InputObject $objActivity -memberType NoteProperty -Name UniqueID -Value $inputObject.UniqueID.'#text'
		}
		default
		{
			# TODO: FINISH CONVERSION OF THE POWERSHELL REPORT FUNCTION
			# Do it smart - exclude 3 items from your object created earlier and loop through them to build each row:
			$arrPropertyList = $inputObject.psObject.Properties | ? { $_.MemberType -eq "Property" }
			
			# For each item in the array we just made
			$arrPropertyList | % {
				
				# Declare item as temp variable
				$thisProperty = $_
				# As long as there is data in the value field
				if ($thisProperty.Value -ne $null -and $thisProperty.Value -ne 'CUSTOM_START_PARAMETERS' -and $thisProperty.Value -ne 'PUBLISH_POLICY_DATA')
				{
					# As long as there is data in the value field
					$propValue = $thisProperty.Value
					$propName = $thisProperty.Name
					#Log-Message -logMsg "Add property '$($propName)' with a value of '$propValue' to the object $($inputObject.Name.'#text')..."
					
					# Add property $propName with a value of $propValue to the object were creating
					Add-Member -InputObject $objActivity -memberType NoteProperty -Name $propName -Value $propValue
				}
				
				# Special Considerations for an 'Initialize Data' Activity's CUSTOM_START_PARAMETERS property.
				if($thisProperty.Value -eq 'CUSTOM_START_PARAMETERS') {
				
					$cStartParams = Get-CustomStartParameters -ActivityID $inputObject.UniqueID.'#text'
					[string]$propArray
					
					Log-Message "[ New-SCORCHObject ]:[CUSTOM_START_PARAMETERS]: This initialize data activity has: $($cStartParams.Count) starting parameters associated with it..."
					
					$cStartParams | % {
						$propArray += "$($_)";
						Log-Message "[ New-SCORCHObject ]:[CUSTOM_START_PARAMETERS]: Adding $($_) to the return string"
					}							
					
					# Add property $propName with a value of $propValue to the object were creating
					Add-Member -InputObject $objActivity -memberType NoteProperty -Name "CUSTOM_START_PARAMETERS" -Value $propArray -Force	
					$propArray = $null;
				}
				
				# Special Considerations for an 'Initialize Data' Activity's CUSTOM_START_PARAMETERS property.
				if($thisProperty.Value -eq 'PUBLISH_POLICY_DATA') {
				
					$thisPolicyData = Get-PublishedPolicyData -ActivityID $inputObject.UniqueID.'#text'
					[string]$propArray
					
					Log-Message "[ New-SCORCHObject ]:[PUBLISH_POLICY_DATA]: This initialize data activity has: $($thisPolicyData) starting parameters associated with it..."
					
					$thisPolicyData | % {
						$propArray += "$($_)";
						Log-Message "[ New-SCORCHObject ]:[PUBLISH_POLICY_DATA]: Adding $($_) to the return string"
					}							
					
					# Add property $propName with a value of $propValue to the object were creating
					Add-Member -InputObject $objActivity -memberType NoteProperty -Name "PUBLISH_POLICY_DATA" -Value $propArray -Force	
					$propArray = $null;
				}				
				
			}
		}
	}
	
	# Assign our newly built object to the return variable
	$results = $objActivity;
	
	# Return the new object to the script
	Return $results
	
}

function Get-ObjectByID ($objectID, $runbookObject)
{
	# Build our return variable but make it empty
	$returnObject = $null
	
	# First make sure there's SOMETHING to try
	if ($objectID -ne $null -and $runbookObject -ne $null)
	{
		# Find the object (activity/link) in this runbook by it's ID and return that object as the result
		$returnObject = $runbookObject.Object | ? { $_.UniqueID.'#text' -eq $objectID }
		
		# If there was nothing there, lets try each of our global variable 'folders' to see if the object is actually there instead
		if($returnObject -eq $null) {
			
			#Log-Message -logMsg "[Get-ObjectByID]: Did not locate $($objectID) in the Runbook folder $($runbookObject.Name.'#text')"
			
			# For every folder in our global configuration
			for($g=0;$g -le ($objGlobalRunbookVarFolders.Folder.Count - 1);$g++) {
				
				# set temp variable to current "Folder" in the global variables cache
				$thisGlobalVarFolder = $objGlobalRunbookVarFolders.Folder[$g]
				
				# Now Look through all the objects in this folder
				$returnObject = $thisGlobalVarFolder.Folder.Objects.Object | ? { $_.UniqueID.'#text' -eq $objectID -or $_.Entry.UniqueID -eq $objectID}
				
				# If we've found the object then jump out of the loop
				if($returnObject -ne $null) {
					Log-Message -logMsg "[Get-ObjectByID]: Found object (in GlobalVariables): $($returnObject.Name.'#text') with id of $($objectID)"					
					break;					
				} else {
					#Log-Message -logMsg "[Get-ObjectByID]: Did not locate $($objectID) in the Global Settings folder $($objGlobalRunbookVarFolders.Folder[$g].Name)"
					
				}
				
			}

		} else {
		
			Log-Message -logMsg "[Get-ObjectByID]: Found object (in Policy.Object): $($returnObject.Name.'#text') with id of $($objectID)"		
			
		}

	}
	
	# Pass the result back to the calling scriptblock
	return $returnObject
}

################################################################################################
# By keeping this function updated when you want to incorporate new icons (see documentation),
# this script will help you find the right icon to use for each "Activity"
function Get-IconForActivity ($object, $icons)
{
	# Declare return var
	$iconBase64 = $null
	
	# If the object isn't null - process it
	if ($object -ne $null)
	{	
		switch ($object.ObjectTypeName.'#text')
		{
			# The icons array contains file objects, so we need the 'fullname' property to get the files path
			# Generic Activities available by default in Orchestrator
			#--------------------------------------------------------
			'Publish Policy Data' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[4].FullName)"; break; } # Return Data Usually
			'Custom Start' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[3].FullName)"; break; } # Usually named 'Initialize Data'
			'Run .Net Script'{ $iconBase64 = Convert-ToBase64Pic -path "$($icons[1].FullName)"; break; }
			'Query Database' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[0].FullName)"; break; }
			'Compare Values' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[6].FullName)"; break; }
			'Send Event Log Message' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[2].FullName)"; break; }
			'Create Folder' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[8].FullName)"; break; }
			'Junction' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[11].FullName)"; break; }
			'Get File Status' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[12].FullName)"; break; }
			'Append Line' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[13].FullName)"; break; }
			'Trigger Policy' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[14].FullName)"; break; }
			
			# Integration Pack Icons: Active Directory
			# NOTE: Having these here is ok even if your runbooks don't use them
			#--------------------------------------------------------
			'Create User' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[7].FullName)"; break; }
			'Enable User' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[10].FullName)"; break; }
			
			# Integration Pack Icons: Exchange
			# NOTE: Having these here is ok even if your runbooks don't use them
			#--------------------------------------------------------
			'Enable Mailbox' { $iconBase64 = Convert-ToBase64Pic -path "$($icons[9].FullName)"; break; }
			
			# For everything that we haven't "snipped" out into icon files yet, we can give it a default one
			#--------------------------------------------------------
			default { $iconBase64 = Convert-ToBase64Pic -path "$($icons[5].FullName)"; break; }
		}
		
		
	}
	
	return $iconBase64;
}

################################################################################################
# Get all the custom start parameters and make them human-readable
# Full-Disclosure: This is a modified version of the method used by the SMART Documentation Toolkit
# that achieves the same result
function Get-CustomStartParameters ($activityID){
	
	$thisActivityID = $activityID.ToString().Replace('{',$null)
	$thisActivityID = $thisActivityID.ToString().Replace('}',$null)

	$returnParams = @();
	
	try {	
		
		$myConnection = New-Object System.Data.SqlClient.SqlConnection;
		$myConnection.ConnectionString = $SQLConnectionString
		
		# Open a session to the SQL Server (Make sure SQL Port is accessible from machine your connecting from)
		$myConnection.Open();
			

		Log-Message "[ Get-CustomStartupParameters]: There were custom startup parameters added in the workflow definition..." 
		
		
		# Query also taken from the SMART Doc util and integrated into this function.  
		$SqlQuery = "select value, type from CUSTOM_START_PARAMETERS where ParentID = '" + $thisActivityID + "'"
		$myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
		
		$dr = $myCommand.ExecuteReader()
		while ($dr.Read())
				{
					$thisStartParameter = ("Variable:`t" + "<b>" + "[" + $dr["type"] + "]" + $dr["value"] + "</b>" + "[ Accessed as: $" + $dr["value"].Replace(" ", "_") + " ]`n")
					$returnParams += $thisStartParameter.ToString()
				}
		$dr.Close()

		$myConnection.Close();
		return $returnParams
		
	} catch {
		# Throw an error that we can use to debug if necessary
		Show-MsgBox -Title "Error Querying SCOrch Database" -Prompt "An error occurred while querying the SCOrch Database.`r`nThe error was:`r'$($Error[0])" -Icon 'Critical' -DefaultButton '1'
		return $null;
	}
}


function Get-PublishedPolicyData ($activityID) {
	
	$thisActivityID = $activityID.ToString().Replace('{',$null)
	$thisActivityID = $thisActivityID.ToString().Replace('}',$null)
	
	$returnData= @();

	try {	
		# Build the connection to the Database 
		$myConnection = New-Object System.Data.SqlClient.SqlConnection;
		$myConnection.ConnectionString = $SQLConnectionString
		
		# Open a session to the SQL Server (Make sure SQL Port is accessible from machine your connecting from)
		$myConnection.Open();


		$SqlQuery = "select [Key],Value from PUBLISH_POLICY_DATA where ParentID = '" + $thisActivityID + "'"
		$myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
		$dr = $myCommand.ExecuteReader()
		while ($dr.Read())
				{
				$TmpProperty = $dr["value"]
				$returnData += ("[ Returned Data ]: <b>" + $dr["key"] + "</b> = $($TmpProperty)" + "`n")
				}
		$dr.Close()
		$myConnection.Close()
		
		return $returnData

	} catch {
		# Throw an error that we can use to debug if necessary
		Show-MsgBox -Title "Error Querying SCOrch Database" -Prompt "An error occurred while querying the SCOrch Database.`r`nThe error was:`r'$($Error[0])" -Icon 'Critical' -DefaultButton '1'
		return $null;
	}
}

################################################################################################
# Yes this is ugly but it works for me - I go back through the HTML files generated in a last
# pass to convert all GUIDs (Most, at least) to something understandable by us mere mortals.
# This function helps us find the Orchestrator object by pulling an ID from the HTML
function Find-ObjectInHTMLFiles($objectID, $rootDirectory)
{
	$foundFile = $null;
	$htmlFile = $null;
	
	if (Test-Path $rootDirectory)
	{
		$htmlFiles = Get-ChildItem "$rootDirectory" | ? { $_.Name -like "*.html" -and $_.Name -ne "Index.html" }
		
		$htmlFiles | % {
			
			$htmlFile = $_
			$foundFile = Get-Content "$($htmlFile.FullName)" | Select-String -Pattern '^.*<td class="alt">UniqueID<\/td>.*</tr>.*$'
			
			# If a match was found (should find something in every activity file)
			if ($foundFile -ne $null -and $foundFile.Line -ne $null)
			{
				# Clean up the line to leave only the ID
				$uniqueID = $foundFile.Line -replace '^.*<td class="alt">UniqueID</td><td>', $null
				$uniqueID = $uniqueID -replace '</td></tr>.*$', $null
				
				# If that ID matches the ID we are looking for
				if ($uniqueID -eq $objectID)
				{
					# return the name of the file to the calling script
					#Log-Message "HTML File found with ID of $($objectID).  The file name is $($htmlFile.Name)"
					return $htmlFile.Name
				}
				
			}
			
		}
		
		return $null;
		
	}
	else
	{
		Log-Message "[Find-ObjectInHTMLFiles]:ERROR: The directory '$($rootDirectory)' could not be found!"
		return $null;
	}
	
}

################################################################################################
# This is a specialized version of the "Process-ActivityReport" function tailored to Link Activities
function Process-LinkActivityReport ($object, $counter, $sourceID, $targetID, $runbook)
{
	$x = $counter
	$thisRunbook = $runbook

	Write-Host "Preparing Report for Link Activity: $($object.Name.'#text') ..."
	$SubRunbookName = $thisRunbook.Name.'#text'
	
	# Import the template contents to work with later
	$tmpHTMLReport = (Get-Content $templateReport) -join "`n"
	
	# Get the javascript from our local source for the menu highlighting jquery
	# The join piece that enables the text to keep the line breaks was found at: http://stackOverflow.com/questions/15041857/powershell-keep-text-formatting-when-reading-in-a-file
	$jQueryCode = (Get-Content $JQueryMenuHighlightScript) -join "`n"

	# Embed the JQuery Code into the file.  This may make the file a tad bigger - but you do not get 
	# The security warnings our Business Web Browser Policies generate.
	# Additionally, no one should ever be trying to directly modify the report files, 
	# since the HTML was all generated from the script, the script, process, and templates
	# are what need to be reviewed for any modifications that would result in modified HTML output
	#-----------------------------------------------------------------------------------------------
	# The only reason you should directly modify a generated report is to validate code changes you 
	# wish to reflect across all subsequent reports generated from the script
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_JQUERY_JAVASCRIPT_%',$jQueryCode)
	
	# Convert our reference pic to a string so we can hardcode it into the report
	$thisReferenceImage = $PictureRepository + $SubRunbookName.Replace(' ', $null) + ".png"
	$picRunbookReference = Convert-ToBase64Pic -path $thisReferenceImage
	$base64LinkActivityImage = Convert-ToBase64Pic -path "$($IconRootFolder)\image_link_activity.png"
	
	# Build a new object so we can see whats up
	# This also acts as a "filter" -as tons of unnecessary
	$objActivity = Generate-NewSCOrchObject -inputObject $object -runbook $thisRunbook
	
	# Update the data in the temporary report text to reflect the actual instance we are working with
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_RUNBOOK_%', $SubRunbookName)
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_ACTIVITY_%', $objActivity.Name)
	
	# Remove Spaces so they are not in a file name (keeping length to a minimum)
	$activityFileName = $($objActivity.Name.Replace(' ', $null))
	
	# Build a file name with the current Runbook and Activity Name in it
	$LinkActivityFile = $filename.Replace(".html", "\$($x)-$($activityFileName).html")
	# Create txt files for which we will input Powershell and SQL Stuff Respectively
	New-Item -ItemType File -Path $LinkActivityFile -Force -Confirm:$false
	
	# Import templates to build our report
	$tmpLinkActivityHTML = (Get-Content $LinkActivityTemplate) -join "`n"
	
	# Add the company logo to the upper left corner
	$base64companyLogo = Convert-ToBase64Pic -path "$($script:strCompanyLogo)"
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_SRC_IMG_COMPANYLOGO_%', $base64companyLogo)
	
	# Source Icon:  Find out which we need
	# available icons: db_query[0], dotnet_script[1], event_log[2], initialize[3], return_data[4], unknown_activity[5], compare_values[6], send_eventlog_message[7], publish_policydata[8]
	$base64SourceActivityIcon = Get-IconForActivity -object $objActivity.SourceObject -icons $arrIcons
	$base64TargetActivityIcon = Get-IconForActivity -object $objActivity.TargetObject -icons $arrIcons
	
	# Get the names of these source/targets so we can update the html appropriately in the template
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_SOURCEACTIVITYNAME_%', $objActivity.SourceObject.Name.'#text')
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_TARGETACTIVITYNAME_%', $objActivity.TargetObject.Name.'#text')
	
	# We have a little image to depict links and failures, lets update the html with the base64-coded image
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_LINK_ACTIVITY_IMAGE_%', $base64LinkActivityImage)
	
	# Now add (embed) the image of the source activity
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_SOURCEICON_%', "$($base64SourceActivityIcon)")
	
	# and do the same for the target activity
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_TARGETICON_%', "$($base64TargetActivityIcon)")
	
	
	# Update the date of this document being generated:
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_TODAYS_DATE_%', "$(Get-Date) by $([Environment]::UserName)")
	
	# Update the Activity Information by Cheating (Replace)
	$strrunBookTitle = $LinkActivityFile.Replace($ReportExportPath, ' ')
	$arrRunbookTitle = $strrunBookTitle.Split('-')
	$runBookTitle = $arrRunbookTitle[0] + "-" + $arrRunbookTitle[1]
	
	# Update report HTML with current variables
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_RUNBOOK_%', $runBookTitle)
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_ACTIVITY_%', $objActivity.Name.ToString())
	
	# if there's no description we need to make sure to have something
	if ($objActivity.Description -ne $null)
	{
		$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_DESCRIPTION_%', "<b>$($objActivity.Description)</b>")
	}
	else
	{
		# Replace the description with some html code and make it red to ensure people see this needs to be updated
		$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_DESCRIPTION_%', "<span style=`"color:red;`">(empty / null)</span>")
	}
	
	# Update Object Type Name
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_OBJECTTYPENAME_%', $objActivity.ObjectTypeName.ToString())
	# And the coup de gras - lets add the reference image for storage in the file itself
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_PLACEHOLDER_RUNBOOK_IMAGE_%', $picRunbookReference)
	#$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_UNIQUEID_%', $objActivity.UniqueID.ToString())
	
	# TODO: FINISH CONVERSION OF THE POWERSHELL REPORT FUNCTION
	# Do it smart - exclude 3 items from your object created earlier and loop through them to build each row:
	$arrPropertyList = @();
	$arrPropertyList += $objActivity.psObject.Properties | ? { $_.MemberType -eq "NoteProperty" -and $_.Name -ne "ActivityName" -and $_.Name -ne "Description"}
	
	# Get a counter for our loop
	[string]$rowHTML
	
	# For each item in the array we just made
	for ($y = 0; $y -le ($arrPropertyList.Count - 1); $y++)
	{	
		# Declare item as temp variable
		$thisProperty = $arrPropertyList[$y]
		#Log-Message -logMsg "[Process-LinkActivityReport]: Found Property: $($thisProperty.Name) with a value of:"
		#Log-Message -logMsg "[Process-LinkActivityReport]: $($thisProperty.Value)"
		
		if ($thisProperty.Value -eq $null) # As long as there is data in the value field
		{
			#Log-Message -logMsg "[Process-LinkActivityReport]: Processing as 'Empty/Null'..."
			# If the property value is empty, update the HTML to reflect it so it shows SOMETHING
			# We will hide or show it depending on the configuration item [General]=>[ShowNullorEmptyProperties]
			
			if($ShowNullorEmptyProps -eq $true) {
				$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">{1}</td></tr>`n" -f $thisProperty.Name, "(null/empty)"
			} 
		}
		else
		{
		
			if ($thisProperty.Name -eq "Triggers") {
			
				# for each item in the Triggers list, let's create some HTML for it
				$rowHTML += "<tr><td class=`"alt`">TRIGGERS</td><td colspan=`"3`">"

				# This is the ripped off function that does a great job of roughly "translating" the trigger conditions:
				$thisLinkCondition = LinkCondition -LinkID $objActivity.UniqueID -SQLConnectionString "$($script:SQLConnectionString)"
				
				if($thisLinkCondition -ne $null)
				{					
					$rowHTML +=  "$($thisLinkCondition)"
				
				}

				$rowHTML += "</td></tr>`n"
	
			}
			elseif($thisProperty.Name -eq "SourceObject")
			{
				#Log-Message -logMsg "[Process-LinkActivityReport]: Processing as 'SourceObject'..."
				$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">{1}</td></tr>`n" -f $thisProperty.Name, $objActivity.SourceObject.UniqueID.'#text'
			}
			elseif ($thisProperty.Name -eq "TargetObject")
			{
				#Log-Message -logMsg "[Process-LinkActivityReport]: Processing as 'TargetObject'..."
				$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">{1}</td></tr>`n" -f $thisProperty.Name, $objActivity.TargetObject.UniqueID.'#text'
			}
			else
			{
				$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">{1}</td></tr>`n" -f $thisProperty.Name, $thisProperty.Value;		
			}
		}
	}
	
	# Now that we have our loop done, update the html we just created to insert the new rows that contain property attributes
	$tmpLinkActivityHTML = $tmpLinkActivityHTML.Replace('%_PLACEHOLDER_LINKPROP_INFORMATION_%', $rowHTML)
	
	# Do our update of what HTML weve built to the temporary report template html:
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_ACTIVITY_INFORMATION_%', $tmpLinkActivityHTML)
	
	# This is the HTML to use if there is not any "Published Data" for the Activity
	$replacementHTML = @"
	<tr colspan=4>
		<td>No Published Data For This Activity</td>
	</tr>
"@
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_PUBLISHED_DATA_%', $replacementHTML)
	
	# Before sending off the data, Clean up the data a bit
	# --------------------------------------------------------------------------------------------
	# Replace the formatting for subscription data
	$tmpHTMLReport = $tmpHTMLReport.Replace('\`d.T.~Ed/', $null)
	$tmpHTMLReport = $tmpHTMLReport.Replace('\`d.T.~Vb/', $null)
	
	# Last but not least: Update the CSS with our style sheet code (this will embed it in the doc so we don't need another location to store the css file)
	$stylesheetData = Get-Content $StyleSheetPath
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_LAYOUT_CSS', $stylesheetData)
	
	# Send our formatted/dolled up string to the file waiting for it, creating our report
	$tmpHTMLReport | Out-File -FilePath $LinkActivityFile -Append -Encoding UTF8
	
}

################################################################################################
# This is the "master" function that is the generic "Activity" report.  
# Edit the Switch Cases in this function to specially handle other types of activities
function Process-ActivityReport ($object, $counter, $icons, $runbook)
{
	
	$thisRunbook = $runbook
	Write-Host "Preparing Report for General Activity: $($object.Name.'#text') ..."
	$SubRunbookName = $thisRunbook.Name.'#text'
	
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Preparing Report for General Activity: $($object.Name.'#text') ..."
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Runbook for this activity is: $($SubRunbookName) ..."
	
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Obtaining template for report from: $($templateReport)"
	# Import the template contents to work with later
	$tmpHTMLReport = (Get-Content $templateReport) -join "`n"
	
	# Get the javascript from our local source for the menu highlighting jquery
	# The join piece that enables the text to keep the line breaks was found at: http://stackOverflow.com/questions/15041857/powershell-keep-text-formatting-when-reading-in-a-file
	$jQueryCode = (Get-Content $JQueryMenuHighlightScript) -join "`n"

	# Embed the JQuery Code into the file.  This may make the file a tad bigger - but you do not get 
	# The security warnings our Business Web Browser Policies generate.
	# Additionally, no one should ever be trying to directly modify the report files, 
	# since the HTML was all generated from the script, the script, process, and templates
	# are what need to be reviewed for any modifications that would result in modified HTML output
	#-----------------------------------------------------------------------------------------------
	# The only reason you should directly modify a generated report is to validate code changes you 
	# wish to reflect across all subsequent reports generated from the script
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_JQUERY_JAVASCRIPT_%',$jQueryCode)
		
	# Convert our reference pic to a string so we can hardcode it into the report
	$thisReferenceImage = $PictureRepository + $SubRunbookName.Replace(' ', $null) + ".png"
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Converting Image: $($thisReferenceImage) to a base64 string..."
	$picRunbookReference = Convert-ToBase64Pic -path $thisReferenceImage
	
	# Build a new object so we can see whats up
	# This also acts as a "filter" -as tons of unnecessary
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Generating a PowerShell Object from object: $($object.Name.'#text')..."
	$objActivity = Generate-NewSCOrchObject -inputObject $object
	
	
	# Get the base64 string of the correct icon by passing the object and our array icon to this function
	$iconForActivity = Get-IconForActivity -object $object -icons $arrIcons
	#Log-Message -logMsg "[ GENERAL ACTIVITY ]: Icon selected for Activity: $($iconForActivity)"
	
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Replacing Template %_% variables in template with newly obtained variables"
	# Update the data in the temporary report text to reflect the actual instance we are working with
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_RUNBOOK_%', $SubRunbookName)
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Runbook: $($SubRunbookName)"
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_ACTIVITY_%', $objActivity.Name.'#text')
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Activity: $($objActivity.Name.'#text')"
	# Remove Spaces so they are not in a file name (keeping length to a minimum)
	$activityFileName = $object.Name.'#text'.ToString().Replace(' ', $null)
	
	# Build a file name with the current Runbook and Activity Name in it
	$GeneralActivityFile = $filename.Replace(".html", "\$($counter)-$($activityFileName).html")
	
	# Create txt files for which we will input Powershell and SQL Stuff Respectively
	New-Item -ItemType File -Path $GeneralActivityFile -Force -Confirm:$false
	
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Activity file: $($GeneralActivityFile) Created..."
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Updating template %_% variables with newly obtained ones..."
	
	# Import templates to build our report
	$tmpGeneralActivityHTML = (Get-Content $GeneralActivityTemplate) -join "`n"
	

	# Add the company logo to the upper left corner
	$base64companyLogo = Convert-ToBase64Pic -path "$($script:strCompanyLogo)"
	$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_SRC_IMG_COMPANYLOGO_%', $base64companyLogo)
	
	$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_RUNBOOK_%', $SubRunbookName)
	$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_ACTIVITY_%', $object.Name.'#text')
	
	# Now add (embed) the image of the individual activity
	$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_ICON_ACTIVITY_%', $iconForActivity)
	
	# Update the date of this document being generated:
	$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_TODAYS_DATE_%', "$(Get-Date) by $([Environment]::UserName)")
	
	# And the coup de gras - lets add the reference image for storage in the file itself
	$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_PLACEHOLDER_RUNBOOK_IMAGE_%', $picRunbookReference)
	
	# if there's no description we need to make sure to have something
	if ($object.Description.'#text' -ne $null)
	{
		$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_DESCRIPTION_%', "<b>" + $object.Description.'#text' + "</b>")
	}
	else
	{
		# Replace the description with some html code and make it red to ensure people see this needs to be updated
		$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_DESCRIPTION_%', "<span style=`"color:red;`">(empty / null)</span>")
	}
	
	# Update Objecttypename
	$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_OBJECTTYPENAME_%', $object.ObjectTypeName.'#text')
	
	Log-Message -logMsg "[ GENERAL ACTIVITY ]: Obtaining property name and values and then creating the html for each property (table row - 2 columns [name][value])"
	
	
	# Do it smart - exclude 3 items from your object created earlier and loop through them to build each row:
	$arrPropertyList = $objActivity | gm | ? { $_.MemberType -eq "NoteProperty" -and $_.Name -ne "Name" -and $_.Name -ne "Description" -and $_.Name -ne "ObjectTypeName" }
	
	# Get a counter for our loop
	$arrPropertyList.Count
	
	# Declare empty string to append row data to
	[string]$rowHTML
	
	# For each item in the array we just made
	for ($y = 0; $y -le ($arrPropertyList.Count - 1); $y++)
	{
		
		# Declare item as temp variable
		$thisProperty = $objActivity."$($arrPropertyList[$y].Name)"
		#Log-Message -logMsg "[ GENERAL ACTIVITY ]: Property Name: $($arrPropertyList[$y].Name)"
		#Log-Message -logMsg "[ GENERAL ACTIVITY ]: Property Value: $($thisProperty.'#text')"
		# As long as there is data in the value field
		if ($thisProperty.datatype -eq "null" -or ([String]::IsNullOrEmpty($thisProperty))){
		
			# If the property value is empty, update the HTML to reflect it so it shows SOMETHING
			if($ShowNullorEmptyProps -eq $true) {
				$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">{1}</td></tr>`n" -f $thisProperty.Name, "(null/empty)"
			}
		}
		else
		{	
			switch($thisProperty.Name) {
			# ADD ADDITIONAL CONDITIONS FOR VARIOUS PARAMETERS YOU WISH TO SPECIALLY FORMAT
			"SourceObject"
			{
				# A link activity has the Target of itself identified here
				#Log-Message -logMsg "[Process-GeneralActivityReport]: Processing as 'SourceObject'..."
				$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">{1}</td></tr>`n" -f $thisProperty.Name, $objActivity.SourceObject.UniqueID.'#text'
				break;
			}
			"TargetObject"
			{
				# A link activity has the Target of itself identified here
				#Log-Message -logMsg "[Process-GeneralActivityReport]: Processing as 'TargetObject'..."
				$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">{1}</td></tr>`n" -f $thisProperty.Name, $objActivity.TargetObject.UniqueID.'#text'
				break;
			}
			"CUSTOM_START_PARAMETERS"
			{
				# This is the data sent to the data bus from a "Initialize Data" Activity - usually from previous runbooks or static parameters
				#Log-Message -logMsg "[Process-GeneralActivityReport]: Processing as 'CUSTOM START PARAMETERS'..."
				$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">`n" -f $thisProperty.Name
				
				$rowHTML += $objActivity.'CUSTOM_START_PARAMETERS'
				
				$rowHTML += "</td></tr>`n"
				#Log-Message "[ Get-CustomStartParameters ]: Processing Complete"
				break;
				
			}
			"PUBLISH_POLICY_DATA"
			{
				# This is the data sent back to the data bus from a "Return Data" Activity
				#Log-Message -logMsg "[Process-GeneralActivityReport]: Processing as 'PUBLISH_POLICY_DATA'..."
				$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">`n" -f $thisProperty.Name
				
				$rowHTML += $objActivity.'PUBLISH_POLICY_DATA'
				
				$rowHTML += "</td></tr>`n"
				#Log-Message "[ Get-PublishedPolicyData ]: Processing Complete"
				break;
			}
			"StringTestOption" {
					# We want to "humanize" the codes for the actual condition we are comparing our values for
					switch($objActivity.StringTestOption.'#text') {
					'1' {
							$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">`n" -f $thisProperty.Name;
							$rowHTML += ($thisProperty.'#text' + " - <i>(is different than)</i>");
							$rowHTML += "</td></tr>`n";
							break;
						}
					'2' {
							$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">`n" -f $thisProperty.Name;
							$rowHTML += ($thisProperty.'#text' + " - <i>(is equal to)</i>");
							$rowHTML += "</td></tr>`n";
							break;
						}
					'7' {
							$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">`n" -f $thisProperty.Name;
							$rowHTML += ($thisProperty.'#text' + " - <i>(matches the pattern)</i>");
							$rowHTML += "</td></tr>`n";
							break;
						}
					'8' {
							$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">`n" -f $thisProperty.Name;
							$rowHTML += ($thisProperty.'#text' + " - <i>(does not match the pattern)</i>");
							$rowHTML += "</td></tr>`n";
							break;
						}
					'9' {
							$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">`n" -f $thisProperty.Name;
							$rowHTML += ($thisProperty.'#text' + " - <i>(is not empty)</i>");
							$rowHTML += "</td></tr>`n";
							break;
						}
					default {
							$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">`n" -f $thisProperty.Name;
							$rowHTML += ($thisProperty.'#text' + " - <i>(undefined parameter 'StringTestOption' - Line 1638)</i>");
							$rowHTML += "</td></tr>`n";
							break;						
						}
					}
					Log-Message "[ Get-PublishedPolicyData ]: Processing Complete"
				}
				"ExecutionData" {
					
					if($rowHTML -notlike "*ExDataDescription*") {
						# Get the data and transform it to XML for easier processing
						$exData = [xml]$thisProperty.'#text';
						
						# Build a table tag and some headers for the data
						$rowHTML += "<tr><td class=`"alt`">{0}</td><td colspan=`"3`">`n" -f $thisProperty.Name;
						$rowHTML += "`t`t<div style=`"float:left;padding-left:10px;margin-top:0;padding-top:0;`" class=`"ExDataDescBox`">`n";
						# this is where we add the dropdown into the row we are in the middle of creating
						$rowHTML += "`t`t`t<fieldset><p><label>Select ID:</label>"
						
						$rowHTML += "<select id=`"ExecutionDataList`">"
						$rowHTML += "`t`t`t`t`t<option value=`"SelectItem`">Select an Item...</option>`n"
						
						# For each Item in the XML data, process it as an item for the drop down
						$exData.ExecutionData.Item | % {
							$thisItem = $_	
							$rowHTML += "`t`t`t`t`t<option value=`"$($thisItem.ID.Replace(' ',$null))`">$($thisItem.Text)</option>`n"
						}
						
						# close the field set and cell, and move to the next cell
						$rowHTML +=  "`t`t`t</select></p></fieldset>`n"
						$rowHTML += "`t`t</div>`n"
						$rowHTML += "`t`t<div style=`"float:left;padding-left:10px;margin-top:0;padding-top:0;`">`n`t`t`t"
						$rowHTML += "<div id=`"SelectItem`" class=`"ExDataDescription`">Choose an option, and the description will appear here</div>"						
						
						$exData.ExecutionData.Item | % {
							$thisItem = $_	
							$rowHTML += "<div id=`"$($thisItem.ID.Replace(' ',$null))`" class=`"ExDataDescription`">$($thisItem.Description)<br><b>Type</b>: $($thisItem.Type)</div>"						
						}
						
						$rowHTML += "`n`t`t</div>`n";

						# Close out the table and the property Row
						$rowHTML += "</td></tr>`n";
					}
					break;
				}
				'Properties' {
				
					[string]$formattedCode = "$($null)";
					# We should process the code a little to make the HTML viewable (too much weird formatting by default to be readable)
					$rowHTML += "<tr>`n`t<td class=`"alt`">{0}</td>`n`t<td>" -f $arrPropertyList[$y].Name
					
					if ($thisProperty.'#text' -like "*ItemRoot*") {

						# Convert these properties to xml for further formatting
						$xmlData = [xml]$thisProperty.'#text';
						
						if($xmlData.ItemRoot.Entry.Count -eq $null -and $xmlData.ItemRoot.Entry -ne $null) {
								# Get the current item in the list
								$thisEntry = $xmlData.ItemRoot.Entry;
								
								# Get the stuff we need
								$propName = $thisEntry.PropertyName.'#cdata-section'.Replace('\`~F/',$null);
								$propValue = $thisEntry.PropertyValue.'#cdata-section'.Replace('\`~F/',$null);

								$formattedCode += "<b>Name</b>: $($propName) - <b>Value</b>: $($propValue)`r"	
								
						} else {
						
							$xmlData.ItemRoot.Entry | % {
								
								# Get the current item in the list
								$thisEntry = $_;
								
								# Get the stuff we need
								$propName = $thisEntry.PropertyName.'#cdata-section'.Replace('\`~F/',$null);
								$propValue = $thisEntry.PropertyValue.'#cdata-section'.Replace('\`~F/',$null);

								$formattedCode += "<b>Name</b>: $($propName) - <b>Value</b>: $($propValue)`r"
							}
						}
					}
					
					$rowHTML += $formattedCode;
					
					$rowHTML += "`n`t</td></tr>`n"; 
					
				}
				default
				{
					#Log-Message -logMsg "[Process-LinkActivityReport]: Processing as normal property (not target or source object)..."
					
					# Build an HTML Row based on the current Property Name and Value
					# If this property isn't some identifier for the current activity
					if ($thisProperty.Name -ne $null -and $thisProperty.Name -notlike "*ID" -and $thisProperty.'#text' -match "{*}")
					{		
						$rowHTML += "<tr><td class=`"alt`">{0}</td><td>{1}</td></tr>`n" -f $arrPropertyList[$y].Name, $thisProperty.'#text'
					}
					else
					{
						# Build an HTML Row based on the current Property Name and Value
						$rowHTML += "<tr><td class=`"alt`">{0}</td><td>{1}</td></tr>`n" -f $arrPropertyList[$y].Name, $thisProperty.'#text'
						
						#Log-Message -logMsg "[Process-ActivityReport]: ================================================================"
						#Log-Message -logMsg "[Process-ActivityReport]: NOT A GUID, PROCESSING AS NORMAL"
						#Log-Message -logMsg "[Process-ActivityReport]: ================================================================"
					}
					break;
				}
			}
		}
	}
	
	# Now that we have our loop done, update the html we just created to insert the new rows that contain property attributes
	$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_PLACEHOLDER_LINKPROP_INFORMATION_%', $rowHTML)
	
	# This is the HTML to use if there is not any "Published Data" for the Activity
	$replacementHTML = @"
	<tr colspan="4">
		<td>No Published Data For This Activity</td>
	</tr>
"@
	$tmpGeneralActivityHTML = $tmpGeneralActivityHTML.Replace('%_PLACEHOLDER_PUBLISHED_DATA_%', $replacementHTML)
	
	
	# Do our update of what HTML weve built to the temporary report template html:
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_ACTIVITY_INFORMATION_%', $tmpGeneralActivityHTML)
	
	# Before sending off the data, Clean up the data a bit
	# --------------------------------------------------------------------------------------------
	# Replace the formatting for subscription data
	$tmpHTMLReport = $tmpHTMLReport.Replace('\`d.T.~Ed/', $null)
	$tmpHTMLReport = $tmpHTMLReport.Replace('\`d.T.~Vb/', $null)
	
	# Instead of retyping wierd code - lets assign the name to a variable
	$strActivityName = $object.Name.'#text'
	$strLinkID = $strActivityName.Replace(' ', $null)
	$strHTMLLinkText = "$($counter)-$($strLinkID)"
	
	# Now get the coordinates of the activity icon and generate an "area" for our map to be highlighted for this page
	# The JQuery setting for this instance of the maphilight plugin is "alwaysOn:true" so we don't need to mouseover the image
	$intYPosition = $object.PositionY.'#text'
	$intXPosition = $object.PositionX.'#text'
	
	# Generate HTML for an area (coordinates for the red dot as an image map)
	$strJQueryAreaMapTemplate = Generate-RunbookCoords -xPath $intXPosition -yPath $intYPosition -id $strHTMLLinkText
	
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_ACTIVITY_COORDS_%', $strJQueryAreaMapTemplate)
	
	# Obtain the template we have for our CSS Menu that will show links (associated with this object/Activity) to the user
	$templateCSSMenu = (Get-Content "$($CSSMenuTemplate)") -join "`n"
	
	# Now update the actual HTML template with our CSS Menu (that also was generated from a template)
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_CSS_LINK_MENU_%', $templateCSSMenu)
	
	# Last but not least: Update the CSS with our style sheet code (this will embed it in the doc so we don't need another location to store the css file)
	$stylesheetData = (Get-Content "$($StyleSheetPath)") -join "`n"
	$tmpHTMLReport = $tmpHTMLReport.Replace('%_PLACEHOLDER_LAYOUT_CSS_%', $stylesheetData)
	
	# Send our formatted/dolled up string to the file waiting for it, creating our report
	$tmpHTMLReport | Out-File -FilePath $GeneralActivityFile -Append -Encoding UTF8
	
}

Function Process-TableOfContentFiles($runbookToCFiles,$ToCMenuCSS){

	# For each runbook directory
	for($r = 0; $r -le ($runbookToCFiles.Count - 1);$r++) {
		
		# Get the current directory
		$tocFile = $runbookToCFiles[$r];
	
		if($tocFile -ne $null) {
			
			Log-Message "Updating Left CSS Menu for $($tocFile.FullName) ..."
			$ToCData = Get-Content $tocFile.FullName
			
			Log-Message "Generated the following HTML:"
			Log-Message "$($ToCMenuCSS)"	
			Log-Message "Writing data to table of contents file: $($tocFile.FullName)"
			$ToCData.Replace('%_PLACEHOLDER_LEFTMENUHTML_%',$ToCMenuCSS) | Set-Content $tocFile.FullName -Force -Confirm:$false
		
		} else { Log-Message "No Index.html file was found in $($thisDir.FullName)" }
		
	}

}

################################################################################################
# Function to return query results in a datatable format
function Query-SCOrchDB($datasource, $query)
{

	# Queries ripped from SMART Documentation Toolkit on 9-25-2014:
	# SCHEDULING INFO:
	<#               If ($dr["name"]-eq "ScheduleTemplateID") {
                                    #This is a Check Schedule activity, let's provide more details about the schedule itself
                                    $SqlQuery = "select Name from OBJECTS where uniqueID='{" + $TmpProperty + "}'"
                                    $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                                    $dr2 = $myCommand2.ExecuteReader()
                                    while ($dr2.Read()) {WriteToFile -ExportMode $ExportMode -Add ("# Schedule name : " + $dr2["Name"]) -NbTab $NbTab}
                                    $dr2.Close()
                                    $SqlQuery = "select * from SCHEDULES where uniqueID = '{" + $TmpProperty + "}'"
                                    $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                                    $dr2 = $myCommand2.ExecuteReader()
                                    while ($dr2.Read())
                                        {
                                        WriteToFile -ExportMode $ExportMode -Add ("# Schedule details : Days of week = " + $dr2["DaysOfWeek"] + " - Days of Month = "+ $dr2["DaysOfMonth"]) -NbTab $NbTab
                                        WriteToFile -ExportMode $ExportMode -Add ("# Schedule details : Monday = " + $dr2["Monday"] + " - Tuesday = "+ $dr2["Tuesday"] + " - Wednesday = "+ $dr2["Wednesday"] + " - Thursday = "+ $dr2["Thursday"] + " - Friday = "+ $dr2["Friday"] + " - Saturday = "+ $dr2["Saturday"] + " - Sunday = "+ $dr2["Sunday"]) -NbTab $NbTab
                                        WriteToFile -ExportMode $ExportMode -Add ("# Schedule details : First = " + $dr2["First"] + " - Second = "+ $dr2["Second"] + " - Third = "+ $dr2["Third"] + " - Fourth = "+ $dr2["Fourth"] + " - Last = "+ $dr2["Fourth"]) -NbTab $NbTab
                                        WriteToFile -ExportMode $ExportMode -Add ("# Schedule details : Days = " + $dr2["Days"] + " - Hours = "+ $dr2["Hours"] + " - Exceptions = "+ $dr2["Exceptions"]) -NbTab $NbTab
                                        }
                                    $dr2.Close() 
                                    $Global:FlagSchedule = $True                                  
                                    If ($Global:FlagScheduleList.Contains($ActivityName))
                                        {$Global:FlagScheduleNumber[$Global:FlagScheduleList.IndexOf($ActivityName)] = $Global:FlagScheduleNumber[$Global:FlagScheduleList.IndexOf($ActivityName)] +1}
                                    else
                                        {
                                        $Global:FlagScheduleList+= $ActivityName
                                        $Global:FlagScheduleNumber+= 1
                                        }
                                    } #>


	# Build our connection
	#$connectionString = "Server=$($datasource);Initial Catalog='Orchestrator';Integrated Security=True;"
	$connectionString = $SQLConnectionString
	$connection = New-Object System.Data.SqlClient.SqlConnection;
	$connection.ConnectionString = $connectionString;
	try
	{
		# Open a session to the SQL Server (Make sure SQL Port is accessible from machine your connecting from)
		$connection.Open();
		
		# Build and execute the Query
		$command = $connection.CreateCommand();
		$command.CommandText = $query;
		
		# Get our results of the query
		$results = $command.ExecuteReader();
		
		# Build a datatable object which we will store our results in
		$table = New-Object "System.Data.DataTable";
		$table.Load($results);
		
		# Close the connection to sql
		$connection.Close();
		
		# Return the data as a DataTable object so we can do other stuff if we want
		return $table;
		
	}
	catch
	{
		# Throw an error that we can use to debug if necessary
		Show-MsgBox -Title "Error Querying SCOrch Database" -Prompt "An error occurred while querying the SCOrch Database.`r`nThe error was:`r'$($Error[0])" -Icon 'Critical' -DefaultButton '1'
		return $null;
	}
}

################################################################################################
<# 
            .SYNOPSIS  
            Shows a graphical message box, with various prompt types available. 
 
            .DESCRIPTION 
            Emulates the Visual Basic MsgBox function.  It takes four parameters, of which only the prompt is mandatory 
 
            .INPUTS 
            The parameters are:- 
             
            Prompt (mandatory):  
                Text string that you wish to display 
                 
            Title (optional): 
                The title that appears on the message box 
                 
            Icon (optional).  Available options are: 
                Information, Question, Critical, Exclamation (not case sensitive) 
                
            BoxType (optional). Available options are: 
                OKOnly, OkCancel, AbortRetryIgnore, YesNoCancel, YesNo, RetryCancel (not case sensitive) 
                 
            DefaultButton (optional). Available options are: 
                1, 2, 3 
 
            .OUTPUTS 
            Microsoft.VisualBasic.MsgBoxResult 
 
            .EXAMPLE 
            C:\PS> Show-MsgBox Hello 
            Shows a popup message with the text "Hello", and the default box, icon and defaultbutton settings. 
 
            .EXAMPLE 
            C:\PS> Show-MsgBox -Prompt "This is the prompt" -Title "This Is The Title" -Icon Critical -BoxType YesNo -DefaultButton 2 
            Shows a popup with the parameter as supplied. 
 
            .LINK 
            http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.msgboxresult.aspx 
 
            .LINK 
            http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.msgboxstyle.aspx 
            #>
# By BigTeddy August 24, 2011
# http://social.technet.microsoft.com/profile/bigteddy/.
function Show-MsgBox
{
	
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true)] [string]$Prompt,
		[Parameter(Position = 1, Mandatory = $false)] [string]$Title = "",
		[Parameter(Position = 2, Mandatory = $false)] [ValidateSet("Information", "Question", "Critical", "Exclamation")] [string]$Icon = "Information",
		[Parameter(Position = 3, Mandatory = $false)] [ValidateSet("OKOnly", "OKCancel", "AbortRetryIgnore", "YesNoCancel", "YesNo", "RetryCancel")] [string]$BoxType = "OkOnly",
		[Parameter(Position = 4, Mandatory = $false)] [ValidateSet(1, 2, 3)] [int]$DefaultButton = 1
	)
	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null
	switch ($Icon)
	{
		"Question" { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Question }
		"Critical" { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Critical }
		"Exclamation" { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Exclamation }
		"Information" { $vb_icon = [microsoft.visualbasic.msgboxstyle]::Information }
	}
	switch ($BoxType)
	{
		"OKOnly" { $vb_box = [microsoft.visualbasic.msgboxstyle]::OKOnly }
		"OKCancel" { $vb_box = [microsoft.visualbasic.msgboxstyle]::OkCancel }
		"AbortRetryIgnore" { $vb_box = [microsoft.visualbasic.msgboxstyle]::AbortRetryIgnore }
		"YesNoCancel" { $vb_box = [microsoft.visualbasic.msgboxstyle]::YesNoCancel }
		"YesNo" { $vb_box = [microsoft.visualbasic.msgboxstyle]::YesNo }
		"RetryCancel" { $vb_box = [microsoft.visualbasic.msgboxstyle]::RetryCancel }
	}
	switch ($Defaultbutton)
	{
		1 { $vb_defaultbutton = [microsoft.visualbasic.msgboxstyle]::DefaultButton1 }
		2 { $vb_defaultbutton = [microsoft.visualbasic.msgboxstyle]::DefaultButton2 }
		3 { $vb_defaultbutton = [microsoft.visualbasic.msgboxstyle]::DefaultButton3 }
	}
	$popuptype = $vb_icon -bor $vb_box -bor $vb_defaultbutton
	$ans = [Microsoft.VisualBasic.Interaction]::MsgBox($prompt, $popuptype, $title)
	return $ans
} #end function

################################################################################################
# This LinkCondition function was ripped blatantly from the SMART Documentation Toolkit on TechNet:
# http://gallery.technet.microsoft.com/SMART-Documentation-and-f28fc304
# MA - 09/25/2014
function LinkCondition
# This function retrieves the condition on a link between activities (if any)
{
    param (
    [String]$LinkID,
	[String]$SQLConnectionString
    )
	$datasource = $script:SCORCHDBServer
    $output = ""
    $OutputID = ""
	
	$connectionString = "$($SQLConnectionString)"
	$myConnection = New-Object System.Data.SqlClient.SqlConnection;
	$myConnection.ConnectionString = $connectionString;
	
	# Open a session to the SQL Server (Make sure SQL Port is accessible from machine your connecting from)
	$myConnection.Open();
		
	# Build and execute the Query

    $SqlQuery = "select condition, data, value from TRIGGERS where ParentID = '" + $LinkID + "'"
    $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
	
    $dr = $myCommand.ExecuteReader()
    $OffsetBracket=0
    while ($dr.Read())
            {
            $OutputID = ($dr["data"]).Substring(0, 38)
            $NumberofGUIDs = ($dr["data"]).Split("{").Count - 1
            Switch ($dr["condition"])
                {
                "isgreaterthan" {$Outputcondition = "-gt"}
                "isgreaterthanorequalto" {$Outputcondition = "-ge"}
                "islessthan" {$Outputcondition = "-lt"}
                "islessthanorequalto" {$Outputcondition = "-le"}
                "equals" {$Outputcondition = "-eq"}
                "doesnotequal" {$Outputcondition = "-ne"}
                "" {$Outputcondition = "{linkcondition:returns}" ; $OffsetBracket = 1}                
                default
                    # "contains" "doesnotcontain" "endswith" "startswith" "doesnotmatchpattern" "matchespattern"
                    {
                    $Outputcondition = "{linkcondition:" + $dr["condition"] + "}"
                    $Global:FlagStringcondition = $True
                    $OffsetBracket = 1
                    }
                }
            $Output = "If (" + $dr["data"] + " " + $Outputcondition + " `"" + $dr["value"] + "`") {"
            }
    $dr.Close()
    If ($OutputID -ne "")
            {
            #There was a condition, let's convert the published data activity (first part)
            $SqlQuery = "select name, ObjectType from objects where UniqueID = '" + $OutputID + "'"
            $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
            $dr = $myCommand.ExecuteReader()
            while ($dr.Read()) {$OutputName = $dr["name"]}
            $dr.Close()
            $Output = $Output.Replace($OutputID, "{Activity:" + $OutputName + "}")
            #Let's check if there is a second GUID to convert - only applicable when it's published data from an initialize data activity
            $NumberofGUIDs = $output.Split("{").Count - 1 -$OffsetBracket
            If ($NumberofGUIDs -eq 3)
                {
                $OutputSuffix = "{" + $output.Split("{")[2].Substring(0, 36) + "}"
                $SqlQuery = "select value from CUSTOM_START_PARAMETERS where UniqueID = '" + $OutputSuffix + "'"
                $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                $dr = $myCommand.ExecuteReader()
                while ($dr.Read()) {$OutputSuffixName = $dr["value"]}                    
                $dr.Close()
                $OutputSuffix = "}." + $OutputSuffix
                $Output = $output.Replace($OutputSuffix, ".PublishedData" + $OutputSuffixName + "}")
                }
            }
    Return $Output
}

################################################################################################
# I'll give you one guess to figure out what this function does... :)
Function Update-GUIDsWithHumanSpeak($file,$runbook){
	# Regex for GUIDs
	$regex = '(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}'

	# I decided to filter out 'UniqueID','SourceObject', and 'TargetObject' from being updated since it might be useful to know the actual GUID during log troubleshooting
	# (Be that from the script logs or Orchestrator logs themselves due to runbook failures)
	$matchedGUIDs = select-string -Path $file -Pattern $regex -AllMatches | ? { 
		($_.Line -notlike "*UniqueID*") -and ($_.Line -notlike "*TargetObject*") -and ($_.Line -notlike "*SourceObject*")
	}

	$htmlContent = (Get-Content $file) -join "`n"
	
	for($l=0;$l -le ($matchedGUIDs.matches.Count - 1);$l++) { 
	
		# Get the current item from the matched set
		$thisGUID = $matchedGUIDS.matches[$l].Value
		
		$thisMatchedLine = $matchedGUIDs[$l].Line 
		
		
		# Try to find where this object is in the runbook
		$thisObject = Get-ObjectByID -objectID "$($thisGUID)" -runbookObject $runbook
		
		
		# If it could not be found locally, we need to then look to the database for info
		if($thisObject -eq $null -and $thisMatchedLine -notlike "*UniqueID*") {
			
			# Before searching the databases - we can search each runbook for anything that matches
			$nodeFolder | % { 
			
				$thisRunbook = $_
				
				if($thisRunbook.Name -ne $runbook.Name) {
				
					Log-Message "[Update-GUIDS]: Searching for object in $($thisRunbook.Name)..."
					# try to find the object in each runbook
					$thisObject = Get-ObjectByID -objectID "$($thisGUID)" -runbookObject $thisRunbook
					
					if($thisObject -ne $null) {
						
						# Use this variable to update the HTML
						$htmlContent = $htmlContent.Replace("$($thisGUID)","<i><b>($($thisObject.Name.'#text'))</b></i>")
						Log-Message "[Update-GUIDS]: Found object in Runbook ($($thisRunbook.Name)...Name: $($thisObject.Name.'#text')..."
						# break out of the loop, we're done here
						break;
						
					} else {
					
						Log-Message "[Update-GUIDS]: Object not found in that runbook...next..."
						
					}
			
				}
			
			}
			
						
			# We've already looked through the export file, so now we need to remove the '{}' in the GUID
			# To properly query the database tables
			$thisGUID = $thisGUID.Replace('{',$null);
			$thisGUID = $thisGUID.Replace('}',$null);

			Log-Message "[Update-GUIDS]: Attempting to find object from SCORCH Database (ObjectTypes)...please wait..."
			$sqlQuery = "SELECT * FROM [ObjectTypes] WHERE UniqueID = '$($thisGUID)'"
			$thisObject = Query-SCOrchDB -datasource $SCORCHDBServer -query $sqlQuery
			
			# THIS CAN BE CLEANED UP TREMENDOUSLY - I WAS RUSHING TO GET A WORKING EXAMPLE OUT INTO THE NETS
			# THIS GIANT IF/ELSE NEST is basically first looking through the export file, then looping through each database table specified and looking for the GUID
			if($thisObject -ne $null){
					
				if(!([String]::IsNullOrEmpty($thisObject.Name))) {
					# Use this variable to update the HTML
					$htmlContent = $htmlContent.Replace("$($thisGUID)","<i><b>($($thisObject.Name))</b></i>")
					Log-Message "[Update-GUIDS]: Found object in SCORCH Database (ObjectTypes)...Name: $($thisObject.Name)..."
				}
		
			} else {
				
				Log-Message "[Update-GUIDS]: Attempting to find object from SCORCH Database (Objects)...please wait..."
				$sqlQuery = "SELECT * FROM [Objects] WHERE UniqueID = '$($thisGUID)'"
				$thisObject = Query-SCOrchDB -datasource $SCORCHDBServer -query $sqlQuery

				if($thisObject -ne $null){
							
					if(!([String]::IsNullOrEmpty($thisObject.Name))) {
						# Use this variable to update the HTML
						$htmlContent = $htmlContent.Replace("$($thisGUID)","<i><b>($($thisObject.Name))</b></i>")
						Log-Message "[Update-GUIDS]: Found object in SCORCH Database (Objects)...Name: $($thisObject.Name)..."
					}
			
				} else {
					
					Log-Message "[Update-GUIDS]: Attempting to find object from SCORCH Database (VARIABLES)...please wait..."
					$sqlQuery = "SELECT * FROM [VARIABLES] WHERE UniqueID = '$($thisGUID)'"
					$thisObject = Query-SCOrchDB -datasource $SCORCHDBServer -query $sqlQuery
			
					if($thisObject -ne $null){
					
						# Use this variable to update the HTML
						$htmlContent = $htmlContent.Replace("$($thisGUID)","<i><b>($($thisObject.Value))</b></i>")
						Log-Message "[Update-GUIDS]: Found object in SCORCH Database (Variables)...Value: $($thisObject.Value)..."
						
					} else {
				
						
						Log-Message "[Update-GUIDS]: Attempting to find object from SCORCH Database (TRIGGER_POLICY_PARAMETERS)...please wait..."
						$sqlQuery = "SELECT * FROM [TRIGGER_POLICY_PARAMETERS] WHERE UniqueID = '$($thisGUID)'"
						$thisObject = Query-SCOrchDB -datasource $SCORCHDBServer -query $sqlQuery
						
						if($thisObject -ne $null){
						
							if(!([String]::IsNullOrEmpty($thisObject.ParameterName))) {
								# Use this variable to update the HTML
								$htmlContent = $htmlContent.Replace("$($thisGUID)","<i><b>($($thisObject.ParameterName))</b></i>")
								Log-Message "[Update-GUIDS]: Found object in SCORCH Database (TRIGGER_POLICY_PARAMETERS)...ParameterName: $($thisObject.ParameterName)..."
							}
							
						} else {
						
							# Last chance - try the custom start parameters table
							Log-Message "[Update-GUIDS]: Attempting to find object from SCORCH Database (CUSTOM_START_PARAMETERS)...please wait..."
							$sqlQuery = "SELECT * FROM [CUSTOM_START_PARAMETERS] WHERE UniqueID = '$($thisGUID)'"
							$thisObject = Query-SCOrchDB -datasource $SCORCHDBServer -query $sqlQuery
						
							if($thisObject -ne $null){

									# Use this variable to update the HTML
									$htmlContent = $htmlContent.Replace("$($thisGUID)","<i><b>($($thisObject.Value))</b></i>")
									Log-Message "[Update-GUIDS]: Found object in SCORCH Database (CUSTOM_START_PARAMETERS)...Value: $($thisObject.Value)..."
								
							} else { 
						
								Log-Message "[Update-GUIDS]:[Warning]: Unable to locate GUID: $($thisGUID) in any of the database tables or XML files"
								Log-Message "[Update-GUIDS]:[Warning]: (Should Investigate if I'm Missing Important GUIDs or if they are irrelavent when doing quick troubleshooting)..."
							}
						}
						
					}	
					
				}
				
			}
		
		} else {
		
			if(!([String]::IsNullOrEmpty($thisObject.Name))) {
				# Use this variable to update the HTML
				$htmlContent = $htmlContent.Replace("$($thisGUID)","<i><b>($($thisObject.Name.'#text'))</b></i>")
				#Log-Message "Found object in Runbook...Name: $($thisObject.Name.'#text')..."
			}
		}

	}
	
	# Take some formatting that was residual out so the properties are in more of a "list" format:
	$htmlContent = $htmlContent.Replace('\`~F/]]>','<br>')
	$htmlContent = $htmlContent.Replace('\`d.T.~Br/',$null)
	

	# Now that all processing has been completed, let's rewrite the HTML with the friendly info
	$htmlContent | out-file $file -force 
}


Function Process-RunbookFolders($runbookPolicy,$runbookFolderName) {

	# Build the string we want to use for the Runbook Folder (No Spaces)
	$runbookExportFolder = $ReportExportPath + $runbookFolderName
	
	# Attempt to create the root folders for runbook activity documentation (returns true/false)
	$folderCreated = Organize-RunbookFolders -rootPath $runbookExportFolder
	
	if ($folderCreated -eq $true)
	{
		# Build a space-less version of the name and tack .html on the end
		$script:fileName = $runbookExportFolder.Replace(' ', $null).Trim() + ".html"
		
		# Build a file name so we can pass it to the function that builds the file
		$ActivityTOCFile = $fileName.Replace(".html", "\Index.html")
		
		# Get all steps for the table of contents
		$arrAllActivities = $runbookPolicy.Object
	
		# This is the function that will build the file
		Build-TableOfContents -arrAllActivities $arrAllActivities -ActivityTOCFile $ActivityTOCFile -runbook $runbookPolicy
		
		$objectCounter = 0
		
		# For each of these objects
		for ($x = 0; $x -le ($arrAllActivities.Count - 1); $x++)
		{
			# Declare a variable for the current object and increment the counter for the links
			$thisObject = $arrAllActivities[$x]
			$objectCounter++
			
			# Depending on the objects "ObjectTypeName' property, let's process the data accordingly
			switch ($thisObject.ObjectTypeName.'#text')
			{
				'Run .Net Script' {
					
					Write-Host "Processing Report for: " -ForegroundColor Yellow -NoNewline
					Write-Host "$($thisObject.ObjectTypeName.'#text')..." -ForegroundColor White -nonewline
					Write-Host ""
					Process-PowerShellReport -object $thisObject -counter $objectCounter -icon $arrIcons[1].FullName -runbook $runbookPolicy
					Write-Host "Processing Complete!" -ForegroundColor White
					Write-Host ""
					break;
				}
				'Query Database' {
					
					Write-Host "Processing Report for: " -ForegroundColor Yellow -NoNewline
					Write-Host "$($thisObject.ObjectTypeName.'#text')..." -ForegroundColor White -nonewline
					Write-Host ""
					Process-DBQueryReport -object $thisObject -counter $objectCounter -icon $arrIcons[0].FullName -runbook $runbookPolicy
					Write-Host "Processing Complete!"
					Write-Host ""
					break;
				}
				'Link' {
					
					Write-Host "Processing Report for: " -ForegroundColor Yellow -NoNewline
					Write-Host "$($thisObject.ObjectTypeName.'#text')..." -ForegroundColor White -nonewline

					# Get the GUIDs of the source and target objects
					$sourceID = $thisObject.SourceObject.'#text'
					$targetID = $thisObject.TargetObject.'#text'
					
					# Now pass the object, counter, and our two GUIDs to a function to spit out the html
					Process-LinkActivityReport -object $thisObject -counter $objectCounter -sourceID $sourceID -targetID $targetID -runbook $runbookPolicy
					Write-Host "Processing Complete!"
					Write-Host ""
					break;
				}
				default
				{
					Write-Host "Processing Report for: " -ForegroundColor Yellow -NoNewline
					Write-Host "$($thisObject.ObjectTypeName.'#text')..." -ForegroundColor White -nonewline
					Write-Host ""
					# Now pass the object, counter, and our two GUIDs to a function to spit out the html
					Process-ActivityReport -object $thisObject -counter $objectCounter -icons $arrIcons -runbook $runbookPolicy
					Write-Host "Processing Complete!"
					Write-Host ""
					break;
				}
				
			}
			
		} #end for-each object (Activity)
		
		
	}else{
				
		Write-Host "Error: The Runbook Folders could not be created, review any errors shown above.  This script is also configured to check for the folders existence, and quit if they are already there.  Delete or backup those folders then re-run the script..."
		start-sleep -seconds 7
	}
}

Function Process-RunbookLinks($runbookFolder,$currentPolicy) {
		
		# Now that the files have completed, we need to go back and dig up our links for each activity accordingly
		$ActivityFiles = @()
		$ActivityFiles = Get-ChildItem "$($runbookFolder)" | ? { $_.Name -like "*.html" -and $_.Name -ne "Index.html" }
		
		# For each file / Activity in the folder
		If ($ActivityFiles.Count -gt 0)
		{
			Write-Host "[ POST-CREATION PROCESSING ]: Found $($ActivityFiles.Count) Activities (files) in the directory..." -ForegroundColor White
			Log-Message -logMsg "[ POST-CREATION PROCESSING ]: Found $($ActivityFiles.Count) Activities (files) in the directory..."
			
			# For each Activity File found, process it and update the CSS Links that may (or may not) be there
			$ActivityFiles | % {
			#for ($a = 0; $a -le ($ActivityFiles.Count - 1); $a++)
			#{
				$isActivityALink = $null;
				# get the current file info
				#$thisActivityFileName = $ActivityFiles[$a].Name
				#$thisActivityFilePath = $ActivityFiles[$a].FullName
				$thisActivityFileName = $_.Name
				$thisActivityFilePath = $_.FullName
				$dataHTML = $null
				
				Log-Message -logMsg "[ POST-CREATION PROCESSING ]: Checking ObjectTypeName of this Activity file..."
				$isActivityALink = Get-Content "$($thisActivityFilePath)" | Select-String -Pattern '^.*<td class="alt">ObjectTypeName<\/td><td colspan="3">Link<\/td><\/tr>.*$'
				
				# As long as the results from the regex query aren't null
				if ($isActivityALink -eq $null -and $isActivityALink.Line -eq $null)
				{
					Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: Beginning Processing CSS Menu data for Activity: $($thisActivityFileName)..."
					
					# Obtain the data of the html file
					$dataHTML = Get-Content "$($thisActivityFilePath)"
					
					# Get the unique ID of this activity
					#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Searching for this string:"
					#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: <td class=`"alt`">UniqueID</td>"
					$results = Get-Content "$($thisActivityFilePath)" | Select-String -Pattern '^.*<td class="alt">UniqueID<\/td>.*<\/tr>'
					
					if ($results -ne $null -and $results.Line -ne $null)
					{
						$ActivityID = $null;
						# Lets parse out the crap html and get our GUID that we needed
						$ActivityID = $results.Line -replace '.*<td class="alt">UniqueID<\/td><td>', $null
						$ActivityID = $ActivityID -replace '.*<td class="alt">UniqueID<\/td><td colspan="3">', $null
						$ActivityID = $ActivityID -replace '<\/td><\/tr>.*$', $null
						$ActivityID = $ActivityID.Trim()
						
						#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Found this activity ID: $($ActivityID)"
						#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Obtaining this activity object from the OIS Export Data..."
						$objThisActivity = Get-ObjectByID -objectID "$($ActivityID)" -runbookObject $currentPolicy
						
						# Find the links associated with this object from the export data
						#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Obtaining Links with this activity as a source or target..."
						$linkSourceObjects = @();
						$linkSourceObjects += $currentPolicy.Object | ? { $_.ObjectTypeName.'#text' -eq 'Link' -and $_.Name.'#text' -ne $null -and $_.SourceObject.'#text' -eq "$($ActivityID)" }
						$linkTargetObjects = @();
						$linkTargetObjects += $currentPolicy.Object | ? { $_.ObjectTypeName.'#text' -eq 'Link' -and $_.Name.'#text' -ne $null -and $_.TargetObject.'#text' -eq "$($ActivityID)" }
						
						#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Links with this Activity as a Source: $($linkSourceObjects.Count)"
						#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Links with this Activity as a Target: $($linkTargetObjects.Count)"
						Write-Host "[ POST-CREATION PROCESSING ]: Links sourced FROM This Activity: $($linkSourceObjects.Count)" -ForegroundColor White
						Write-Host "[ POST-CREATION PROCESSING ]: Links targeted AT This Activity: $($linkTargetObjects.Count)" -ForegroundColor White
						Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Processing Target Link(s)..."
						$sourceHTML = $null;
						$targetHTML = $null;
						
						# For each Target Link object in our array
						$linkTargetObjects | % {
		
							# Get the current link object
							$thisLinkObject = $_;
							
							$thisLinkName = $thisLinkObject.Name.'#text';
							
							# Obtain its unique ID
							$thisLinkID = $thisLinkObject.UniqueID.'#text';
							
							# Get the current activity's runbook name
							#$runbookFolderName = $runbookExportFolder
							Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Obtaining link files from:"
							Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: $($runbookFolder)"
							#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: that match the name '*$($thisLinkName).html'"
							
							# Get all the link files from the directory that have a name like the name of this link activity
							$arrLinkHTMLFiles = @()
							$arrLinkHTMLFiles = GCI "$($runbookFolder)" | ? { $_.Name -like "*$($thisLinkName).html" };
							
							Write-Host "[ POST-CREATION PROCESSING ]: Processing Link Files (If Any)...." -ForegroundColor White
							
							# For each of these HTML Files
							$arrLinkHTMLFiles | % {
							
								$thisLinkFile = $_;
								
								# Get the content of this link file
								$linkHTMLContent = Get-Content "$($thisLinkFile.FullName)"
								
								# Search for the unique ID that matches what we found earlier (thisLinkID)
								$findLinkResults = $linkHTMLContent | Select-String -AllMatches "<td class=`"alt`">UniqueID</td><td colspan=`"3`">$($thisLinkID)</td>"
								
								if ($findLinkResults.Line -ne $null)
								{
									# Do not change this: SourceObject is there deliberately (I picked a really crappy labeling system I'm discovering)
									$findTargetID = $linkHTMLContent | Select-String -AllMatches "<td class=`"alt`">SourceObject</td><td colspan=`"3`">"
									$foundTargetID = $null;
									
									if ($findTargetID -ne $null)
									{
										# Remove everything but the GUID itself from the HTML
										$foundTargetID = $findTargetID.Line -replace '.*<td class="alt">SourceObject<\/td><td colspan="3">', $null
										$foundTargetID = $foundTargetID.Replace("<td>", $null)
										$foundTargetID = $foundTargetID.Replace("</td>", $null)
										$foundTargetID = $foundTargetID.Replace("<tr>", $null)
										$foundTargetID = $foundTargetID -replace '<\/tr>.*$', $null
										$foundTargetID = $foundTargetID.Trim()
									}
									
									Log-Message "Obtaining Target Object by id ( $($foundTargetID) ) to get name for CSS link..."
									$targetObject = Get-ObjectByID -objectID $foundTargetID -runbookObject $currentPolicy
									
									$strActName = $objThisActivity.Name.'#text'.Replace(' ', $null);
									
									# We found the right link, so we need to build an html string that points to this file
									# Then we will update the original Activity HTML with the links we build from this
									
									$tmpHTML = "<li><a href=`"$($thisLinkFile.Name)`"><span>$($thisLinkName) - Source: $($targetObject.Name.'#text')</span></a></li>"
									
									# Update our variable with the next bit of html we just made:
									$targetHTML += $tmpHTML
									
								}
							}
						}
						Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: $($linkTargetObjects.Count) Target Link(s) Processed for: $($objThisActivity.Name.'#text')"
						Write-Host "[ POST-CREATION PROCESSING ]: $($linkTargetObjects.Count) Target Link(s) Processed for $($objThisActivity.Name.'#text')" -ForegroundColor Yellow
						Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Processing Source Link(s) originating from this activity..."
						
						# For each Source Link object in our array
						$linkSourceObjects | % {
							# Get the current link object
							$thisLinkObject = $_;
							
							$thisLinkName = $thisLinkObject.Name.'#text';
							
							# Obtain its unique ID
							$thisLinkID = $thisLinkObject.UniqueID.'#text';
							
							# Get the current activity's runbook name
							#$runbookFolderName = $SubRunbookName.Replace(' ', $null);
							Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Obtaining link files from:"
							Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: $($runbookFolder)"
							#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: that match the name '$($thisLinkName).html'"
							
							# Get all the link files from the directory that have a name like the name of this link activity
							$arrLinkHTMLFiles = @()
							$arrLinkHTMLFiles = GCI "$($runbookFolder)" | ? { $_.Name -like "*$($thisLinkName).html" };
							
							# For each of these HTML Files
							for ($b = 0; $b -le ($arrLinkHTMLFiles.Count - 1); $b++)
							{
								$thisLinkFile = $arrLinkHTMLFiles[$b];
								
								# Get the content of this link file
								$linkHTMLContent = Get-Content "$($thisLinkFile.FullName)"
								
								# Search for the unique ID that matches what we found earlier (thisLinkID)
								$findLinkResults = Get-Content "$($thisLinkFile.FullName)" | Select-String -AllMatches "<td class=`"alt`">UniqueID</td><td colspan=`"3`">$($thisLinkID)</td>"
								
								if ($findLinkResults.Line -ne $null)
								{
									$findSourceID = Get-Content "$($thisLinkFile.FullName)" | Select-String -AllMatches "<td class=`"alt`">TargetObject</td><td colspan=`"3`">"
									
									if ($findSourceID -ne $null)
									{
										# Remove everything but the GUID itself from the HTML
										$foundSourceID = $findSourceID.Line -replace '.*<tr><td class="alt">TargetObject<\/td><td colspan="3">', $null
										$foundSourceID = $foundSourceID.Replace("<td>", $null)
										$foundSourceID = $foundSourceID -replace '<\/td><\/tr>.*$', $null
										$foundSourceID = $foundSourceID.Replace("<tr>", $null)
										$foundSourceID = $foundSourceID.Trim()
										
									}
									
									#Log-Message "Obtaining Source Object by id ( $($foundSourceID) ) to get name for CSS link..."
									$srcObject = Get-ObjectByID -objectID $foundSourceID -runbookObject $currentPolicy
									
									# We found the right link, so we need to build an html string that points to this file
									# Then we will update the original Activity HTML with the links we build from this
									
									$tmpHTML = "<li><a href=`"$($thisLinkFile.Name)`"><span>$($thisLinkName) - Target: $($srcObject.Name.'#text')</span></a></li>"
									
									# Update our variable with the next bit of html we just made:
									$sourceHTML += $tmpHTML

								}
								
							}
						}
						Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: $($linkSourceObjects.Count) Source Link(s) Processed for $($objThisActivity.Name.'#text')"
						Write-Host "[ POST-CREATION PROCESSING ]: $($linkSourceObjects.Count) Source Link(s) Processed for $($objThisActivity.Name.'#text')" -ForegroundColor Yellow
						
						if ($targetHTML -ne $null)
						{
							# Update this activity's html with the new css links we built from above:
							#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Replacing temp placeholder for targets in Activity HTML File..."
							$dataHTML = $dataHTML.Replace('%_PLACEHOLDER_LINKS_TARGETED_HERE_%', $targetHTML);
						}
						else
						{
							
							$dataHTML = $dataHTML.Replace('%_PLACEHOLDER_LINKS_TARGETED_HERE_%', "<li><a href='#'><span>(No links targeted at this Activity)</span></a></li>");
							
						}
						
						if ($sourceHTML -ne $null)
						{
							#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]: Replacing temp placeholder for sources in Activity HTML File..."
							$dataHTML = $dataHTML.Replace('%_PLACEHOLDER_LINKS_SOURCED_HERE_%', $sourceHTML);
						}
						else
						{
							$dataHTML = $dataHTML.Replace('%_PLACEHOLDER_LINKS_SOURCED_HERE_%', "<li><a href='#'><span>(No links from this Activity)</span></a></li>");
						}
						
						# Now recommit the file and overwrite what we had before:
						$dataHTML | Out-File "$($thisActivityFilePath)" -Force -Confirm:$false -Encoding UTF8
						Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:  $($thisActivityFilePath) - Update of Menu CSS Complete!"
					}
					
				}else{
					Write-Host "[ POST-CREATION PROCESSING ]: Beginning Processing Area Map data for Activity: $($thisActivityFileName)..." -ForegroundColor White
					Write-Host "[ POST-CREATION PROCESSING ]: Link Activity Identified, skipping CSS Menu Generation for:"	-foregroundcolor Yellow
					Write-Host "[ POST-CREATION PROCESSING ]: $($thisActivityFilePath)" -foregroundcolor white
					
					Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: Beginning Processing Area Map data for Activity: $($thisActivityFileName)..."
					Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: Link Activity Identified, skipping CSS Menu Generation for: $($thisActivityFilePath)"
					
					# Get the unique ID of this activity
					#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: Searching for this string:"
					#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: <td class=`"alt`">UniqueID</td>"
					
					$result = Get-Content "$($thisActivityFilePath)" | Select-String -AllMatches '^.*<td class="alt">UniqueID<\/td>.*<\/tr>.*$'
					
					$srcURL = $null;
					$targetURL = $null;
					$thisLinkURL = $null;
					$source = $null;
					$target = $null;
					#$thisRunbookExportFolder = "$($ReportExportPath)" + $thisRunbook.Name.Replace(' ',$null)
					
					if ($result -ne $null -and $result.Line -ne $null)
					{
						# Lets parse out the crap html and get our GUID that we needed
						$ActivityID = $result.Line -replace '.*<td class="alt">UniqueID<\/td><td>', $null
						$ActivityID = $ActivityID -replace '.*<td class="alt">UniqueID<\/td><td colspan="3">', $null
						$ActivityID = $ActivityID -replace '<\/td><\/tr>.*$', $null
						$ActivityID = $ActivityID.Trim()
						
						#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:  Found this activity ID: $($ActivityID)"
						#Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:  Current Runbook: $($thisRunbook.Name)"
						Log-Message -logMsg "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:  Obtaining this activity object from the OIS Export Data..."
						$objThisActivity = Get-ObjectByID -objectID "$($ActivityID)" -runbookObject $currentPolicy
						
						$fileLinkName = $objThisActivity.Name.'#text';
						$fileLinkID = $objThisActivity.UniqueID.'#text';
						
						
						$source = Get-ObjectByID -objectID $objThisActivity.SourceObject.'#text' -runbookObject $currentPolicy
						$target = Get-ObjectByID -objectID $objThisActivity.TargetObject.'#text' -runbookObject $currentPolicy
						
						if ($source -ne $null -and $target -ne $null)
						{
							#Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:  Searching directory: $($ReportExportPath)$($SubRunbookFolder) for Link ID as it's uniqueID: $($fileLinkID)..."
							$thisLinkURL = Find-ObjectInHTMLFiles -objectID $fileLinkID -rootDirectory "$($ReportExportPath)$($SubRunbookFolder)" 
							
							# Find the source and target objects
							#Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: Searching directory: $($ReportExportPath)$($SubRunbookFolder) for Source ID as it's uniqueID: $($source.UniqueID.'#text')..."
							$srcURL = Find-ObjectInHTMLFiles -objectID $source.UniqueID.'#text' -rootDirectory "$($ReportExportPath)$($SubRunbookFolder)"
							
							#Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: Searching directory: $($ReportExportPath)$($SubRunbookFolder) for Target ID as it's uniqueID: $($target.UniqueID.'#text')..."
							$targetURL = Find-ObjectInHTMLFiles -objectID $target.UniqueID.'#text' -rootDirectory "$($ReportExportPath)$($SubRunbookFolder)"
							
						}
						
						if ($thisLinkURL -eq $null -and $srcURL -eq $null -and $targetURL -eq $null -and $source -eq $null -and $target -eq $null)
						{
							Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:[ WARNING ]: One of the pieces of data required were missing!!!"
							Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:[ WARNING ]: thisLinkURL : $($thisLinkURL)"
							Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:[ WARNING ]: srcURL : $($srcURL)"
							Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:[ WARNING ]: targetURL : $($targetURL)"
							Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:[ WARNING ]: source : $source"
							Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]:[ WARNING ]: target : $target"
							Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: Getting Link Activity HTML - File: $($thisLinkURL)"
							
						}
						else
						{
							$thisLinkContent = Get-Content "$($thisActivityFilePath)"
							# Update the link with the correct urls for the user's file
							$areaMapHTML = Generate-LinkCoords -source $source -target $target -sourceFile $srcURL -targetFile $targetURL
							
							#Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: The following HTML was generated:"
							#Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: $($areaMapHTML)"
							
							$thisLinkContent = $thisLinkContent.Replace('%_PLACEHOLDER_ACTIVITY_COORDS_%', $areaMapHTML)
							Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: HTML updated - Re-writing file to disk...."
							
							$thisLinkContent | Out-File -FilePath $thisActivityFilePath -Force -Confirm:$false -Encoding UTF8
							Log-Message "[ MAIN SCRIPT BLOCK ]:[ POST-CREATION PROCESSING ]: $($thisLinkURL) has been updated with activity coordinates" 
							Log-Message "[ POST-CREATION PROCESSING COMPLETE ]: $($thisLinkURL)"
						}
						
						#endregion
						
					}
					
				}
				
			}
			# End for each Activity File found (Process CSS for 'Link' Activities)
			
			Write-Host "[ POST-CREATION PROCESSING ]: Runbook Report Generation Process Complete!" -ForegroundColor Yellow -BackgroundColor:Blue
			
		}

}
#endregion

#region Main Script

# available icons can be reviewed/edited by looking at the Get-IconForActivity function at the top of the script
$script:IconRootFolder = $PictureRepository.Replace('runbook_images', 'icons\activities')
$script:arrIcons = GCI $IconRootFolder | Sort-Object $ToNatural

# Get your Exported SCOrch Integration Pack as XML Content (this is the source for all our data from this point forward)
$script:oisExport = [xml](Get-Content $Path)

#$rootRunbookFolderName = $oisExport.ExportData.Policies.Folder.Name.Replace(' ',$null)
$today = (Get-Date -f 'MM-dd-yy')
$rootRunbookFolderName = "ExportedRunbook_$($today)"

# Get each of the Activity Nodes
$nodeFolder = @()
#$masterRunbook = $oisExport.ExportData.Policies.Folder

# Add the root master runbook folder
$oisExport.ExportData.Policies.Folder | % { $nodeFolder += $_ }

# For each subfolder, add this object into our nodeFolder array for processing
$oisExport.ExportData.Policies.Folder.Folder | % { $nodeFolder += $_ }

# Get all our global variables also so we can search those for names of things later
$script:objGlobalRunbookVarFolders = $oisExport.ExportData.GlobalSettings.Variables.Folder

# Create a folder in the reports section for the root runbook
if(!(test-path "$($ReportExportPath)$($rootRunbookFolderName)\")) {
	
	# We need to create the new root folder for our runbook exports
	try {
		
		New-Item -ItemType Directory -Path "$($ReportExportPath)$($rootRunbookFolderName)\" -Force
	
		# Now that we have our directory there, lets reset the global report path to this new folder
		$script:ReportExportPath = "$($ReportExportPath)$($rootRunbookFolderName)\"
	
	} catch {
		
		# Log what info we could and exit script - without this folder the rest doesnt matter
		Log-Message "[ ERROR ]: There was an error creating the directory: $($ReportExportPath)$($rootRunbookFolderName)"
		Log-Message "[ ERROR ]: The error was:`r$($Error[0])"
		Write-Host "***** ERROR *******" -BackgroundColor Yellow -ForegroundColor Black 
		Write-Host "$($Error[0])" 
		Write-Host "Closing script in 20 seconds.  Copy what error info you want to or check the log.."
		Start-Sleep -seconds 20
		Stop-Process $PID
	}
	

}

# For Each Runbook in the array - Obtain the following Data
$nodeFolder | % {
	
	# Get the root runbook folder folder 
	$thisRunbookFolder = $_;
	$AllRunbooksinFolder = @()
	$AllRunbooksinFolder += $thisRunbookFolder.Policy
	
	# For each Policy (subFolder) we have
	$AllRunbooksInFolder | % {
		
		# Assign the current policy a variable
		$currentPolicy = $_;
		
		# Get the sub folder if any within this runbook folder
		$SubRunbookFolder = $currentPolicy.Name.'#text';
		$SubRunbookFolder = $SubRunbookFolder.Replace(' ',$null);
		
		# Now pass the variables to our main function that processes everything
		Process-RunbookFolders -runbookPolicy $currentPolicy -runbookFolderName $SubRunbookFolder

		# Get the current runbooks name and turn it into what it's corresponding report folder should be
		#$runbookFolderName = $runbookExportFolder
		Write-Host "[ POST-CREATION PROCESSING ]: Checking for HTML files in: $($ReportExportPath)$($SubRunbookFolder)..." -ForegroundColor White
		Log-Message -logMsg "[ POST-CREATION PROCESSING ]: Checking for HTML files in: $($ReportExportPath)$($SubRunbookFolder)..."
	
		Process-RunbookLinks -runbookFolder "$($ReportExportPath)$($SubRunbookFolder)" -currentPolicy $currentPolicy
		
	}
	
}
# end for each nodeFolder

Write-Host "----------------------------------------------------------------------------------------------------------------------" -ForegroundColor Blue -BackgroundColor White
Write-Host "" -ForegroundColor Blue -BackgroundColor White
Write-Host "[ OPERATION COMPLETE ]: DOCUMENTATION FOR RUNBOOK HAS BEEN GENERATED! CONGRATS! LOOK AT ALL THAT TIME YOU SAVED!!! ;)" -ForegroundColor Blue -BackgroundColor White
Write-Host "" -ForegroundColor Blue -BackgroundColor White
Write-Host "----------------------------------------------------------------------------------------------------------------------" -ForegroundColor Blue -BackgroundColor White
Write-Host ""
Write-Host "Beginning Final Phase - Make GUIDs become understandable object names..."
start-sleep -seconds 2
# ---------------------------------------------------------------------------
# 
# I found one issue towards the end, when I realized one piece of the documentation
# that would be confusing would be wherever the object ID (GUIDs) were specified for
# subscribed or global variables, or objects referenced in the database.
# we will loop through each activity HTML file and scrub the data, and replace our
# GUIDs with meaningful data
#
# ---------------------------------------------------------------------------
$nodeFolder | % {
# Get the root runbook folder folder 
	$thisRunbookFolder = $_;
	
	# For each Policy (subFolder) we have
	$thisRunbookFolder.Policy | % {

		# Get the current runbook
		$thisRunbookObject = $_		
		
		$RunbookName = $thisRunbookObject.Name.'#text'.ToString().Replace(' ',$null) 
		$runbookReportPath = "$($ReportExportPath)$($RunbookName)"

		Log-Message "[ UPDATE GUIDS WITH HUMAN SPEAK ]: Begin processing runbook $($RunbookName)"
		Log-Message "[ UPDATE GUIDS WITH HUMAN SPEAK ]: Searching Folder $($runbookReportPath)..."
		
		# Now loop through each folder in our export directory, find the NON Table of Contents files, and update the GUIDs with human readable information
		$runbookActivityFiles = Get-ChildItem -path $runbookReportPath | ? { $_.PSIsContainer -eq $false -and $_.Name -ne "Index.html" }
		
		# for each activity file
		for($t=0;$t -le ($runbookActivityFiles.Count - 1);$t++) {
		
			# Here is your activity file info
			$thisActivityFilePath = $runbookActivityFiles[$t].FullName;
			$thisActivityFileName = $runbookActivityFiles[$t].Name;
			Write-Host "Updating GUID References on: " -foregroundcolor yellow -nonewline
			Write-Host "$($thisActivityFileName)" -foregroundcolor White
			Write-Host "---"
		
			# Lets process this file and fix the stuff we need clarification for
			Update-GUIDsWithHumanSpeak -file $thisActivityFilePath -runbook $thisRunbookObject
		}
	
	}
	
}

Write-Host "Updating GUID to Reader Friendly Names Complete! " -foregroundcolor yellow


# ---------------------------------------------------------------------------
#
# 	LAST STEP IN REPORT: "MASSAGING OF DATA"
# 	This step includes the provisioning of the left side menu
# 	on each table of contents page, which should have all been already created.
#
# ---------------------------------------------------------------------------
# Now loop through each folder in our export directory, find the Index.html file, and update it with a menu on the left hand side
$runbookToCFiles = Get-ChildItem -path "$($script:ReportExportPath)" -recurse | ? { $_.PSIsContainer -eq $false -and $_.Name -like "Index.html" }

# Get a list of all runbooks in this grouping, and generate some HTML code to inject into the Index.html files for each runbook
$ToCMenuCSS = Generate-RunbookMenuSelections -ReportExportPath "$($script:ReportExportPath)" -oisExportXML $oisExport

# Update Each Table of contents file with the data we have created
Process-TableOfContentFiles $runbookTocFiles $ToCMenuCSS	

Write-Host ""
Write-Host "***********************************************************" -foregroundcolor yellow
Write-Host "Script Execution Complete! Reports are ready for viewing!" -foregroundcolor White
Write-Host "***********************************************************" -foregroundcolor yellow
Write-Host ""
Log-Message "**********************************************************" 
Log-Message "Script Execution Complete! Reports are ready for viewing!"
Log-Message "**********************************************************" 

# If running in the console, wait for input before closing. 
if ($Host.Name -eq "ConsoleHost") {     
	Write-Host "Press any key to continue..."    
	$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null
}
#endregion
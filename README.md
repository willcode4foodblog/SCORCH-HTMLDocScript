  .NOTES
	===========================================================================
	OIS Export Parser - HTML Document Generator v0.60
	===========================================================================
	 Created by:   	Michael Adams <michael_adams@outlook.com>
	 Organization: 	http://willcode4foodblog.wordpress.com
	 Created on:   	8/18/2014 7:01 PM
	 Last Updated:	10/25/2014 2:00 PM 
	 
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
		

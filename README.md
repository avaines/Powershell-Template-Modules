Module Repository
===================
Here is a repository of pre-built modules for the script template.
Add the required PS1 file to the "Modules" folder of your projects copy of the "New-ScriptTemplate"

------------------

Connect-AD
-------------
**More details are available here: [Blog Post](http://vaines.org/)**
Connect-AD allows you to well connect to AD
Add '*connect-ad*' to your script in the "Main Script Block" of the template.
The script will check if the module is loaded and import the ActiveDirectory module if not

> **Note:**
 RSAT needs to be installed and configured, this module will abort the script if not installed


Get-ADTools
-------------
**More details are available here: [Blog Post](http://vaines.org/)**
Common AD manipulating tools
> **Note:**
> If this module is used "*Connect-AD*" should also be included in this project

**Get-NestedGroupmember**
Returns all members including members of nested AD groups


Connect-O365
-------------
**More details are available here: [Blog Post](http://vaines.org/)**
Contains a suite of functions for connecting to office 365 powershell services

Connect to individual services:
	
    Connect-SFBOnline
    	write-host "Do Some Skype For Business stuff"
    	Disconnect-SFBOnline
    Connect-ExOnline
    	write-host "Do Some Exchange stuff"
    Disconnect-ExOnline
    Connect-SCCOnline
    	write-host "Do Some Stuff with Security & Compliance"
    Disconnect-SCCOnline	

or connect to all of them

    Connect-Office365
    	write-host "Do Some Stuff with Exchange, Skype and Security & Compliance"
    Disconnect-Office365


> **Note:**
> The Skype for Business/Lync module seems to only work if run as administrator, 
> The user context is checked and the script will abort if its not got admin rights


Class-HTMLReport
----------------
A simple HTML report generator
**More details are available here: [Blog Post](http://vaines.org/)**

Include the class somewhere in your project, either as code or as a dot-sourced module

        #create some content
		$proc = Get-Process | Sort-Object CPU -desc | Select-Object CPU, ProcessName, ID -first 10
        $files = Get-ChildItem | select mode,LastWriteTime,Length,Name
        $TextBlock = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Quisque et nisl laoreet, malesuada urna a, semper risus."
		
		
		#create a new HTMLReportObject
		$MyReport = New-Object HtmlReport
				
		#Give the new report a output location, logo image and a title
		$MyReport.Outputpath = ".\report.html" 
        $MyReport.LogoPath = ".\logo.png"
        $MyReport.Title = "My Report"
        
		#Add a text block (title, contents, width of block[%], colour [blue/red/green/yellow])
        $MyReport.addTextBlock("Free Text Block", $TextBlock, "30%" ,"red")		
		
		#Add a table
        $MyReport.addTableBlock("Files", $files,"30%","blue")
		
		#Add an image 
        $MyReport.addImageBlock("Image", ".\graph.png", "20%" , "yellow")
		
           
        $MyReport.publish()

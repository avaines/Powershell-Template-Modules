Class-HTMLReport
===================
A simple HTML report generator
**More details are available here: [Blog Post](http://vaines.org/)**


Example usage
------------
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


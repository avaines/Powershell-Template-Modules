Class HtmlReport{
    [String]$LogoPath
    [String]$Title = "MyReport"
    $Sections=@()
    $Outputpath = ".\report.html" 
    $content
    $htmlreport
    
             
    publish () {         
        $date = ( get-date ).ToString('yyyy/MM/dd')  
        $this.htmlreport += "<html>
            <head>
                <meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
                <style>
                    h1 {font-family:Tahoma;font-size:14pt; color: #313131;}
                    body p {font-family:Tahoma;font-size:10pt; color: #0C0B07;}
                    h2{width:95%;font-family:Tahoma;padding-left:10px;padding-top:2px;padding-bottom:2px;font-size:12pt;color: #0C0B07;}
                    .headingParent{width:100%;}
                    .headingTitle{float:left;padding-left: 2%;top: 50%;transform:translateY(50%);}
                    #sectionparent{width:100%;position: absolute;}
                    #section{position: relative;float:left;}
                    #footer p{font-family:Tahoma;font-size:8pt; font-style: italic;color: #0C0B07;}
                    #footer {width:100%;float:left;}
                    table {font-family: Tahoma, Geneva, sans-serif;table-layout: auto;text-align: center;border-collapse: collapse;border: 1px solid #FFFFFF;}
                    table td, table th {border: 1px solid #FFFFFF;padding: 3px 2px;}
                    table tbody td {font-size: 12px;}
                    table thead th {font-size: 14px;color: #FFFFFF;text-align: center;border-left: 2px solid #FFFFFF;}
                    table thead th:first-child {border-left: none;}
                    table tfoot td {font-size: 12px;}
                    table thead {border-bottom: 5px solid #FFFFFF;}
                    div.Blue h2{background-color:#442E84;color:#f2f2f2;}
                    div.Blue table tr:nth-child(even) {background: #C6BAE8;}
                    div.Blue table thead {background: #5B4699;}
                    div.Yellow h2{background-color:#B0AC00;color:#f2f2f2;}
                    div.Yellow table tr:nth-child(even) {background: #FFFC6F;}
                    div.Yellow table thead {background: #E2DE01;text-shadow: -1px -1px 0 #B0AC00,1px -1px 0 #B0AC00,-1px 1px 0 #B0AC00, 1px 1px 0 #B0AC00;}
                    div.Green h2{background-color:#1A7C15;color:#f2f2f2;}
                    div.Green table tr:nth-child(even) {background: #BFF1BC;}div.Green table thead {background: #4BB446;}
                    div.Red h2{background-color:#BE3337;color:#f2f2f2;}
                    div.Red table tr:nth-child(even) {background: #FEC6C7;}
                    div.Red table thead {background: #DD5659;}
                </style>"

         $this.htmlreport += ("<title>" + $this.Title + "</title>")
         $this.htmlreport += ("<body>
             <div id='headingParent'>
	            <div'>
		            <img src='" + $this.LogoPath + "'/>
	            </div>
	            <div id='headingTitle'>
		            <h1>"+$this.Title+"</h1>
	            </div>
            </div>
            <div class='sectionParent'>") 
   
        $this.htmlreport += $this.content

        $this.htmlreport += "</div>"


        $this.htmlreport += ("</div><div id='footer'>
	        <br><hr><p>Created "+$(get-Date) +" on " + $env:ComputerName + " by "+ $env:UserName +"</p>
            </div></Body></html>")



       $this.htmlreport | out-file $this.Outputpath

    } #Publish


    addTableBlock($Title,$data,$width,$colour){
        #start a new "section" with the specified width, colour and title
        $this.content += ("<div id='section' class='"+$colour+ "' style='width:" + $width + ";' >
	        <h2>"+ $Title+"</h2>")


        #Generate the table tags
        $this.content += "<table><thead><tr>"
        
        #we need a <th> tag for each heading in the data provided
        $ObjectHeadings = $data[0].psobject.Properties.name 
        
        #for each heading in the data object, create a <th> tag
        foreach($heading in $ObjectHeadings){
            $this.content += ("<th>" + $heading + "</th>")
        }
 
        #end the row
        $this.content += "</tr></thead><tbody>"


        #create the table rows with the correct data
        foreach($row in $data){
            $this.content += "<tr>"
            
            #for each column, in the current row, create a 'cell'
            foreach($objectHeading in $ObjectHeadings){    
                $this.content += ("<td>"+ ($row).$objectHeading + "</td>"  )
            }

            #end of the current row
            $this.content +=  "</tr>"

        }
        $this.content +=  "</tbody></table></div>" 
        
          
    }
    

    addTextBlock($Title,$text,$width,$colour){
        #start a new "section" with the specified width, colour and title
        $this.content += ("<div id='section' class='"+$colour+ "' style='width:" + $width + ";' >
	        <h2>"+ $Title+"</h2>")


        #Generate the table tags
        $this.content += ("<p>" +$text + "</p>")
        $this.content += "</div>"
    }

    
    addImageBlock($Title,$image,$width,$colour){
        #start a new "section" with the specified width, colour and title
        $this.content += ("<div id='section' class='"+$colour+ "' style='width:" + $width + ";' >
	        <h2>"+ $Title+"</h2>")


        #Generate the table tags
        $this.content += ("<img src='" + $image + "'width='100%'/>")
        $this.content += "</div>"
    }

}
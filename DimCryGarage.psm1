# MIT License
# 
# Copyright (c) 2020 DimCry's Garage (Cristian Dimofte)
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# Description:
    ### This module contains common useful functions

# Author:
    ### DimCry (Cristian Dimofte)

#####################################
# Common space for Global variables #
#####################################
#region "Global variables"

### TheWorkingDirectory (Global: Script) variable is used to list the exact value for the Working Directory
$Global:TheWorkingDirectory = $null

[System.Collections.ArrayList]$Global:LogObject = @()
[string]$Global:LogPath = "$Env:Temp\DimCryGarage\Log.log"
[string]$Global:HTMLLogPath = "$Env:Temp\DimCryGarage\Log.html"
[string]$GLobal:LogEntry = $null
    $Global:HTMLLogEntry = New-Object PSObject
        $Global:HTMLLogEntry | Add-Member -NotePropertyName Date -TypeName string -NotePropertyValue ""
        $Global:HTMLLogEntry | Add-Member -NotePropertyName Type -TypeName string -NotePropertyValue ""
        $Global:HTMLLogEntry | Add-Member -NotePropertyName Description -TypeName string -NotePropertyValue ""

#endregion "Global variables"


#####################################
# Common space for Functions #
#####################################
#region "Functions"

<#
.SYNOPSIS
    Function to export content from a String or TableStructure to HTML report

.DESCRIPTION
    Function "Export-ReportToHTML" is used to export content from a String or TableStructure to HTML report.
    It can create the HTML report using different themes (for the moment just the "Blue" / default one is defined).

.EXAMPLE
    Export-ReportToHTML -FilePath $FilePath -PageTitle $PageTitle -ReportTitle $ReportTitle -TheObjectToConvertToHTML $TheObjectToConvertToHTML

	Description
	-----------
	This example is exporting the information into the $FilePath HTML report

.NOTES
    This function can be used to convert any correctly formatted PowerShell content to an HTML report.
#>
function Export-ReportToHTML {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,

        [Parameter(Mandatory=$false)]
        [string]$PageTitle,

        [Parameter(Mandatory=$false)]
        [string]$ReportTitle,

        [Parameter(Mandatory=$true)]
        [System.Collections.ArrayList]$TheObjectToConvertToHTML,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Blue")]
        [string]$Theme = "Blue"        
    )

### Create header of the HTML file
if ($Theme -eq "Blue") {
$HTMLBeginning = @"
<!DOCTYPE html>
<html>
	<head>
	
		<meta charset="UTF-8">
		<meta name="description" content="DimCry's Garage">
		<meta name="keywords" content="PowerShell, HTML, Report, DimCry, Garage">
		<meta name="author" content="Cristian Dimofte">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
		
        <title>$PageTitle</title>
		<link rel="titlebar icon" href="data:image/png;base64,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" />
		
		<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css" />
		<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css" />
		<script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>
		
		<style type="text/css">
			html {
				box-sizing: border-box;
				font-family: FabricMDL2Icons;
			}
			
			*, *:before, *:after {
				box-sizing: inherit;
			}			
			
			.ms-font-su {
				font-family:FabricMDL2Icons;
				-webkit-font-smoothing:antialiased;
				font-weight:100
			}
			
			.ms-fontColor-themePrimary,.ms-fontColor-themePrimary--hover:hover{
				color:#0078d7
			}
			
			@font-face{
				font-family:FabricMDL2Icons;
				src:url(https://spoprod-a.akamaihd.net/files/fabric/assets/icons/fabricmdl2icons.woff) format('woff'),url(https://spoprod-a.akamaihd.net/files/fabric/assets/icons/fabricmdl2icons.ttf) format('truetype');
				font-weight:400;
				font-style:normal
			}

			body {
				padding: 10px;
				font-family: FabricMDL2Icons;
				background: #f6f6f6;
			}
			
			a {
				color: #06c;
			}

			h1 {
				margin: 0 0 1.5em;
				font-weight: 600;
				font-size: 1.5em;
			}

			h2 {
				color: #1a0e0e;
				font-size: 20px;
				font-weight: 700;
				margin: 0;
				line-height: normal;
				text-transform:uppercase;
				margin-block-start: 0.3em;
				margin-block-end: 0.3em;
				text-indent: 0px;
				margin-left: 0px;
			}
    
			h2 span {
				display: block;
				padding: 0;
				font-size: 18px;
				opacity: 0.7;
				margin-top: 5px;
				text-transform:uppercase;
			}

			h3 {
				color: #000000;
				font-size: 17px;
				font-weight: 300;
				margin-left: 10px;
				font-family: FabricMDL2Icons;
			}
    
			h3 span {
				color: #000000;
				font-size: 17px;
				font-weight: 300;
				margin-left: 10px;
				font-family: FabricMDL2Icons;
			}			
			
			h5 {
				display: block;
				padding: 0;
				margin: 5px 0px 0px 0px;
				font-size: 42px;
				font-weight: 100;
				font-family: FabricMDL2Icons;
			}
			
			p {
				margin: 0 0 1em;
				padding: 10px;
				font-family: FabricMDL2Icons;
			}

			.ms-icon--description:before {
				font-family: FabricMDL2Icons;
				content: "\EA16";
				color: #0078d7;
				margin-right: 5px;
				font-style: normal;
				font-size: 20px;
			}

			.ms-icon--emojineutral:before {
				font-family: FabricMDL2Icons;
				content: "\EA87";
				margin-left: 25px;
				margin-right: -15px;
				font-style: normal;
			}

			.ms-icon--emojihappy:before {
				font-family: FabricMDL2Icons;
				content: "\E76E";
				margin-left: 25px;
				margin-right: -15px;
				font-style: normal;
			}

			.ms-icon--emojidisappointed:before {
				font-family: FabricMDL2Icons;
				content: "\EA88";
				margin-left: 25px;
				margin-right: -15px;
				font-style: normal;
			}

			.label {
				width: 100%;
				height: 80px;
				background-color: #0078d7;
				color: #ffffff;
				font-size: 46px;
				display: inline-block;
			}
			
			.label1 {
				width: 100%;
				height: 80px;
				background-color: #f6f6f6;
				color: #0078d7;
				font-size: 46px;
				display: inline-block;
				box-sizing: border-box;
			}
			
			.body-panel {
				font-size: 14px;
				padding: 15px;
				margin-bottom: 5px;
				margin-top: 5px;
				box-shadow: 0 1px 2px 0 rgba(0,0,0,.1);
				overflow-x: auto;
				box-sizing: border-box;
				background-color: #fff;
				color: #333;
				-webkit-tap-highlight-color: rgba(0,0,0,0);
				padding-top: 15px;
				font-family: FabricMDL2Icons;
			}

			.accordion {
				margin-bottom: 1em;
			}

			.accordion p:last-child {
				margin-bottom: 0;
			}
			
			.accordion > input[name="collapse"] {
				display: none;
			}
			
			.accordion > input[name="collapse"]:checked ~ .content {
				height: auto;
				transition: height 0.5s;
			}
			
			.accordion label, .accordion .content {
				max-width: 3200px;
				width: 99%;
			}
			
			.accordion .content {
				background: #fff;
				overflow: hidden;
				overflow-x: auto;
				height: 0;
				transition: 0.5s;
				box-shadow: 1px 2px 4px rgba(0, 0, 0, 0.3);
			}
			
			.accordion label {
				display: block;
			}
			
			.accordion > input[name="collapse"]:checked ~ .content {
				border-top: 0;
				transition: 0.3s;
			}
			
			.accordion .handle {
				margin: 0;
				font-size: 16px;
			}
			
			.accordion label {
				color: #0078d7;
				cursor: pointer;
				font-weight: 300;
				padding: 10px;
				background: #e6f3ff;
				user-select: none;
			}
			
			.accordion label:hover, .accordion label:focus {
				background: #cce6ff;
				color: #0000ff;
			}
			
			.accordion .handle label:before {
				font-family: FabricMDL2Icons;
				content: "\E972";
				display: inline-block;
				margin-right: -20px;
				font-size: 1em;
				line-height: 1.556em;
				vertical-align: middle;
				transition: 0.4s;
			}
			
			.accordion > input[name="collapse"]:checked ~ .handle label:before {
				transform: rotate(180deg);
				transform-origin: center;
				transition: 0.4s;
			}
			
			section {
				float: left;
				width: 100%;
			}
    
			.container{
				max-width: 3200px;
				width:99%;
				margin: 0 auto;
			}

		   table {
				font-size: 14px;
				font-weight: 800;
				border: 1px solid #0078d7;
				border-collapse: collapse;
				font-family: FabricMDL2Icons;
				margin-left: 10px;
				margin-bottom: 10px;
				margin-right: 10px;
				text-align: center;
			} 
			
			td {
				padding: 4px;
				margin: 0px;
				border: 1px solid #0078d7;
				border-collapse: collapse;
			}
			
			th {
				color: #0078d7;
				font-family: FabricMDL2Icons;
				text-transform: uppercase;
				padding: 10px 5px;
				vertical-align: middle;
				border: 1px solid #0078d7;
				border-collapse: collapse;
			}
			
			tbody tr:nth-child(odd) {
				background: #f0f0f2;
			}			

			#output {
				color: #004377;
				font-size: 14px;
			}
			
			@media screen and (max-width:639px){
				h2 {
					color: #1a0e0e;
					font-size: 16px;
					font-weight: 700;
					margin: 0;
					line-height: normal;
					text-transform:uppercase;
					text-indent: 0px;
					margin-left: 0px;
				}
			
				h5 {
					margin: 5px 0px 0px 0px;
					font-size: 8vw;
					font-weight: 100;
					font-family: FabricMDL2Icons;
				}
				
				.accordion label, .accordion .content {
					max-width: 3200px;
					width: 99%;
				}
				
				.accordion .content {
					background: #fff;
					overflow: hidden;
					overflow-x: auto;
					height: 0;
					transition: 0.5s;
					box-shadow: 1px 2px 4px rgba(0, 0, 0, 0.3);
				}
				
				.accordion > input[name="collapse"]:checked ~ .content {
					height: auto;
				}
			}
			
			@media (max-device-width:480px) and (orientation:landscape) {
				.body-panel {
					max-height:200px
				}
			}

			/* For Desktop */
			@media only screen and (min-width: 620px) {
				.accordion > input[name="collapse"]:checked ~ .content {
					height: auto;
				}
			}	
			
			.ms-icon--HTMLReport:before {
				font-family: FabricMDL2Icons;
				content: "\E9D9";
				font-size: 120px;
				color: #0078d7;
				font-style: normal;
			}
		
		</style>
		
	</head>
	<body class="ms-Fabric" dir="ltr">
		<header>
		<div class="label1">
			<a href="https://github.com/dimcry/DimCry-Garage" target="_blank">
				<img src="data:image/png;base64,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" alt="DimCry's Garage" width="auto" height="110%" style="float: left; margin-right: 5px">
			</a>
			<div style="width: 100%; height: 100%; ">
				<h5 class="ms-font-su ms-fontColor-themePrimary" style="font-weight: 350; ">
                    $ReportTitle
                </h5>
            </div>
        </div>
        </header>
        <div class="body-panel" style="margin: 30px 0px 0px 0px; display: block;">
"@

$HTMLEnd = @"
		</div>
		
		<div class="ms-font-su ms-fontColor-themePrimary" style="font-size: 15px; margin-left: 10px; text-align: right;">
			<ul>Creation Date: $((Get-date).ToUniversalTime()) UTC</ul>
			<ul>&copy; 2020 DimCry's Garage</ul>
		</div>
	</body>
</html>
"@
}

    [int]$i = 1
    [string]$TheBody = $HTMLBeginning

    ### For each scenario, convert the data to HTML
    foreach ($Entry in $TheObjectToConvertToHTML) {
		
		[string]$Emoji = $null
		if ($Entry.SectionTitleColor -eq "Green") {
			$Emoji = "happy"
		}
		elseif ($Entry.SectionTitleColor -eq "Red") {
			$Emoji = "disappointed"
		}
		else {
			$Emoji = "neutral"
		}

        $TheBody = $TheBody + "
        	`<section class=`"accordion`"`>
				`<input type=`"checkbox`" name=`"collapse`" id=`"handle$i`" `>
				`<h2 class=`"handle`"`>
					`<label for=`"handle$i`" `>
						`<h2 class=`"ms-font-su ms-fontColor-themePrimary`" style=`"display: inline-block; color: $($Entry.SectionTitleColor); font-size: 20px; font-weight: 650;`"`>`<i class=`"ms-icon ms-icon--emoji$Emoji`"`>`</i`>&nbsp;&nbsp;&nbsp;&nbsp;$($Entry.SectionTitle)`</h2`>
					`</label`>
				`</h2`>
				`<div class=`"content`"`>
					`<h3`>`<i class=`"ms-icon ms-icon--description`"`>`<`/i`>$($Entry.Description)`</h3`>
        "

        if ($Entry.DataType -eq "String") {
            $TheValue = "					`<p style=`"font-family: FabricMDL2Icons; font-weight: 800; margin-left: 10px;`"`>$($Entry.EffectiveData)`<`/p>"
        }
        else {
            $TheProperties = ($($Entry.EffectiveData)| Get-Member -MemberType NoteProperty).Name
            $TheValue = $($Entry.EffectiveData) | ConvertTo-Html -As $($Entry.TableType) -Property $TheProperties -Fragment
            
            if ($Entry.TableType -eq "List") {
                [int]$z = 0
                foreach ($NewEntryFound in $TheValue) {
                    if ($NewEntryFound -like "*<table>*") {
                        $TheValue.Item($z) = $NewEntryFound.Replace("<table>", "<table style=`"text-align: left;`">")
                    }
                    elseif ($NewEntryFound -like "*<tr><td>*") {
                        $TheValue.Item($z) = $NewEntryFound.Replace("<tr><td>", "<tr><th>")
                        $TheValue.Item($z) = (($TheValue.Item($z) -split "<td")[0] -replace "`/td>", "`/th>") + ($TheValue.Item($z).Substring((($TheValue.Item($z) -split "<td")[0].Length),($TheValue.Item($z).Length-($TheValue.Item($z) -split "<td")[0].Length)))
                    }
                    $z++
                }
            }
        }

        ### Adding sections in the body of the HTML report
        $TheBody = $TheBody + $TheValue
        $TheBody = $TheBody + "
        	`<`/div`>
		`<`/section`>
        "

        $i++
    }
    $TheBody = $TheBody + $HTMLEnd
    $TheBody | Out-File $FilePath -Force
}



<#
.SYNOPSIS
    "Prepare-ObjectForHTMLReport" function is used to prepare the objects to be converted to HTML file

.DESCRIPTION
    "Prepare-ObjectForHTMLReport" function is used to prepare the objects to be converted to HTML file

.EXAMPLE
	[PSCustomObject]$TheCommand = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description -DataType "String" -EffectiveDataString $TheString
	
	Description
	-----------
	This example is preparing an object that has a string as DataType, to be listed into the HTML report

.EXAMPLE	
	[PSCustomObject]$TheCommand = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description -DataType "TableStructure" -EffectiveDataTableStructure $TheTable -TableType "List"

	Description
	-----------
	This example is preparing an object that has a TableStructure as DataType, to be listed as a List into the HTML report
	
.EXAMPLE	
	[PSCustomObject]$TheCommand = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description -DataType "TableStructure" -EffectiveDataTableStructure $TheTable -TableType "Table"

	Description
	-----------
	This example is preparing an object that has a TableStructure as DataType, to be listed as a Table into the HTML report

.NOTES
    This function can be used to correctly prepare any PowerShell content to be exported to an HTML report
#>
function Prepare-ObjectForHTMLReport {
param (
    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [Parameter(ParameterSetName = "TableStructure", Mandatory=$false)]
    [string]$SectionTitle,

    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [Parameter(ParameterSetName = "TableStructure", Mandatory=$false)]
    [ValidateSet("Black", "Green", "Red")]
    [ConsoleColor]$SectionTitleColor,

    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [Parameter(ParameterSetName = "TableStructure", Mandatory=$false)]
    [string]$Description,

    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [Parameter(ParameterSetName = "TableStructure", Mandatory=$false)]
    [ValidateSet("TableStructure", "String")]
    [string]$DataType,

    [Parameter(ParameterSetName = "String", Mandatory=$false)]
    [string]$EffectiveDataString,

    [Parameter(ParameterSetName = "TableStructure", Mandatory=$false)]
    [PSCustomObject]$EffectiveDataTableStructure,

    [Parameter(ParameterSetName = "TableStructure", Mandatory=$false)]
    [ValidateSet("List", "Table")]
    [string]$TableType
)

    ###Create the object, with all needed Properties, that will be used to convert into an HTML report
    $TheObject = New-Object PSObject
        $TheObject | Add-Member -NotePropertyName SectionTitle -NotePropertyValue $SectionTitle
        $TheObject | Add-Member -NotePropertyName SectionTitleColor -NotePropertyValue $SectionTitleColor
        $TheObject | Add-Member -NotePropertyName Description -NotePropertyValue $Description
        $TheObject | Add-Member -NotePropertyName DataType -NotePropertyValue $DataType
        if ($DataType -eq "TableStructure") {
            $TheObject | Add-Member -NotePropertyName EffectiveData -NotePropertyValue $EffectiveDataTableStructure
            $TheObject | Add-Member -NotePropertyName TableType -NotePropertyValue $TableType
        }
        else {
            $TheObject | Add-Member -NotePropertyName EffectiveData -NotePropertyValue $EffectiveDataString
        }

    ### Return the created object
    return $TheObject

}

function Test-ExportReportToHTMLUsingDummyData {
	
	[System.Collections.ArrayList]$TheObjectToConvertToHTML = @()

	### "String1" section
	[string]$SectionTitle = "This is the title of the `"String1`" section"
	[ConsoleColor]$SectionTitleColor = "Green"
	[string]$Description = "This is the description of the `"String1`" section"
	[string]$TheString = "This is the content of the `"String1`" section"
		[PSCustomObject]$TheCommand = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description -DataType "String" -EffectiveDataString $TheString
	$null = $TheObjectToConvertToHTML.Add($TheCommand)

	### "List" section
	[string]$SectionTitle = "This is the title of the `"List`" section"
	[ConsoleColor]$SectionTitleColor = "Black"
	[string]$Description = "This is the description of the `"List`" section"

		[System.Collections.ArrayList]$TheTable = @()
		$AList = New-Object PSObject
			$AList | Add-Member -NotePropertyName FirstRow -NotePropertyValue "A value of the first row"
			$AList | Add-Member -NotePropertyName SecondRow -NotePropertyValue "A value of the second row"
			$AList | Add-Member -NotePropertyName ThirdRow -NotePropertyValue "A value of the third row"
			$AList | Add-Member -NotePropertyName FourthRow -NotePropertyValue "A value of the fourth row"
			$AList | Add-Member -NotePropertyName FifthRow -NotePropertyValue "A value of the fifth row"
		$null = $TheTable.Add($AList)

	[PSCustomObject]$TheCommand = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description -DataType "TableStructure" -EffectiveDataTableStructure $TheTable -TableType "List"
	$null = $TheObjectToConvertToHTML.Add($TheCommand)

	### "Table" section
	[string]$SectionTitle = "This is the title of the `"Table`" section"
	[ConsoleColor]$SectionTitleColor = "Red"
	[string]$Description = "This is the description of the `"Table`" section"

		[System.Collections.ArrayList]$TheTable = @()
		$ATable = New-Object PSObject
			$ATable | Add-Member -NotePropertyName FirstRow -NotePropertyValue "A value of the first column"
			$ATable | Add-Member -NotePropertyName SecondRow -NotePropertyValue "A value of the second column"
			$ATable | Add-Member -NotePropertyName ThirdRow -NotePropertyValue "A value of the third column"
			$ATable | Add-Member -NotePropertyName FourthRow -NotePropertyValue "A value of the fourth column"
			$ATable | Add-Member -NotePropertyName FifthRow -NotePropertyValue "A value of the fifth column"
		$null = $TheTable.Add($ATable)

		$ATable = New-Object PSObject
			$ATable | Add-Member -NotePropertyName FirstRow -NotePropertyValue "Another value of the first column"
			$ATable | Add-Member -NotePropertyName SecondRow -NotePropertyValue "Another value of the second column"
			$ATable | Add-Member -NotePropertyName ThirdRow -NotePropertyValue "Another value of the third column"
			$ATable | Add-Member -NotePropertyName FourthRow -NotePropertyValue "Another value of the fourth column"
			$ATable | Add-Member -NotePropertyName FifthRow -NotePropertyValue "Another value of the fifth column"
		$null = $TheTable.Add($ATable)

	[PSCustomObject]$TheCommand = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description -DataType "TableStructure" -EffectiveDataTableStructure $TheTable -TableType "Table"
	$null = $TheObjectToConvertToHTML.Add($TheCommand)

	### "String2" section
	[string]$SectionTitle = "This is the title of the `"String2`" section"
	[ConsoleColor]$SectionTitleColor = "Black"
	[string]$Description = "This is the description of the `"String2`" section"
	[string]$TheString = "This is the content of the `"String2`" section"
		[PSCustomObject]$TheCommand = Prepare-ObjectForHTMLReport -SectionTitle $SectionTitle -SectionTitleColor $SectionTitleColor -Description $Description -DataType "String" -EffectiveDataString $TheString
	$null = $TheObjectToConvertToHTML.Add($TheCommand)


	# Export the content to an HTML report
	### Create the folder, if doesn't exist
	$null = New-Item -ItemType Directory "$Env:Temp\DimCryGarage" -Force

	### Manually set a name of the file
	$FilePath = "$Env:Temp\DimCryGarage\HTMLReport.HTML"
	$PageTitle = "This is the title of the page"
	$ReportTitle = "This is the title of the report"

	Export-ReportToHTML -FilePath $FilePath -PageTitle $PageTitle -ReportTitle $ReportTitle -TheObjectToConvertToHTML $TheObjectToConvertToHTML

	Write-Host "The HTML report is located on:" -ForegroundColor White
	Write-Host "`tLong path: " -ForegroundColor White -NoNewline
	Write-Host "$Env:Temp\DimCryGarage\HTMLReport.HTML" -ForegroundColor Cyan
	Write-Host "`tShort path: " -ForegroundColor White -NoNewline
	Write-Host "%temp%\DimCryGarage\HTMLReport.HTML" -ForegroundColor Cyan

}

function Write-Log {
    param (
        [ValidateSet("INFO","SUCCESS","WARNING","ERROR")]
        [string]$Type,
        [PSCustomObject]$Description,
        [ConsoleColor]$ForegroundColor,
        [switch]$Interactive = $true,
        [switch]$LinkedToNextEntry,
        [switch]$SameLineLikePreviousEntry,
        [string]$KeyEntry = $null
    )

    if (!$ForegroundColor) {
        switch ($Type) {
            "Success" {$ForegroundColor = "Green"}
            "Warning" {$ForegroundColor = "Yellow"}
            "Error"   {$ForegroundColor = "Red"}
            Default   {
                if (@($Global:LogObject).Count -gt 0) {
                    $ForegroundColor = $($Global:LogObject[0].ForegroundColor)
                }
                else {
                    $ForegroundColor = "White"
                }
            }
        }
    }

    if ($Interactive) {
        [bool]$Interactive = $true
    }
    else {
        [bool]$Interactive = $false
    }

    $TheObject = New-Object PSObject
        if ($Type) {
            [string]$Date = $((Get-Date).ToUniversalTime().ToString("dd.MM.yyyy HH:mm:ss.fff")) + " UTC"

            $TheObject | Add-Member -MemberType NoteProperty -Name Date -Value $Date
            $TheObject | Add-Member -MemberType NoteProperty -Name Type -Value $Type

        }
        else {
            $TheObject | Add-Member -MemberType NoteProperty -Name Date -Value ""
            $TheObject | Add-Member -MemberType NoteProperty -Name Type -Value ""
        }

        $TheObject | Add-Member -MemberType NoteProperty -Name Description -Value $Description
        $TheObject | Add-Member -MemberType NoteProperty -Name ForegroundColor -Value $ForegroundColor
        $TheObject | Add-Member -MemberType NoteProperty -Name SameLineLikePreviousEntry -Value $SameLineLikePreviousEntry
        $TheObject | Add-Member -MemberType NoteProperty -Name KeyEntry -Value $KeyEntry
        $TheObject | Add-Member -MemberType NoteProperty -Name LinkedToNextEntry -Value $LinkedToNextEntry
        $TheObject | Add-Member -MemberType NoteProperty -Name Interactive -Value $Interactive
    
    if (!(Test-Path $Global:HTMLLogPath)) {
        Create-HTMLLogFile
    }        

    if ($TheObject.Date) {
        $Global:HTMLLogEntry.Date = $TheObject.Date
    }
    if ($TheObject.Type) {
        $Global:HTMLLogEntry.Type = $TheObject.Type
    }

    if (!$LinkedToNextEntry) {
        [string]$WriteHostTypeString = $null

        $null = $Global:LogObject.Add($TheObject)

        if (@($($TheObject.Description)).Count -gt 1) {
            if (@($($TheObject.Description) | Get-Member -MemberType NoteProperty).Count -gt 1) {

                [System.Collections.ArrayList]$TheProperties = @()
                $ListOfMembers = ($($TheObject.Description)| Get-Member -MemberType NoteProperty).Name
                $null = $TheProperties.Add($($TheObject.KeyEntry))

                foreach ($MemberEntry in $ListOfMembers) {
                    if ($MemberEntry -ne $($TheObject.KeyEntry)) {
                        $null = $TheProperties.Add($MemberEntry)
                    }
                }

                [string]$EntryToUse = $($TheObject.Description) | ConvertTo-Html -As Table -Property $TheProperties -Fragment
                $Global:HTMLLogEntry.Description = $Global:HTMLLogEntry.Description + $EntryToUse
            }
            else {
                [int]$i = 1
                foreach ($DescriptionEntry in $($TheObject.Description)) {
                    if ($i -le $(@($TheObject.Description).Count - 1)) {
                        $Global:HTMLLogEntry.Description = $Global:HTMLLogEntry.Description + $DescriptionEntry + "<br>"
                    }
                    else {
                        $Global:HTMLLogEntry.Description = $Global:HTMLLogEntry.Description + $DescriptionEntry
                    }
                    $i++
                }
            }
        }
        else {
            if ($Global:HTMLLogEntry.Description) {
                if ($SameLineLikePreviousEntry) {
                    if ($Global:HTMLLogEntry.Description.EndsWith("<br>")) {
                        $Global:HTMLLogEntry.Description = $($Global:HTMLLogEntry.Description).Substring(0,$($Global:HTMLLogEntry.Description).Length-4) + $($TheObject.Description) + "<br>"
                    }
                    else {
                        $Global:HTMLLogEntry.Description = $Global:HTMLLogEntry.Description + $($TheObject.Description) + "<br>"
                    }
                }
                else {
                    $Global:HTMLLogEntry.Description = $Global:HTMLLogEntry.Description + $($TheObject.Description) + "<br>"
                }
            }
            else {
                $Global:HTMLLogEntry.Description = $($TheObject.Description)
            }
        }
        
        [int]$i = 1
        foreach ($ObjectEntry in $Global:LogObject) {
            if ($i -eq 1) {
                ### For HTML

                ### For CSV and Write-Host
                Write-EntryOnScreenOrLOGFileSameOrNextLine -ObjectEntry $ObjectEntry -IsFirstEntry $true
            }
            else {
                ### For HTML

                ### For CSV and Write-Host
                Write-EntryOnScreenOrLOGFileSameOrNextLine -ObjectEntry $ObjectEntry -IsFirstEntry $false
            }
            $i++
        }

        Add-EntryInHTMLLogFile -theLogToAddInHTML $Global:HTMLLogEntry -description $($Global:HTMLLogEntry.Description)

        $Global:LogObject.Clear()
        $Global:HTMLLogEntry.Date = ""
        $Global:HTMLLogEntry.Type = ""
        $Global:HTMLLogEntry.Description = ""
        
    }
    else {
        $null = $Global:LogObject.Add($TheObject)

        if (@($($TheObject.Description) | Get-Member -MemberType NoteProperty).Count -gt 1) {
            
            [System.Collections.ArrayList]$TheProperties = @()
            $ListOfMembers = ($($TheObject.Description)| Get-Member -MemberType NoteProperty).Name
            $null = $TheProperties.Add($($TheObject.KeyEntry))

            foreach ($MemberEntry in $ListOfMembers) {
                if ($MemberEntry -ne $($TheObject.KeyEntry)) {
                    $null = $TheProperties.Add($MemberEntry)
                }
            }

            [string]$EntryToUse = $($TheObject.Description) | ConvertTo-Html -As Table -Property $TheProperties -Fragment
        }
        else {
            [string]$EntryToUse = $($TheObject.Description)
        }
        
        if (@($($TheObject.Description)).Count -gt 1) {
            $Global:HTMLLogEntry.Description = $Global:HTMLLogEntry.Description + $EntryToUse
        }
        else {
            if ($SameLineLikePreviousEntry) {
                if ($Global:HTMLLogEntry.Description.EndsWith("<br>")) {
                    $Global:HTMLLogEntry.Description = $($Global:HTMLLogEntry.Description).Substring(0,$($Global:HTMLLogEntry.Description).Length-4) + $($TheObject.Description) + "<br>"
                }
                else {
                    $Global:HTMLLogEntry.Description = $Global:HTMLLogEntry.Description + $($TheObject.Description) + "<br>"
                }
            }
            else {
                $Global:HTMLLogEntry.Description = $Global:HTMLLogEntry.Description + $($TheObject.Description) + "<br>"
            }
        }
    }
}


function Write-EntryOnScreenOrLOGFileSameOrNextLine {
    param (
        [bool]$IsFirstEntry,
        $ObjectEntry
    )

    [string]$WriteHostDescriptionString = " " * 40

    if ($ObjectEntry.Type) {
        if ($($ObjectEntry.Type) -eq "INFO") {
            $WriteHostTypeString = " " * 3
        }
        elseif ($($ObjectEntry.Type) -eq "ERROR") {
            $WriteHostTypeString = " " * 2
        }
    }

    if ($($ObjectEntry.SameLineLikePreviousEntry)) {
        $SameLineLikePreviousEntry = $true
    }
    else {
        $SameLineLikePreviousEntry = $false
    }

    if ($($ObjectEntry.LinkedToNextEntry)) {
        $LinkedToNextEntry = $true
    }
    else {
        $LinkedToNextEntry = $false
    }

    [int]$i = 1
    foreach ($DescriptionEntry in $($ObjectEntry.Description)) {
        if (@($DescriptionEntry | Get-Member -MemberType NoteProperty).Count -gt 1) {
            [System.Collections.ArrayList]$DescriptionEntryMembers = Create-DescriptionTableMembers -ListOfMembers (($DescriptionEntry | Get-Member -MemberType NoteProperty).Name) -KeyEntry $($ObjectEntry.KeyEntry)
            [string]$WriteHostMemberString = Create-DescriptionTableEntries -DescriptionEntryMembers $DescriptionEntryMembers -ObjectEntryDescription $($ObjectEntry.Description) -DescriptionEntry $DescriptionEntry
            [string]$EntryToUse = $WriteHostMemberString
        }
        else {
            [string]$EntryToUse = $DescriptionEntry
        }

        if (($i -eq 1) -and ($IsFirstEntry)) {
            if ($ObjectEntry.Interactive) {
                Write-Host $("`n" + "-" * $($ObjectEntry.Date).Length + "   " + "-" * $($ObjectEntry.Type + $WriteHostTypeString).Length + "   " + "-" * $($EntryToUse).Length) -ForegroundColor $($ObjectEntry.ForegroundColor) -NoNewline
                Write-Host "`n$($ObjectEntry.Date)   $($ObjectEntry.Type)$WriteHostTypeString   $EntryToUse" -ForegroundColor $($ObjectEntry.ForegroundColor) -NoNewline
            }

            if (!(Test-Path $Global:LogPath)) {
                Add-Content -LiteralPath $Global:LogPath -Value $("Date                       " + "   " + "Type   " + "   " + "Description") -Force
            }
            
            if ($LinkedToNextEntry) {
                $Global:LogEntry = $("-" * $($ObjectEntry.Date).Length + "   " + "-" * $($ObjectEntry.Type + $WriteHostTypeString).Length + "   " + "-" * $($EntryToUse).Length)
                $Global:LogEntry = $Global:LogEntry + "`n$($ObjectEntry.Date)   $($ObjectEntry.Type)$WriteHostTypeString   $EntryToUse"
            }
            else {
                Add-Content -LiteralPath $Global:LogPath -Value $("-" * $($ObjectEntry.Date).Length + "   " + "-" * $($ObjectEntry.Type + $WriteHostTypeString).Length + "   " + "-" * $($EntryToUse).Length) -Force
                Add-Content -LiteralPath $Global:LogPath -Value "$($ObjectEntry.Date)   $($ObjectEntry.Type)$WriteHostTypeString   $EntryToUse" -Force
            }
        }
        else {
            if ($SameLineLikePreviousEntry) {
                if ($ObjectEntry.Interactive) {
                    Write-Host "$EntryToUse" -ForegroundColor $($ObjectEntry.ForegroundColor) -NoNewline
                }

                if ($LinkedToNextEntry) {
                    $Global:LogEntry = $Global:LogEntry + $EntryToUse
                }
                else {
                    if ($i -le $(@($TheObject.Description).Count - 1)) {
                        $Global:LogEntry = $Global:LogEntry + $EntryToUse
                    }
                    else {
                        Add-Content -LiteralPath $Global:LogPath -Value "$Global:LogEntry$EntryToUse" -Force
                        [string]$Global:LogEntry = $null
                    }
                }
            }
            else {
                if ($ObjectEntry.Interactive) {
                    Write-Host "`n$WriteHostDescriptionString$EntryToUse" -ForegroundColor $($ObjectEntry.ForegroundColor) -NoNewline
                }

                if ($LinkedToNextEntry) {
                    if ($Global:LogEntry) {
                        $Global:LogEntry = $Global:LogEntry + "`n" + $WriteHostDescriptionString + $EntryToUse
                    }
                    else {
                        $Global:LogEntry = $WriteHostDescriptionString + $EntryToUse
                    }
                }
                else {
                    if ($i -le $(@($TheObject.Description).Count - 1)) {
                        if ($Global:LogEntry) {
                            $Global:LogEntry = $Global:LogEntry + "`n" + $WriteHostDescriptionString + $EntryToUse
                        }
                        else {
                            $Global:LogEntry = $WriteHostDescriptionString + $EntryToUse
                        }
                    }
                    else {
                        if ($Global:LogEntry) {
                            Add-Content -LiteralPath $Global:LogPath -Value $Global:LogEntry -Force
                            [string]$Global:LogEntry = $null
                        }
                        Add-Content -LiteralPath $Global:LogPath -Value "$WriteHostDescriptionString$EntryToUse" -Force
                    }
                }
            }
        }
        $i++
    }
}


function Create-DescriptionTableMembers {
    param (
        $ListOfMembers,
        $KeyEntry
    )

    [System.Collections.ArrayList]$DescriptionEntryMembers = @()
    $null = $DescriptionEntryMembers.Add($KeyEntry)

    foreach ($MemberEntry in $ListOfMembers) {
        if ($MemberEntry -ne $KeyEntry) {
            $null = $DescriptionEntryMembers.Add($MemberEntry)
        }
    }

    return $DescriptionEntryMembers
}


function Create-DescriptionTableEntries {
    param (
        [System.Collections.ArrayList]$DescriptionEntryMembers,
        $ObjectEntryDescription,
        $DescriptionEntry
    )

    [int]$i = 1
    [string]$WriteHostMemberString = ""

    foreach ($Member in $DescriptionEntryMembers) {
        $MaximumLengthOfMember = ($ObjectEntryDescription.$Member | sort {$_.ToString().Length} -Descending | select -First 1).ToString().Length
        $WriteHostMemberStringToAdd = " " * $($MaximumLengthOfMember - $DescriptionEntry.$Member.ToString().Length)

        if ($i -eq 1) {
            $WriteHostMemberString = "| " + $DescriptionEntry.$Member + $WriteHostMemberStringToAdd
        }
        elseif ($i -eq @($DescriptionEntry | Get-Member -MemberType NoteProperty).Count) {
            $WriteHostMemberString = $WriteHostMemberString + " | " + $DescriptionEntry.$Member + $WriteHostMemberStringToAdd + " |"
        }
        else {
            $WriteHostMemberString = $WriteHostMemberString + " | " + $DescriptionEntry.$Member + $WriteHostMemberStringToAdd
        }

        $i++
    }
    
    return $WriteHostMemberString
}


function Create-HTMLLogFile {
    param (
        [string]$lineToAdd
    )

$HTMLBeginning = @"
<!DOCTYPE html>
<html>
	<head>
	
		<meta charset="UTF-8">
		<meta name="description" content="DimCry's Garage">
		<meta name="keywords" content="PowerShell, HTML, Report, DimCry, Garage">
		<meta name="author" content="Cristian Dimofte">
		<meta name="viewport" content="width=device-width, initial-scale=1.0">
		
        <title>Filtrable / Searchable log file</title>
		<link rel="titlebar icon" href="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD0AAABOCAYAAACT1B6GAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAGeSURBVHhe7duxSsNQFMbxvoCT7+Ab+Aw+hg/g6uDk4Kq4ODm4OwiC4uKkkwgigkipiBYVsWrVKmY95sBHPCml5EoOSXq/4Q8hvS3nBx1u0rSVJInEFtGxRPTeZV9aS+2JSk3WqBFN9IQUjJ5a7sjaUU92LvqNSGfVma0hGD290pHTu8/cG+qczqozWwPRiGii04huQEQjolEwevPkRWZWb7Lmtx/kbfCTva7Hes6u8UpnsbO5oXXXY9fMbXXl1aD1WM/ZNV7pLHY2ohHRKBh9cPUuC7uPWevHPfn4/kPrsZ6za7zSWexsbug6RzQiGhFNdBrRDYhoVAq6qguO2Y1bOb8f5GYZzg1d1Ta0yLeOaEQ0CkZXdcGxuP8k189fuVmGc0PXOaIR0YhootOIbkBEo1LQXpuTs+74i4kiuaG9tqGjhguN6DGfS3QRtNdNhMN2/teK/+SGrnNEI6IR0USnEd2AiEaloKN8/KKq+96j0lnsbEQjolEwOsrHJOsc0YhoRDTRaUQ3IKJRMDrKP5hOQkQjoqNExxLRsRQhOpFfossxSEw+JV4AAAAASUVORK5CYII=" />
		
		<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css" />
		<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css" />
		<script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>

		<style type="text/css">
			.input {
			  width: 15px;
			  height: 24px;
			  font-size: 18px;
			}

			#myInput {
				font-size: 16px;
				box-sizing: border-box;
				font-family: FabricMDL2Icons;
				font-size: 16px;
				border: none;
				outline: 1px solid #ddd;
				padding-left: 5px;
				padding-bottom: 5px;
			}
			
			#myInput:focus {outline: 3px solid #ddd;}
			
			.tooltip {
			  position: relative;
			  display: inline-block;
			}

			.tooltip .tooltiptext {
			  visibility: hidden;
			  width: 120px;
			  background-color: #fff;
			  color: #0078d7;
			  text-align: center;
			  border-radius: 6px;
			  padding: 5px 0;
			  position: absolute;
			  z-index: 1;
			  bottom: 125%;
			  left: 50%;
			  margin-left: -60px;
			  opacity: 0;
			  transition: opacity 0.5s;
			}

			.tooltip .tooltiptext::after {
			  content: "";
			  position: absolute;
			  top: 100%;
			  left: 75%;
			  margin-left: -5px;
			  border-width: 5px;
			  border-style: solid;
			  border-color: #555 transparent transparent transparent;
			}

			.tooltip:hover .tooltiptext {
			  visibility: visible;
			  opacity: 10;
			}

			:root {
				font-size: calc(16px + (24 - 16)*(100vw - 320px)/(1920 - 320));
			}
			input {
				line-height: 1.5em;
				color: #171717;
			}

			.dropButtonType {
				border: none;
				cursor: pointer;
				border-style: none;
				float: right;
			}

			.dropButtonType:hover, .dropButtonType:focus {
				border: none;
				border-style: none;
			}

			.dropdown-content {
				display: none;
				position: absolute;
				background-color: #f6f6f6;
				overflow: auto;
				border-style: none;
				z-index: 1;
				resize: both;
				min-height: 110px;
				min-width: 130px;
				margin-top: 10px;
			}

			.dropdown-content a {
				color: black;
				padding: 12px 16px;
				text-decoration: none;
				display: block;
				border-style: none;
			}

			.show {
				display: block;
				border-style: none;
			}
		
			html {
				box-sizing: border-box;
				font-family: FabricMDL2Icons;
			}
			
			*, *:before, *:after {
				box-sizing: inherit;
			}			
			
			.ms-font-su {
				font-family:FabricMDL2Icons;
				-webkit-font-smoothing:antialiased;
				font-weight:100
			}
			
			.ms-fontColor-themePrimary,.ms-fontColor-themePrimary--hover:hover{
				color:#0078d7
			}
			
			@font-face{
				font-family:FabricMDL2Icons;
				src:url(https://spoprod-a.akamaihd.net/files/fabric/assets/icons/fabricmdl2icons.woff) format('woff'),url(https://spoprod-a.akamaihd.net/files/fabric/assets/icons/fabricmdl2icons.ttf) format('truetype');
				font-weight:400;
				font-style:normal
			}

			body {
				padding: 10px;
				font-family: FabricMDL2Icons;
				background: #f6f6f6;
				overflow: auto;
			}
			
			a {
				color: #06c;
			}

			h5 {
				display: block;
				padding: 0;
				margin: 5px 0px 0px 0px;
				font-size: 42px;
				font-weight: 100;
				font-family: FabricMDL2Icons;
			}
			
			p {
				margin: 0 0 1em;
				padding: 10px;
				font-family: FabricMDL2Icons;
			}

			.ms-icon--filtersolid:before {
				font-family: FabricMDL2Icons;
				content: "\F412";
				color: #0078d7;
				margin-right: 0px;
				font-style: normal;
				font-size: 12px;
				float: right;
			}
			
			.ms-icon--sort:before {
				font-family: FabricMDL2Icons;
				content: "\E8CB";
				color: #0078d7;
				margin-right: 5px;
				font-style: normal;
				font-size: 12px;
				float: right;
			}			

			.label {
				width: 100%;
				height: 80px;
				background-color: #f6f6f6;
				color: #0078d7;
				font-size: 46px;
				display: inline-block;
				box-sizing: border-box;
			}			
			
			.logtypeentryerror {
				color: Red;
				font-style: normal;
			}
			
			.logtypeentrywarning {
				color: #ffba00;
				font-style: normal;
			}
			
			.logtypeentrysuccess {
				color: Green;
				font-style: normal;
			}
			
			.logtypeentryinfo {
				color: Black;
				font-style: normal;
			}			
			
		   table {
				font-size: 14px;
				font-weight: 800;
				border: 1px solid #fff;
				border-collapse: collapse;
				font-family: FabricMDL2Icons;
				text-align: left;
				white-space: normal;
			}
			
			td {
				padding: 4px;
				margin: 0px;
				border: 1px solid #fff;
				border-collapse: collapse;
				vertical-align: top;
			}

			th.description {
				border: 1px solid #fff;
				border-collapse: collapse;
				vertical-align: middle;
				text-align: center;
			}
			
			td.description {
				border: 1px solid #fff;
				border-collapse: collapse;
				vertical-align: middle;
			}
			
			th {
				color: #0078d7;
				font-family: FabricMDL2Icons;
				text-transform: uppercase;
				padding: 10px 5px;
				border: 1px solid #fff;
				border-collapse: collapse;
			}
			
			tbody tr:nth-child(odd), tbody.description tr:nth-child(even) {
				background: #e6f3ff;
			}
			
			tbody.description tr:nth-child(odd), tbody tr:nth-child(even) {
				background: #fff;
			}

			@media (max-device-width:480px) and (orientation:landscape) {
				.dropdown-content {
					margin-top: 30px;
				}
			}			
			
			@media screen and (max-width:639px){
				h5 {
					margin: 5px 0px 0px 0px;
					font-size: 8vw;
					font-weight: 100;
					font-family: FabricMDL2Icons;
				}
				
				.dropdown-content {
					margin-top: 20px;
				}
			}
		</style>
	</head>
	<body class="ms-Fabric" dir="ltr">
		<header>
			<div class="label">
				<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAD0AAABOCAYAAACT1B6GAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAGeSURBVHhe7duxSsNQFMbxvoCT7+Ab+Aw+hg/g6uDk4Kq4ODm4OwiC4uKkkwgigkipiBYVsWrVKmY95sBHPCml5EoOSXq/4Q8hvS3nBx1u0rSVJInEFtGxRPTeZV9aS+2JSk3WqBFN9IQUjJ5a7sjaUU92LvqNSGfVma0hGD290pHTu8/cG+qczqozWwPRiGii04huQEQjolEwevPkRWZWb7Lmtx/kbfCTva7Hes6u8UpnsbO5oXXXY9fMbXXl1aD1WM/ZNV7pLHY2ohHRKBh9cPUuC7uPWevHPfn4/kPrsZ6za7zSWexsbug6RzQiGhFNdBrRDYhoVAq6qguO2Y1bOb8f5GYZzg1d1Ta0yLeOaEQ0CkZXdcGxuP8k189fuVmGc0PXOaIR0YhootOIbkBEo1LQXpuTs+74i4kiuaG9tqGjhguN6DGfS3QRtNdNhMN2/teK/+SGrnNEI6IR0USnEd2AiEaloKN8/KKq+96j0lnsbEQjolEwOsrHJOsc0YhoRDTRaUQ3IKJRMDrKP5hOQkQjoqNExxLRsRQhOpFfossxSEw+JV4AAAAASUVORK5CYII=" alt="DimCry's Garage" width="auto" height="90%" style="float: left; margin-right: 15px">
			
				<div style="width: 100%; height: 100%; ">
					<h5 class="ms-font-su ms-fontColor-themePrimary" style="font-weight: 350; ">
						Filtrable / Searchable log file
					</h5>
				</div>
			</div>
        </header>
	
		<div style="margin-top: 10px;">
					<form >
						<select name="searchIn" id="searchIn" onchange="selectFromWhereToSearch(value)">
							<option value="allEntries">Search through all entries</option>
							<option value="descriptionEntries">Search through description entries</option>
						</select>
						<input type="text" id="myInput" onkeyup="searchFunction(window.theValue)" placeholder="Type here to search..." name="search" title="Type anything to search..." autocomplete="off" style="margin-top: 10px;">
					</form>
		</div>

		<div style="margin: 10px 0px 0px 0px; display: block;">
			<table id="myTable">
				<tr class="header">
					<th>
						Date
						<i onclick="sortTable(0)" style="font-style: normal; cursor: pointer;">
							<i class="ms-icon ms-icon--sort"></i>
						</i>
					</th>
					<th class="tooltip">
						Type&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<i class="tooltiptext">Sort or filter</i>
						<i onclick="showDropDownFilter('myTypeDropdown')" class="dropButtonType ms-icon ms-icon--filtersolid" ></i>
						
						<div id="myTypeDropdown" class="dropdown-content" >
							
							<form style="float:left; font-size: 15px;">
								
								<input type="checkbox" id="showall" onclick="window.visibleEntries = filterTypeColumn('showall')" checked="checked" value="showall">
									<label for="showall">Show all...</label><br />
								<input type="checkbox" id="info" onclick="window.visibleEntries = filterTypeColumn('info')" checked="checked" value="info">
									<label for="info">INFO</label><br>
								<input type="checkbox" id="success" onclick="window.visibleEntries = filterTypeColumn('success')" checked="checked" value="success">
									<label for="success">SUCCESS</label><br />
								<input type="checkbox" id="warning" onclick="window.visibleEntries = filterTypeColumn('warning')" checked="checked" value="warning">
									<label for="warning">WARNING</label><br />
								<input type="checkbox" id="error" onclick="window.visibleEntries = filterTypeColumn('error')" checked="checked" value="error">
									<label for="error">ERROR</label><br />
							</form>
												
						</div>
						<i onclick="sortTable(1)" style="font-style: normal; cursor: pointer;">
							<i class="ms-icon ms-icon--sort"></i>
						</i>
					</th>
					<th>Description</th>
				</tr>
"@

$HTMLEnd = @"
			</table></div>

		<div class="ms-font-su ms-fontColor-themePrimary" style="font-size: 15px; text-align: left; margin-left: 0px;">
			<p style="margin:0px;line-height:0.5;"></p>
			<p style="margin:0px;line-height:0.5;">Creation Date: $((Get-date).ToUniversalTime()) UTC</p>
			<p style="margin:0px;line-height:0.5;">&copy; 2020 <a href="https://github.com/dimcry/DimCry-Garage" target="_blank" title="https://github.com/dimcry/DimCry-Garage">DimCry's Garage</a></p>
		</div>
		
		<script>
			var theValue = "allEntries";
			var theTypeFilter = ['info', 'success', 'warning', 'error'];
			var searchValue = "";
			var visibleEntries = [];
			var hiddenEntries = [];
			myTable = document.getElementById("myTable");
			trsFromMyTable = myTable.getElementsByClassName("filtrableTR");
			for (var a = 0; a < trsFromMyTable.length; a++) {
				if (trsFromMyTable[a].style.display === "") {
					visibleEntries.push(trsFromMyTable[a]);
				}
			}
			
			function selectFromWhereToSearch(value) {
				window.theValue = value;
			}			

			function filterTypeColumn(n) {
				var filter, theTable, tr, i, j, txtValue;
				var theVisibleEntries = [];
				
				theTable = document.getElementById("myTable");
				tr = theTable.getElementsByClassName("filtrableTR");
				
				if (document.getElementById(String(n)).checked == true) {
					if (n.toLowerCase() == "showall") {
						theCheckBox("info", true);
						theCheckBox("success", true);
						theCheckBox("warning", true);
						theCheckBox("error", true);
						theTypeFilter = ['info', 'success', 'warning', 'error'];
					} else {
						if (n.toLowerCase() == "info") {
							theTypeFilter.push('info');
						}
						if (n.toLowerCase() == "success") {
							theTypeFilter.push('success');
						}
						if (n.toLowerCase() == "warning") {
							theTypeFilter.push('warning');
						}
						if (n.toLowerCase() == "error") {
							theTypeFilter.push('error');
						}
						
						if (theTypeFilter.includes('info') && theTypeFilter.includes('success') && theTypeFilter.includes('warning') && theTypeFilter.includes('error')) {
							theCheckBox("showall", true);
						}
					}
				} else {
					if (n.toLowerCase() == "showall") {
						theCheckBox("info", false);
						theCheckBox("success", false);
						theCheckBox("warning", false);
						theCheckBox("error", false);
						theTypeFilter = [];
					}
					else {
						theCheckBox("showall", false);
						if (n.toLowerCase() == "info") {
							theTypeFilter.splice(theTypeFilter.indexOf('info'), 1);
						}
						if (n.toLowerCase() == "success") {
							theTypeFilter.splice(theTypeFilter.indexOf('success'), 1);
						}
						if (n.toLowerCase() == "warning") {
							theTypeFilter.splice(theTypeFilter.indexOf('warning'), 1);
						}
						if (n.toLowerCase() == "error") {
							theTypeFilter.splice(theTypeFilter.indexOf('error'), 1);
						}
					}
				}

				for (i = 0; i < tr.length; i++) {
					td = tr[i].getElementsByTagName("td")[1];
					if (td) {
						txtValue = td.textContent || td.innerText;
						if (theTypeFilter.length > 0) {
							if (theTypeFilter.includes(txtValue.toLowerCase())) {
								tr[i].style.display = "";
								theVisibleEntries.push(tr[i]);
							} else {
								tr[i].style.display = "none";
							}
						}
						else {
							tr[i].style.display = "none";
						}
					}
				}
					searchFunction(window.searchValue);
					
					return theVisibleEntries;				
			}		

			function theCheckBox(elementID, boolValue) {
				let inputs = document.getElementById(elementID);
				inputs.checked = boolValue;
			}

			function triggerAlert(n) {
				alert(n);
			}
			
			/* When the user clicks on the filter button,
			toggle between hiding and showing the dropdown content */
			function showDropDownFilter(filter) {
			  document.getElementById(filter).classList.toggle("show");
			}

			function sortTable(n) {
				var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
				table = document.getElementById("myTable");
				switching = true;
				
				//Set the sorting direction to ascending:
				dir = "asc"; 
				
				/*Make a loop that will continue until
				no switching has been done:*/
				while (switching) {
					//start by saying: no switching is done:
					switching = false;
					rows = table.rows;
					
					/*Loop through all table rows (except the
					first, which contains table headers):*/
					for (i = 1; i < (rows.length - 1); i++) {
						//start by saying there should be no switching:
						shouldSwitch = false;
						
						/*Get the two elements you want to compare,
						one from current row and one from the next:*/
						x = rows[i].getElementsByTagName("TD")[n];
						y = rows[i + 1].getElementsByTagName("TD")[n];
						
						/*check if the two rows should switch place,
						based on the direction, asc or desc:*/
						if (dir == "asc") {
							if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
								//if so, mark as a switch and break the loop:
								shouldSwitch= true;
								break;
							}
						} else if (dir == "desc") {
							if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
								//if so, mark as a switch and break the loop:
								shouldSwitch = true;
								break;
							}
						}
					}
					if (shouldSwitch) {
						/*If a switch has been marked, make the switch
						and mark that a switch has been done:*/
						rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
						switching = true;
						//Each time a switch is done, increase this count by 1:
						switchcount ++;      
					} else {
						/*If no switching has been done AND the direction is "asc",
						set the direction to "desc" and run the while loop again.*/
						if (switchcount == 0 && dir == "asc") {
							dir = "desc";
							switching = true;
						}
					}
				}
			}

			window.addEventListener("click", function(event) {
				if (!event.target.matches('.dropButtonDate')) {
					if (!event.target.matches('.dropButtonType')) {
						var dropdowns = document.getElementsByClassName("dropdown-content");
					}	
					else {
						var dropdowns = document.getElementById("myDateDropdown");
					}
				}
				else if (!event.target.matches('.dropButtonType')) {
					if (!event.target.matches('.dropButtonDate')) {
						var dropdowns = document.getElementsByClassName("dropdown-content");
					}
					else {
						var dropdowns = document.getElementById("myTypeDropdown");
					}
				}

				if (dropdowns) {
					for (let i = 0; i < dropdowns.length; i++) {
						var openDropdown = dropdowns[i];
						if (openDropdown.classList.contains('show')) {
							openDropdown.classList.remove('show');
						}
					}
				}
			})
		
			function searchFunction(value) {
				var input, filter, theTable, tr, td, i, j;
				input = document.getElementById("myInput");
				window.searchValue = value;
				filter = input.value.toUpperCase();
				theTable = document.getElementById("myTable");

				tr = window.visibleEntries;

				if (theTypeFilter.length > 0) {
					if (value == "allEntries") {
						for (i = 0; i < tr.length; i++) {
							var tds = tr[i].getElementsByTagName("td");
							txtValue = tds[1].textContent || tds[1].innerText;
							var flag = false;
							if (theTypeFilter.includes(txtValue.toLowerCase())) {
								for (j = 0; j < tds.length; j++){
									var td = tds[j];
									if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
										flag = true;
									}
								}
							}
							if (flag){
								tr[i].style.display = "";
							}
							else {
								tr[i].style.display = "none";
							}
						}
					}
					else if (value == "descriptionEntries") {
						for (i = 0; i < tr.length; i++) {
							var tds = tr[i].getElementsByTagName("td");
							txtValue = tds[1].textContent || tds[1].innerText;
							var flag = false;
							if (theTypeFilter.includes(txtValue.toLowerCase())) {
								var td = tds[2];
								if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {
									tr[i].style.display = "";
								}
								else {
									tr[i].style.display = "none";
								}
							}
						}
					}
				}
				else {
					for (i = 0; i < tr.length; i++) {
							tr[i].style.display = "none";
					}
				}
			}
		</script>
	</body>
</html>			
"@
    [string]$TheBody = $HTMLBeginning
    $TheBody = $TheBody + $lineToAdd
    $TheBody = $TheBody + $HTMLEnd
    $TheBody | Out-File $Global:HTMLLogPath -Force

}


function Add-EntryInHTMLLogFile {
    param (
        $theLogToAddInHTML,
        $description
    )

    [string]$theLineToAdd = "
				`<tr class=`"filtrableTR`"`>
					`<td class=`"filtrableTD`"`>$($theLogToAddInHTML.Date)`<`/td`>
					`<td class=`"filtrableTD`"`>`<i class=`"logtypeentry$($theLogToAddInHTML.Type.ToLower())`"`>$($theLogToAddInHTML.Type)`<`/i`>`<`/td`>
					`<td class=`"filtrableTD`"`>$description`<`/td`>
				`<`/tr`>
`<`/table`>`<`/div`>
"

    $theHTMLFile = Get-Content $Global:HTMLLogPath -Force
    $newHTMLFile = $theHTMLFile.Replace("</table></div>",$theLineToAdd)

    $newHTMLFile | Out-File $Global:HTMLLogPath -Force

}


function Test-WriteLogWithDummyData {

    ### Testing single entry with a description that contains an table with one column
    [System.Collections.ArrayList]$MultipleDescriptionEntries = @()
    $null = $MultipleDescriptionEntries.Add("First entry")
    $null = $MultipleDescriptionEntries.Add("Second entry")
    $null = $MultipleDescriptionEntries.Add("Third entry")
    $null = $MultipleDescriptionEntries.Add("Fourth entry")
        Write-Log -Type ERROR -Description $MultipleDescriptionEntries

    ### Testing multiple entries, that contains string or table with one column
    Write-Log -Type WARNING -Description "Testing Multiple entries:" -LinkedToNextEntry
    Write-Log -Description $MultipleDescriptionEntries -ForegroundColor White

    ### Testing multiple entries, that contains string or table with multiple columns - first entry is string
    Write-Log -Type WARNING -Description "The list of the first 5 services:" -LinkedToNextEntry
    Write-Log -Description (Get-Service | select -First 5 Name, Status) -LinkedToNextEntry -KeyEntry Name
    Write-Log -Description "Dummy text to check if it will be listed as expected" -ForegroundColor Cyan

    ### Testing single entries, that contains a table with multiple columns
    Write-Log -Type INFO -Description (Get-PSDrive | select -First 3 Name, @{n="Provider"; e={$_.Provider.ToString()}}, Root) -KeyEntry Name

    ### Testing multiple entries, that contains string or table with multiple columns - first entry is table
    Write-Log -Type WARNING -Description (Get-Service | select -First 5 Name, Status) -LinkedToNextEntry -KeyEntry Name
    Write-Log -Description "Dummy text to check if it will be listed as expected" -ForegroundColor White

    ### Testing multiple entries, that contains string or table with multiple columns, and ensure the first column in listed table is the selected one (KeyEntry) - first entry is string
    Write-Log -Type ERROR -Description "The list of the first 5 processes:" -LinkedToNextEntry
    Write-Log -Description (Get-Process | select -First 5 ProcessName, Id, Handles, SI) -KeyEntry ProcessName

    ### Testing single entries with a description that contains string
    Write-Log -Type INFO -Description "An INFO text"
    Write-Log -Type WARNING -Description "An WARNING text"
    Write-Log -Type SUCCESS -Description "An SUCCESS text"

    ### Testing single entries with a description that contains string - NonInteractive, means will not be listed on screen
    Write-Log -Type INFO -Description "An INFO text, NonInteractive" -Interactive:$false
    Write-Log -Type WARNING -Description "An WARNING text, NonInteractive" -Interactive:$false
	Write-Log -Type SUCCESS -Description "An SUCCESS text, NonInteractive" -Interactive:$false

    ### Testing multiple string entries, that need to list test on the same line, or next line, with different colors
    Write-Log -Type INFO -Description "Details for " -LinkedToNextEntry
    Write-Log -Description "DimCry" -ForegroundColor Cyan -SameLineLikePreviousEntry -LinkedToNextEntry
    Write-Log -Description " user:" -SameLineLikePreviousEntry -LinkedToNextEntry
    Write-Log -Description "Alias: " -LinkedToNextEntry
    Write-Log -Description "DimCry" -ForegroundColor Cyan -SameLineLikePreviousEntry

    ### Testing single entries with a description that contains string
    Write-Log -Type WARNING -Description "An WARNING text"

    ### Testing multiple entries, that contains multiple tables with one or multiple columns
    Write-Log -Type ERROR -Description (Get-Process | select -First 5 ProcessName, Id, Handles, SI) -KeyEntry ProcessName -LinkedToNextEntry
    Write-Log -Description $MultipleDescriptionEntries -ForegroundColor Cyan

    Write-Host
    Write-Host
    Write-Host
	Write-Host "The filtrable and searchable .html log file is located on:" -ForegroundColor White
	Write-Host "`t$Global:HTMLLogPath" -ForegroundColor Cyan
	Write-Host "The .log log file is located on:" -ForegroundColor White
	Write-Host "`t$Global:LogPath" -ForegroundColor Cyan
}


#endregion "Functions"
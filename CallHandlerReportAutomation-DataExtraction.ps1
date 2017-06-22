###############################################################################################
# READ-ME
###############################################################################################
#
# Script Name: Call_report_automation
# Author: Simar Chawla
# 
###############################################################################################
# Script Description:
# This script transfers data from a list of specified pdf files, and puts them
# in specific locations in a text file where they can then be easily copied
# into the approriate excel file. This will be used to automate part of the
# Call Report making process.
# Note: Avoided directly transfering data into excel file to allow for user
# checking and avoid potentially deleting valuable data in excel file.
#################################################################################################
#PowerShell version details:
#PSVersion	4.0	
#WSManStackVersion	3.0	
#SerializationVersion	1.1.0.1	
#CLRVersion	4.0.30319.18063	
#BuildVersion	6.3.9600.16406	
#PSCompatibleVersions	{1.0, 2.0, 3.0, 4.0}	
#PSRemotingProtocolVersion	2.2	#>
#################################################################################################
#Libraries Needed:
#isharptext
#################################################################################################
write-host "For Janet and Teresa and Eric's report"
$month = Read-Host -Promt "Enter the month (Jan, Feb, Mar, Apr, May, June, July, Aug, Sept, Oct, Nov, Dec)"
$year = Read-Host -Promt "Enter the year (Ex: 2017)"
write-host "For Brian's report (G:\Administration\A27 - Communication Systems\Telephone Data\COURTS\YEAR"
$date1 = Read-host -Prompt "Enter Date 1 (01,02, 03, 04...) and make sure date is in the file"
$date2 = Read-host -Prompt "Enter Date 2 (make sure date is in the file)"
$date3 = Read-host -Prompt "Enter Date 3 (make sure date is in the file)"
$date4 = Read-host -Prompt "Enter Date 4 (make sure date is in the file)"

#Copy dll file into user
If (!(Test-Path (($env:UserProfile) + "\itextsharp.dll"))){
        copy-item -path "\\ykr-fs2\FI_Information_Technology\Administration\A27 - Communication Systems\CallreportAutomation\itextsharp\itextsharp.dll"
                  -destination "$env:UserProfile"
}

$mainpath = "\\ykr-fs2\FI_Information_Technology\Administration\A27 - Communication Systems\Telephone Data"

#Janet files
#HC-mainmenu-afterhrs
$JanetPath1 = $mainpath + "\HEALTH REPORTS\HC menu call counts\"+$year+"\"+$month+" HC-UPD -main menu -afterhrs.pdf"
#HC-Inspectors afterhrs
$JanetPath2 = $mainpath + "\HEALTH REPORTS\HC menu call counts\"+$year+"\"+$month+" HC -Inspectors afterhrs.pdf"
#HC-UPD main menu
$JanetPath3 = $mainpath + "\HEALTH REPORTS\HC menu call counts\"+$year+"\"+$month+" HC -UPD main menu.pdf"
#HPC HC IDCD
$JanetPath4 = $mainpath + "\HEALTH REPORTS\HC menu call counts\"+$year+"\"+$month+" HC IDCD.pdf"
#HC- Sexual line
$JanetPath5 = $mainpath + "\HEALTH REPORTS\HC menu call counts\"+$year+"\"+$month+" HC -Sexual line.pdf"
#HC-Inspectors
$JanetPath6 = $mainpath + "\HEALTH REPORTS\HC menu call counts\"+$year+"\"+$month+" HC -Inspectors.pdf"

#Teresa's files
#CHS-CC -main menu
$TeresaPath1= $mainpath + "\" +$year+"\"+$month+" Admin CHS CC -main menu.pdf"
#CHS CC -Opt 3 Housing
$TeresaPath2= $mainpath + "\" +$year+"\"+$month+" Admin CHS CC -Opt 3 Housing.pdf"
#Child Care line
$TeresaPath3= $mainpath + "\CS&Hdata\Kids Line Counts\"+$year+"\"+$month+" Child Care line.pdf"
#CHS CC- Opt 1 OW
$TeresaPath4 = $mainpath + "\" +$year+"\"+$month+" Admin CHS CC -Opt 1 OW.pdf"
#CHS CC- Opt 3 subopt 1 Housing Inq
$TeresaPath5 = $mainpath + "\" +$year+"\"+$month+" Admin-CHS CC -Opt 3 -subopt 1 Housing Inq.pdf"
#CHS CC- Opt 3 subopt 3 Housing Stability
$TeresaPath6 = $mainpath + "\" +$year+"\"+$month+" Admin CHS CC -Opt 3 -subopt 3 Housing Stability.pdf"
#CHS CC -Opt 3-1 Market Rent
$TeresaPath7 = $mainpath + "\" +$year+"\"+$month+" Admin-CHS CC -Opt 3 -1 Market Rent.pdf"
#CHS CC -Opt 3-2 Subsidized Housing
$TeresaPath8 = $mainpath + "\" +$year+"\"+$month+" Admin-CHS CC -Opt 3 -2 Subsidized Housing.pdf"

#Eric's files
#EMS AUto Attendant
$EricPath1 = $mainpath + "\EMS\"+$year+"\"+$month+" EMS Auto Attendant.pdf"
#EMS Opt 1
$EricPath2 = $mainpath + "\EMS\"+$year+"\"+$month+" EMS Opt 1.pdf"
#EMS Opt 2
$EricPath3 = $mainpath + "\EMS\"+$year+"\"+$month+" EMS Opt 2.pdf"
#EMS Opt 3
$EricPath4 = $mainpath + "\EMS\"+$year+"\"+$month+" EMS Opt 3.pdf"
#EMS Opt 4
$EricPath5 = $mainpath + "\EMS\"+$year+"\"+$month+" EMS Opt 4.pdf"
#EMS Opt 6
$EricPath6 = $mainpath + "\EMS\"+$year+"\"+$month+" EMS Opt 6.pdf"

#Brian's files
$BrianPath1 = $mainpath + "\COURTS\" + $year + "\" + $month + " x73337 -SSC Collection Line.pdf"

#Textfile where pdf is parsed
$parsetextfile = $env:UserProfile + "\pdftext.txt"
$parsetextfile2 = $env:UserProfile + "\parsetext.txt"
#Textfile where final data is shown and ready to be copy and pasted
$textfileJanet = $env:UserProfile + "\HCRevisedMenucounts.txt"
$textfileTeresa = $env:UserProfile + "\C&HSContactCenterCallCounts.txt"
$textfileEric = $env:UserProfile + "\EMSCallCounts.txt"
$textfileBrian = $env:UserProfile + "\BrianEmail.txt"


$monthnums = @{"Jan" = 01; "Feb" = 02;"Mar" = 03;"Apr" = 04; "May" = 05; "June" = 06; "July" = 07; "Aug" = 08; "Sept" =09;
               "Oct" = 10; "Nov" = 11; "Dec" = 12}

$monthnumber=$monthnums.$month
$daysinmonth= [datetime]::DaysInMonth($year,$monthnumber)

#Empty the text files, so that previous reports do not remain
function ClearOrMake-files{
    param(
        [Parameter(Mandatory=$true)]
        $file
        )

    If (Test-Path $file){Clear-Content $file}
        else{New-item $file -type file}
}

ClearOrMake-files -file $parsetextfile
ClearOrMake-files -file $parsetextfile2
ClearOrMake-files -file $textfileJanet
ClearOrMake-files -file $textfileTeresa
ClearOrMake-files -file $textfileEric
ClearOrMake-files -file $textfileBrian


function Add-rowstotextfile{
    param(
        [Parameter(Mandatory=$true)]
        $file
        )

    for ($i = 0; $i -lt 200; $i++){
    add-content $file "`n"
    }
}

Add-rowstotextfile -file $textfileJanet
Add-rowstotextfile -file $textfileTeresa
Add-rowstotextfile -file $textfileEric

#Function uses isharptext library
function Get-PDFContent {
    param(
        [Parameter(Mandatory=$true)]
        $pdfFile
        ) 
           
    $commandPath = $env:UserProfile
    Add-Type -Path "$($commandPath)\itextsharp.dll"    

    $reader = New-Object iTextSharp.text.pdf.PdfReader $pdfFile
    $strategy = New-Object iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy

    for ($i = 1; $i -lt ($reader.NumberOfPages)+1; ++$i) { 
        [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i, $strategy)
    }

    $Reader.Close();

}


function Get-TotalsObject {
    param(
        [Parameter(Mandatory=$true)]
        $pdfFile
        )

    #Convert PDF File into Textfile
    Get-PDFContent -pdfFile $pdfFile| out-file $parsetextfile
    #Textfile now holds the data and each line in textfile represents an array. This allows us to extract the data we need
    $total = (get-content $parsetextfile)[-2]
    $keys = (get-content $parsetextfile)[-5]
    #Create an array for the key values
    $keyarray = $keys -Split " "
    #Create an object with appropriate names
    $result = New-Object -TypeName PSObject -Property @{TotalCalls = $total; "Key 0"=$keyarray[0];"Key 1"=$keyarray[1];"Key 2"=$keyarray[2];"Key 3"=$keyarray[3];"Key 4"=$keyarray[4];
                                                   "Key 5"=$keyarray[5];"Key 6"=$keyarray[6];"Key 7"=$keyarray[7];"Key 8"=$keyarray[8];"Key 9"=$keyarray[9]; "*"=$keyarray[10];
                                                   "#" = $keyarray[11]; "DTMF" = $keyarray[12]; "InvalidDTMF" = $keyarray[13]; "AfterGreeting" = $keyarray[14]; "HangUp"= $keyarray[15];}
    #Output the object
    $result
}


function Add-ContentAtLine{
    param(
       [Parameter(Mandatory=$true)]
        $pdfFile,
       [Parameter(Mandatory=$true)]
        $textfile,
        [Parameter(Mandatory=$true)]
        $key,
       [Parameter(Mandatory=$true)]
        $row
       )

    if (Test-path $pdfFile){
    $filecontent = get-content $textfile
    $filecontent[$row] = ((Get-TotalsObject -pdfFile $pdfFile).($key))
    $filecontent| set-content $textfile
    }
    }


#Janet's
$firstrowJanet = 4
Add-ContentAtLine -pdfFile $JanetPath1 -key "TotalCalls" -row (4-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath1 -key "Key 0" -row (5-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath2 -key "TotalCalls" -row (6-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "TotalCalls" -row (7-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "Key 0" -row (8-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "Key 1" -row (9-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "Key 2" -row (15-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "Key 3" -row (16-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "Key 4" -row (24-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "Key 5" -row (30-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "Key 6" -row (31-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "Key 8" -row (33-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "AfterGreeting" -row (34-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath3 -key "HangUp" -row (35-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath4 -key "Key 0" -row (17-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath4 -key "Key 1" -row (18-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath4 -key "Key 2" -row (19-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath4 -key "Key 3" -row (20-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath4 -key "Key 4" -row (21-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath4 -key "Key 5" -row (22-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath4 -key "AfterGreeting" -row (23-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath5 -key "Key 0" -row (10-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath5 -key "Key 1" -row (11-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath5 -key "Key 2" -row (12-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath5 -key "Key 3" -row (13-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath5 -key "AfterGreeting" -row (14-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath6 -key "Key 0" -row (25-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath6 -key "Key 1" -row (26-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath6 -key "Key 2" -row (27-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath6 -key "Key 3" -row (28-$firstrowJanet) -textfile $textfileJanet
Add-ContentAtLine -pdfFile $JanetPath6 -key "AfterGreeting" -row (29-$firstrowJanet) -textfile $textfileJanet

#Teresa's
$firstrowTeresa = 3
Add-ContentAtLine -pdfFile $TeresaPath1 -key "TotalCalls" -row (3-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 0" -row (4-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 1" -row (5-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 2" -row (13-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 3" -row (24-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 4" -row (52-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 5" -row (53-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 6" -row (54-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 7" -row (55-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 8" -row (56-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "Key 9" -row (57-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "*" -row (58-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "#" -row (59-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "DTMF" -row (60-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "AfterGreeting" -row (61-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath1 -key "HangUp" -row (62-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath2 -key "TotalCalls" -row (25-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath2 -key "Key 0" -row (26-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath2 -key "Key 1" -row (27-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath2 -key "Key 2" -row (43-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath2 -key "Key 3" -row (44-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath2 -key "AfterGreeting" -row (50-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath2 -key "HangUp" -row (51-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "TotalCalls" -row (14-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "Key 0" -row (15-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "Key 1" -row (16-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "Key 2" -row (17-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "Key 3" -row (18-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "Key 4" -row (19-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "Key 5" -row (20-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "Key 6" -row (21-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "AfterGreeting" -row (22-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath3 -key "HangUp" -row (23-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath4 -key "TotalCalls" -row (6-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath4 -key "Key 0" -row (7-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath4 -key "Key 1" -row (8-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath4 -key "Key 2" -row (9-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath4 -key "Key 3" -row (10-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath4 -key "AfterGreeting" -row (11-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath4 -key "HangUp" -row (12-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath5 -key "TotalCalls" -row (25-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath5 -key "Key 0" -row (28-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath5 -key "Key 1" -row (29-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath5 -key "Key 2" -row (33-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath5 -key "AfterGreeting" -row (41-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath5 -key "HangUp" -row (42-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath5 -key "TotalCalls" -row (45-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath6 -key "Key 0" -row (46-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath6 -key "Key 1" -row (47-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath6 -key "AfterGreeting" -row (48-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath6 -key "HangUp" -row (49-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath7 -key "Key 0" -row (30-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath7 -key "AfterGreeting" -row (31-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath7 -key "HangUp" -row (32-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath8 -key "Key 0" -row (34-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath8 -key "Key 1" -row (35-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath8 -key "Key 2" -row (36-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath8 -key "Key 3" -row (37-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath8 -key "Key 4" -row (38-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath8 -key "AfterGreeting" -row (39-$firstrowTeresa) -textfile $textfileTeresa
Add-ContentAtLine -pdfFile $TeresaPath8 -key "HangUp" -row (40-$firstrowTeresa) -textfile $textfileTeresa


#Eric's
$firstrowEric = 3
Add-ContentAtLine -pdfFile $EricPath1 -key "TotalCalls" -row (3-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath1 -key "Key 1" -row (4-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath1 -key "Key 2" -row (11-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath1 -key "Key 3" -row (15-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath1 -key "Key 4" -row (21-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath1 -key "Key 5" -row (28-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath1 -key "Key 6" -row (29-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath1 -key "Key 9" -row (32-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath1 -key "HangUp" -row (33-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath2 -key "Key 1" -row (5-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath2 -key "Key 2" -row (6-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath2 -key "Key 3" -row (7-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath2 -key "Key 4" -row (8-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath2 -key "Key 5" -row (9-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath2 -key "Key 6" -row (10-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath3 -key "Key 1" -row (12-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath3 -key "Key 2" -row (13-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath3 -key "Key 3" -row (14-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath4 -key "Key 1" -row (16-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath4 -key "Key 2" -row (17-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath4 -key "Key 3" -row (18-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath4 -key "Key 4" -row (19-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath4 -key "Key 6" -row (20-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath5 -key "Key 1" -row (22-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath5 -key "Key 2" -row (23-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath5 -key "Key 3" -row (24-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath5 -key "Key 4" -row (25-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath5 -key "Key 5" -row (26-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath5 -key "Key 6" -row (27-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath6 -key "Key 1" -row (30-$firstrowEric) -textfile $textfileEric
Add-ContentAtLine -pdfFile $EricPath6 -key "Key 2" -row (31-$firstrowEric) -textfile $textfileEric

invoke-item $textfileJanet
invoke-item $textfileTeresa
invoke-item $textfileEric
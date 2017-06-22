###############################################################################################
# READ-ME
###############################################################################################
#
# Script Name: transfer_files
# 
###############################################################################################
# Script Description:
# This script transfers the downloaded files into their respetive folders
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

#The folder where the downloads from browser are stored
$folderpath = "\\ykr-fs2\FI_Information_Technology\Administration\A27 - Communication Systems\CallreportAutomation\Temp File Storage"

#User enters the month and year
$month = Read-Host -Promt "Enter the month (Jan, Feb, Mar, Apr, May, June, July, Aug, Sept, Oct, Nov, Dec)"
$year = Read-Host -Promt "Enter the year (Ex: 2017)"

#folder contains the items in the folder specified
$folder = Get-ChildItem $folderpath


$mainpath = "\\ykr-fs2\FI_Information_Technology\Administration\A27 - Communication Systems\Telephone Data"

#Locations where files are copied to
$path1 = $mainpath + "\COURTS\"+$year
$path2 = $mainpath + "\HEALTH REPORTS\HC menu call Counts\"+$year
$path3 = $mainpath + "\IDCD\"+$year
$path4 = $mainpath + "\IDCD\Immunization\" +$year
$path5 = $mainpath + "\EMS\"+$year
$path6 = $mainpath +"\" + $year
$path7 = $mainpath + "\CS&Hdata\Kids Line Counts\" +$year
$path8 = $mainpath + "\HR Reports\" +$year
$path9 = $mainpath + "\LTC\LTC Scheduling\" +$year
$path10 = $mainpath + "\IT REPORTS\" + $year
$path11 = $mainpath + "\ROAD PERMITS\" + $year



#Textfile where pdf is parsed (pdf raw data converted to text in this file)
$temptextfile = $env:UserProfile + "\temptxt.txt"

#Empty the text files, so that previous reports do not remain
function ClearOrMake-files{
    param(
        [Parameter(Mandatory=$true)]
        $file
        )
`   #if there exists a file, clear it, if not, create one (which is empty)
    If (Test-Path $file){Clear-Content $file}
        else{New-item $file -type file}
}



#Copy itextsharp dll file into user
If (!(Test-Path (($env:UserProfile) + "\itextsharp.dll"))){
        copy-item -path "\\ykr-fs2\FI_Information_Technology\Administration\A27 - Communication Systems\CallreportAutomation\itextsharp\itextsharp.dll"
                  -destination "$env:UserProfile"
}

#Function uses isharptext library
function Get-PDFContent {
    param(
        [Parameter(Mandatory=$true)]
        $pdfFile
        ) 
    #Store the userprofile in the variable commandpath       
    $commandPath = $env:UserProfile
    #This command adds the itextsharp class to the session
    Add-Type -Path "$($commandPath)\itextsharp.dll"    

     #Create object that allows reading of the pdffile
    $reader = New-Object iTextSharp.text.pdf.PdfReader $pdfFile
    #How you want pdf to be extraxted
    $strategy = New-Object iTextcharp.text.pdf.parser.SimpleTextExtractionStrategy

    #Extract pdf data from page 1 to the number of pages in the pdf
    for ($i = 1; $i -lt ($reader.NumberOfPages)+1; ++$i) { 
        [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i, $strategy)
    }

    $Reader.Close();
}

#How 
$filenum = ($folder | measure-object).Count



(Get-PDFContent -pdfFile 'C:\Users\chawlas\Documents\CallReportTemp\report(13).pdf'|Out-File $temptextfile)


foreach ($file in $folder){
    ClearOrMake-files -file $temptextfile
    $filepath = $file.FullName
    
   (Get-PDFContent -pdfFile ($filepath)|Out-File $temptextfile)
   if ((get-content $temptextfile)[16] -eq "Key"){
       $filename= (get-content $temptextfile)[16-1]
        }
    else{
    $filename = (get-content $temptextfile)[16]
        }
    if($filename -eq "x73337 -SSC Courts Collection Line"){$filename = $month+ " x73337 -SSC Collection line"}
    if($filename -eq "Tann -Prosecutors -general line"){$filename = $month+ " Tann -Prosecutors -general line"}
    if($filename -eq "SSC-Prosecutors-General-Line"){$filename = $month+ " SSC -Prosecutors -general line"}
    if($filename -eq "HC -IDCD"){$filename = $month+ " HC IDCD"}
    if($filename -eq "HC -IDCD Opt3"){$filename = $month+ " HC IDCD Opt 3"}
    if($filename -eq "HC -Inspectors"){$filename = $month+ " HC -Inspectors"}
    if($filename -eq "HC -Inspectors afterhrs"){$filename = $month+ " HC -Inspectors afterhrs"}
    if($filename -eq "HC -PreRecorded -main"){$filename = $month+ " HC -Prerecorded -main"}
    if($filename -eq "HC -PreRecorded -Opt 1"){$filename = $month+ " HC -Prerecorded -Opt 1"}
    if($filename -eq "HC -PreRecorded -Opt 2"){$filename = $month+ " HC -Prerecorded -Opt 2"}
    if($filename -eq "HC -PreRecorded -Opt 3"){$filename = $month+ " HC -Prerecorded -Opt 3"}
    if($filename -eq "HC -PreRecorded -Opt 4"){$filename = $month+ " HC -Prerecorded -Opt 4"}
    if($filename -eq "HC -PreRecorded -Opt 5"){$filename = $month+ " HC -Prerecorded -Opt 5"}
    if($filename -eq "HC -PreRecorded -Opt 6"){$filename = $month+ " HC -Prerecorded -Opt 6"}
    if($filename -eq "HC -Sexual line"){$filename = $month+ " HC -Sexual line"}
    if($filename -eq "HC-UPD-Main-Menu"){$filename = $month+ " HC -UPD main menu"}
    if($filename -eq "HC-UPD-Main-Menu-After-Hours"){$filename = $month+ " HC-UPD -main menu -afterhrs"}
    if($filename -eq "IDC main-menu"){$filename = $month+ " IDC-main menu"}
    if($filename -eq "IDC Opt 3 menu"){$filename = $month+ " IDC opt 3"}
    if($filename -eq "Immun -main menu"){$filename = $month+ " Imm-main menu"}
    if($filename -eq "EMS Operations -Opt 1"){$filename = $month+ " EMS Opt 1"}
    if($filename -eq "EMS Operations -Opt 2"){$filename = $month+ " EMS Opt 2"}
    if($filename -eq "EMS Operations -Opt 3"){$filename = $month+ " EMS Opt 3"}
    if($filename -eq "EMS Operations -Opt 4"){$filename = $month+ " EMS Opt 4"}
    if($filename -eq "EMS Operations -Opt 5"){$filename = $month+ " EMS Opt 5"}
    if($filename -eq "EMS Operations -Opt 6"){$filename = $month+ " EMS Opt 6"}
    if($filename -eq "EMS Operations -Opt 9"){$filename = $month+ " EMS Opt 9"}
    if($filename -eq "EMS Operations Auto Attendant"){$filename = $month+ " EMS Auto Attendant"}
    if($filename -eq "Admin-Access York main menu -Opt 3"){$filename = $month+ " Admin-Access York main menu -Opt 3"}
    if($filename -eq "Admin-Access York main menu-Opt 1"){$filename = $month+ " Admin-Access York main menu -Opt 1"}
    if($filename -eq "Admin-Access York-Go4York main menu"){$filename = $month+ " Admin- Access York-Go4York main menu"}
    if($filename -eq "Admin-Auto-Attendant main menu"){$filename = $month+ " Auto Attendant"}
    if($filename -eq "Admin-C-HS -Main menu -After hrs"){$filename = $month+ " Admin-C-HS -Main menu -After hrs"}
    if($filename -eq "Admin-CHS CC -Opt 1 -subopt 3 -not taking app"){$filename = $month+ " Admin CHS CC -Opt 1 -not taking app"}
    if($filename -eq "Admin-CHS CC -Opt 1 -subopt 3 -taking app"){$filename = $month+ " Admin CHS CC -Opt 1 -taking app"}
    if($filename -eq "Admin-CHS CC -Opt 1 OW"){$filename = $month+ " Admin CHS CC -Opt 1 OW"}
    if($filename -eq "Admin-CHS CC -Opt 3 -subopt 1 -1 Market Rent"){$filename = $month+ " Admin-CHS CC -Opt 3 -1 Market Rent"}
    if($filename -eq "Admin-CHS CC -Opt 3 -subopt 1 -2 Subsidized Housing"){$filename = $month+ " Admin-CHS CC -Opt 3 -2 Subsidized Housing"}
    if($filename -eq "Admin-CHS CC -Opt 3 -subopt 1 Housing Inq"){$filename = $month+ " Admin-CHS CC -Opt 3 -subopt 1 Housing Inq"}
    if($filename -eq "Admin-CHS CC -Opt 3 -subopt 3 Housing Stability"){$filename = $month+ " Admin CHS CC -Opt 3 -subopt 3 Housing Stability"}
    if($filename -eq "Admin-CHS CC -Opt 3 Housing"){$filename = $month+ " Admin CHS CC -Opt 3 Housing"}
    if($filename -eq "Admin-CHS-CC-After-Hours"){$filename = $month+ " Admin-CHS-CC-After-Hours"}
    if($filename -eq "Admin-CHS-CC-Main-Menu"){$filename = $month+ " Admin CHS CC -main menu"}
    if($filename -eq "Admin-ENVR-Main-Menu"){$filename = $month+ " Admin- ENVR -main menu"}
    if($filename -eq "Admin-UPD-Health-Main-Menu"){$filename = $month+ " UPD Admin- Health -main menu"}
    if($filename -eq "Admin-TRN-Main-Menu"){$filename = $month+ " Admin- TRN -main menu"}
    if($filename -eq "Child Care-Kids Line -main menu"){$filename = $month+ " Child Care line"}
    if($filename -eq "HR-Career-Line x75508 main menu"){$filename = $month+ " HR Career Line"}
    if($filename -eq "1091Gorham-Housing-Main-Menu"){$filename = $month+ " 1091 Gorham- main menu"}
    if($filename -eq "1091Gorham-Housing-Emerg-After-Hours"){$filename = $month+ " 1091 Gorham- Emerg - After Hours"}
    if($filename -eq "24262Woodbine-Main-Menu"){$filename = $month+ " 24262 Woodbine- main menu"}
    if($filename -eq "380 Bayview -main menu"){$filename = $month+ " 380 Bayview -main menu"}
    if($filename -eq "62Bayview-Main-Menu"){$filename = $month+ " 62 Bayview- main menu"}
    if($filename -eq "9060 Jane C-HS main-Menu"){$filename = $month+ " 9060 Jane C-HS- main menu"}
    if($filename -eq "SSC-Auto-Attendant"){$filename = $month+ " SSC- Auto Attendant"}
    if($filename -eq "SSC-CHS-3rd-Flr-Main-Menu"){$filename = $month+ " SSC-CHS-3rd flr -main menu"}
    if($filename -eq "SSC-EIS-4th-Floor-Main-Menu"){$filename = $month+ " SSC-EIS-4th flr -main menu"}
    if($filename -eq "Courts Tann-Eng main menu"){$filename = $month+ " Tann Court-main menu"}
    if($filename -eq "Courts-SSC-Eng-Main-Menu"){$filename = $month+ " SSC Court-main menu"}
    if($filename -eq "LTC Scheduling -after hrs"){$filename = $month+ " LTC Scheduling -after hrs"}
    if($filename -eq "LTC Scheduling -main menu"){$filename = $month+ " LTC Scheduling -main menu"}
    if($filename -eq "LTC Scheduling -opt3"){$filename = $month+ " LTC Scheduling -opt3"}
    if($filename -eq "ITServiceDesk-Afterhour"){$filename = $month+ " IT Servicedesk-Afterhour"}
    if($filename -eq "ITServiceDesk-Afterhour support"){$filename = $month+ " ITServiceDesk-Afterhour support"}
    if($filename -eq "ITServiceDesk-Main menu"){$filename = $month+ " IT Servicedesk-main menu"}
    if($filename -eq "ITServicedesk-Opt 1 to queue"){$filename = $month+ " ITServicedesk-Opt 1 to queue"}
    if($filename -eq "ITServiceDesk-option2"){$filename = $month+ " ITServiceDesk-option2"}
    if($filename -eq "Road Permits Main"){$filename = $month+ " - Road Permits Main"}
    if($filename -eq "Road Permits App -Subopt 1 Sign"){$filename = $month+ " - Road Permits -Subopt 1"}
    if($filename -eq "Road Permits App -Subopt 2 Occupy Road"){$filename = $month+ " - Road Permits -Subopt 2"}
    if($filename -eq "Road Permits App -Subopt 3 Excess Load"){$filename = $month+ " - Road Permits -Subopt 3"}
    if($filename -eq "Road Permits App -Subopt 4 Entrance"){$filename = $month+ " - Road Permits -Subopt 4"}
    $filename = $filename+".pdf"
    
    Rename-Item -path "$filepath" -NewName ($filename)

    $filepath = $file.fullname
    
     if($filename -eq ($month+ " x73337 -SSC Collection line.pdf")){
         If (!(Test-Path ($path1 +"\"+ $month + " x73337 -SSC Collection line"))){
            Copy-Item -Path $filepath -Destination $path1}}
    
     if($filename -eq ($month+ " Tann -Prosecutors -general line.pdf")){
         If (!(Test-Path ($path1+ "\"+$month+ " Tann -Prosecutors -general line.pdf"))){
            Copy-Item -Path $filepath -Destination $path1}}
    
     if($filename -eq ($month+ " SSC -Prosecutors -general line.pdf")){
         If (!(Test-Path ($path1+"\"+$month + " SSC -Prosecutors -general line.pdf"))){
            Copy-Item -Path $filepath -Destination $path1}}
   
     if($filename -eq ($month+ " HC IDCD.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC IDCD.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
    
     if($filename -eq ($month+ " HC IDCD Opt 3.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC IDCD Opt 3.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
    
     if($filename -eq ($month+ " HC -Inspectors.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC -Inspectors.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
    
     if($filename -eq ($month+ " HC -Inspectors afterhrs.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC -Inspectors afterhrs.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
    
     if($filename -eq ($month+ " HC -Prerecorded -main.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC -Prerecorded -main.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
   
     if($filename -eq ($month+ " HC -Prerecorded -Opt 1.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC -Prerecorded -Opt 1.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
    
     if($filename -eq ($month+ " HC -Prerecorded -Opt 2.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC -Prerecorded -Opt 2.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
    
     if($filename -eq ($month+ " HC -Prerecorded -Opt 3.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " EMS Opt 3.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
   
     if($filename -eq ($month+ " HC -Prerecorded -Opt 4.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " EMS Opt 4.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
    
     if($filename -eq ($month+ " HC -Prerecorded -Opt 5.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " EMS Opt 5.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
   
     if($filename -eq ($month+ " HC -Prerecorded -Opt 6.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC -Prerecorded -Opt 6.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
    
     if($filename -eq ($month+ " HC -Sexual line.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC -Sexual line.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
   
     if($filename -eq ($month+ " HC -UPD main menu.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC -UPD main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
    
     if($filename -eq ($month+ " HC-UPD -main menu -afterhrs.pdf")){
         If (!(Test-Path ($path2+"\"+$month + " HC-UPD -main menu -afterhrs.pdf"))){
            Copy-Item -Path $filepath -Destination $path2}}
  
     if($filename -eq ($month+ " IDC-main menu.pdf")){
         If (!(Test-Path ($path3+"\"+$month + " IDC-main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path3}}
    
     if($filename -eq ($month+ " IDC opt 3.pdf")){
         If (!(Test-Path ($path3+"\"+$month + " IDC opt 3.pdf"))){
            Copy-Item -Path $filepath -Destination $path3}}
    
     if($filename -eq ($month+ " Imm-main menu.pdf")){
         If (!(Test-Path ($path4+"\"+$month + " Imm-main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path4}}
    
     if($filename -eq ($month+ " EMS Opt 1.pdf")){
         If (!(Test-Path ($path5+"\"+$month + " EMS Opt 1.pdf"))){
            Copy-Item -Path $filepath -Destination $path5}}
   
     if($filename -eq ($month+ " EMS Opt 2.pdf")){
         If (!(Test-Path ($path5+"\"+$month + " EMS Opt 2.pdf"))){
            Copy-Item -Path $filepath -Destination $path5}}
   
     if($filename -eq ($month+ " EMS Opt 3.pdf")){
         If (!(Test-Path ($path5+"\"+$month + " EMS Opt 3.pdf"))){
            Copy-Item -Path $filepath -Destination $path5}}
   
     if($filename -eq ($month+ " EMS Opt 4.pdf")){
         If (!(Test-Path ($path5+"\"+$month + " EMS Opt 4.pdf"))){
            Copy-Item -Path $filepath -Destination $path5}}
   
     if($filename -eq ($month+ " EMS Opt 5.pdf")){
         If (!(Test-Path ($path5+"\"+$month + " EMS Opt 5.pdf"))){
            Copy-Item -Path $filepath -Destination $path5}}
   
     if($filename -eq ($month+ " EMS Opt 6.pdf")){
         If (!(Test-Path ($path5+"\"+$month + " EMS Opt 6.pdf"))){
            Copy-Item -Path $filepath -Destination $path5}}
   
     if($filename -eq ($month+ " EMS Opt 9.pdf")){
         If (!(Test-Path ($path5+"\"+$month + " EMS Opt 9.pdf"))){
            Copy-Item -Path $filepath -Destination $path5}}
   
     if($filename -eq ($month+ " EMS Auto Attendant.pdf")){
         If (!(Test-Path ($path5+"\"+$month + " EMS Auto Attendant.pdf"))){
            Copy-Item -Path $filepath -Destination $path5}}
   
     if($filename -eq ($month+ " Admin-Access York main menu -Opt 3.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin-Access York main menu -Opt 3.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin-Access York main menu -Opt 1.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin-Access York main menu -Opt 1.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " Admin- Access York-Go4York main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin- Access York-Go4York main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " Auto Attendant.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Auto Attendant.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
    
     if($filename -eq ($month+ " Admin-C-HS -Main menu -After hrs.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin-C-HS -Main menu -After hrs.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin CHS CC -Opt 1 -not taking app.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin CHS CC -Opt 1 -not taking app.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin CHS CC -Opt 1 -taking app.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin CHS CC -Opt 1 -taking app.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin CHS CC -Opt 1 OW.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin CHS CC -Opt 1 OW.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " Admin-CHS CC -Opt 3 -1 Market Rent.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin-CHS CC -Opt 3 -1 Market Rent.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin-CHS CC -Opt 3 -2 Subsidized Housing.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin-CHS CC -Opt 3 -2 Subsidized Housing.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin-CHS CC -Opt 3 -subopt 1 Housing Inq.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin-CHS CC -Opt 3 -subopt 1 Housing Inq.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin CHS CC -Opt 3 -subopt 3 Housing Stability.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin CHS CC -Opt 3 -subopt 3 Housing Stability.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin CHS CC -Opt 3 Housing.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin CHS CC -Opt 3 Housing.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin-CHS-CC-After-Hours.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin-CHS-CC-After-Hours.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin CHS CC -main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin CHS CC -main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " Admin- ENVR -main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin- ENVR -main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " UPD Admin- Health -main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " UPD Admin- Health -main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " Admin- TRN -main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " Admin- TRN -main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " Child Care line.pdf")){
         If (!(Test-Path ($path7+"\"+$month + " Child Care line.pdf"))){
            Copy-Item -Path $filepath -Destination $path7}}
   
     if($filename -eq ($month+ " HR Career Line.pdf")){
         If (!(Test-Path ($path8+"\"+$month + " HR Career Line.pdf"))){
            Copy-Item -Path $filepath -Destination $path8}}
   
     if($filename -eq ($month+ " 1091 Gorham- main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " 1091 Gorham- main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
  
     if($filename -eq ($month+ " 1091 Gorham- Emerg - After Hours.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " 1091 Gorham- Emerg - After Hours.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " 24262 Woodbine- main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " 24262 Woodbine- main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " 380 Bayview -main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " 380 Bayview -main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
    
     if($filename -eq ($month+ " 62 Bayview- main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " 62 Bayview- main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " 9060 Jane C-HS- main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " 9060 Jane C-HS- main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " SSC- Auto Attendant.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " SSC- Auto Attendant.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " SSC-CHS-3rd flr -main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " SSC-CHS-3rd flr -main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " SSC-EIS-4th flr -main menu.pdf")){
         If (!(Test-Path ($path6+"\"+$month + " SSC-EIS-4th flr -main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path6}}
   
     if($filename -eq ($month+ " Tann Court-main menu.pdf")){
         If (!(Test-Path ($path1+"\"+$month + " Tann Court-main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path1}}
  
     if($filename -eq ($month+ " SSC Court-main menu.pdf")){
         If (!(Test-Path ($path1+"\"+$month + " SSC Court-main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path1}}
    
     if($filename -eq ($month+ " LTC Scheduling -after hrs.pdf")){
         If (!(Test-Path ($path9+"\"+$month + " LTC Scheduling -after hrs.pdf"))){
            Copy-Item -Path $filepath -Destination $path9}}
   
     if($filename -eq ($month+ " LTC Scheduling -main menu.pdf")){
         If (!(Test-Path ($path9+"\"+$month + " LTC Scheduling -main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path9}}
    
     if($filename -eq ($month+ " LTC Scheduling -opt3.pdf")){
         If (!(Test-Path ($path9+"\"+$month + " LTC Scheduling -opt3.pdf"))){
            Copy-Item -Path $filepath -Destination $path9}}
   
     if($filename -eq ($month+ " IT Servicedesk-Afterhour.pdf")){
         If (!(Test-Path ($path10+"\"+$month + " IT Servicedesk-Afterhour.pdf"))){
            Copy-Item -Path $filepath -Destination $path10}}
   
     if($filename -eq ($month+ " ITServiceDesk-Afterhour support.pdf")){
         If (!(Test-Path ($path10+"\"+$month + " ITServiceDesk-Afterhour support.pdf"))){
            Copy-Item -Path $filepath -Destination $path10}}
  
     if($filename -eq ($month+ " IT Servicedesk-main menu.pdf")){
         If (!(Test-Path ($path10+"\"+$month + " IT Servicedesk-main menu.pdf"))){
            Copy-Item -Path $filepath -Destination $path10}}
  
     if($filename -eq ($month+ " ITServicedesk-Opt 1 to queue.pdf")){
         If (!(Test-Path ($path10+"\"+$month + " ITServicedesk-Opt 1 to queue.pdf"))){
            Copy-Item -Path $filepath -Destination $path10}}
   
     if($filename -eq ($month+ " ITServiceDesk-option2.pdf")){
         If (!(Test-Path ($path10+"\"+$month + " ITServiceDesk-option2.pdf"))){
            Copy-Item -Path $filepath -Destination $path10}}
   
     if($filename -eq ($month+ " - Road Permits Main.pdf")){
         If (!(Test-Path ($path11+"\"+$month + " - Road Permits Main.pdf"))){
            Copy-Item -Path $filepath -Destination $path11}}
    
     if($filename -eq ($month+ " - Road Permits -Subopt 1.pdf")){
         If (!(Test-Path ($path11+"\"+$month + " - Road Permits -Subopt 1.pdf"))){
            Copy-Item -Path $filepath -Destination $path11}}
   
     if($filename -eq ($month+ " - Road Permits -Subopt 2.pdf")){
         If (!(Test-Path ($path11+"\"+$month + " - Road Permits -Subopt 2.pdf"))){
            Copy-Item -Path $filepath -Destination $path11}}
   
     if($filename -eq ($month+ " - Road Permits -Subopt 3.pdf")){
         If (!(Test-Path ($path11+"\"+$month + " - Road Permits -Subopt 3.pdf"))){
            Copy-Item -Path $filepath -Destination $path11}}
   
     if($filename -eq ($month+ " - Road Permits -Subopt 4.pdf")){
         If (!(Test-Path ($path11+"\"+$month + " - Road Permits -Subopt 4.pdf"))){
            Copy-Item -Path $filepath -Destination $path11}}    #>
}




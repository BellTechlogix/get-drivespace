## BEGGINNING OF SCRIPT ###

#Set execution policy to Unrestricted (-Force suppresses any confirmation)
#Execution policy stopped the script from running via task scheduler
#As a work-around, I added an action in the task scheduler to run first before this script runs
# Set-ExecutionPolicy Unrestricted -Force

Set-ExecutionPolicy Unrestricted -Force

#delete reports older than 7 days

$OldReports = (Get-Date).AddDays(-7)

#edit the line below to the location you store your disk reports# It might also
#be stored on a local file system for example, D:\ServerStorageReport\DiskReport

Get-ChildItem C:\Belltech\DiskReports\*.* | `
Where-Object { $_.LastWriteTime -le $OldReports} | `
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue  

#Create variable for log date

$LogDate = get-date -f yyyyMMddhhmm


#Define location of text file containing your servers. It might also
#be stored on a local file system for example, D:\ServerStorageReport\DiskReport

#$File = (import-csv C:\belltech\DiskReports\CrowleyWSUSPlan.csv).DeviceHostname1
$File =  get-adcomputer -Filter {OperatingSystem -like "*Server*"} -Properties OperatingSystem|select Name
#Now Set your Output File Locations, use your temp folder
$File1 = "$env:TEMP\diskspace.xml"
$File2 = "$env:TEMP\diskspace.xlsx"
#Discard Old Copies
Remove-Item $file1,$file2



#Define admin account variable (Uncommented it and the line in Get-WmiObject
#Line 44 below is commented out because I run this script via task schedule.
#If you wish to run this script manually, please uncomment line 44.

$RunAccount = get-Credential  

# $DiskReport = ForEach ($Servernames in ($File)) 

#the disk $DiskReport variable checks all servers returned by the $File variable (line 37)

$DiskReport = ForEach ($Servernames in ($file.Name)) 
{Get-WmiObject win32_logicaldisk -Credential $RunAccount -ComputerName $Servernames -Filter "Drivetype=3" -ErrorAction SilentlyContinue|

#return only C: Drives
where-object{$_.DeviceID -eq 'C:'}
$servernames
#return only disks with
#free space less  
#than or equal to 0.1 (10%)

#Where-Object {   ($_.freespace/$_.size) -le '0.1'}

} 


#create reports

#Lets create our XML File, this is the initial formatting that it will need to understand what it is, and what styles we are using.
(
 '<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Kristopher Roy</Author>
  <LastAuthor>'+$env:USERNAME+'</LastAuthor>
  <Created>'+(get-date)+'</Created>
   <Version>16.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>8145</WindowHeight>
  <WindowWidth>20490</WindowWidth>
  <WindowTopX>0</WindowTopX>
  <WindowTopY>0</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s62">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
    ss:Bold="1"/>
  </Style>
  <Style ss:ID="s63">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
  </Style>
  <Style ss:ID="s70">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0%"/>
  </Style>
    <Style ss:ID="s65">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Interior ss:Color="#FFFF00" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0%"/>
  </Style>
    <Style ss:ID="s69">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Interior ss:Color="#00B050" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="0%"/>
  </Style>
  <Style ss:ID="s66">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
    ss:Bold="1"/>
   <Interior ss:Color="#D0CECE" ss:Pattern="Solid"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="DiskReport">
  <Table ss:ExpandedColumnCount="5" ss:ExpandedRowCount="'+($Number+1)+'" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:Width="65.25"/>
   <Column ss:Width="60"/>
   <Column ss:Width="93.75"/>
   <Column ss:Width="78.75"/>
   <Column ss:Width="95.25"/>
   <Row ss:AutoFitHeight="0" ss:Height="15.75">
    <Cell ss:StyleID="s66"><Data ss:Type="String">Server Name</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Drive Letter</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Total Capacity (GB)</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Free Space (GB)</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Free Space Percent</Data></Cell>
   </Row>')> $file1


$NewDiskReport = ($DiskReport |
Select-Object @{Label = "Server Name";Expression = {$_.SystemName}},
@{Label = "Drive Letter";Expression = {$_.DeviceID}},
@{Label = "Total Capacity (GB)";Expression = {"{0:N1}" -f( $_.Size / 1gb)}},
@{Label = "Free Space (GB)";Expression = {"{0:N1}" -f( $_.Freespace / 1gb ) }},
@{Label = 'Free Space Percent'; Expression = {($_.freespace/$_.size)}})|where{$_.'Server Name' -ne $null}


$Number = $NewDiskReport.'Server Name'.count

#FOREACH($Report in $NewDiskReport){$Report.'Free Space Percent'}

#$report = $NewDiskReport|where{$_."Server Name" -eq "JAXXEN65APP12"}

FOREACH($Report in $NewDiskReport)
{
#I wanted any sites with Prod in them to be highlighted, you can change this to whatever you like
If($Report.'Free Space Percent' -eq $null -or $Report.'Free Space Percent' -eq ""){$Report.'Free Space Percent' = ($Report."Free Space (GB)"/$Report."Total Capacity (GB)")}
$Report.'Free Space Percent' = [math]::Round($Report.'Free Space Percent',2)
add-content $File1 ('<Row ss:AutoFitHeight="0">')
add-content $File1 ('<Cell ss:StyleID="s63"><Data ss:Type="String">'+($Report.'Server Name')+'</Data></Cell>')
add-content $File1 ('<Cell ss:StyleID="s63"><Data ss:Type="String">'+($Report.'Drive Letter')+'</Data></Cell>')
add-content $File1 ('<Cell ss:StyleID="s63"><Data ss:Type="Number">'+($Report.'Total Capacity (GB)')+'</Data></Cell>')
add-content $File1 ('<Cell ss:StyleID="s63"><Data ss:Type="Number">'+($Report.'Free Space (GB)')+'</Data></Cell>')
IF($Report.'Free Space Percent'-le '0.10'){add-content $File1 ('<Cell ss:StyleID="s70"><Data ss:Type="Number">'+($Report.'Free Space Percent')+'</Data></Cell>')}
ELSEIF($Report.'Free Space Percent'-ge '0.111' -and $Report.'Free Space Percent'-le '0.50'){add-content $File1 ('<Cell ss:StyleID="s65"><Data ss:Type="Number">'+($Report.'Free Space Percent')+'</Data></Cell>')}
ELSEIF($Report.'Free Space Percent'-ge '0.51'){add-content $File1 ('<Cell ss:StyleID="s69"><Data ss:Type="Number">'+($Report.'Free Space Percent')+'</Data></Cell>')}
add-content $File1 '</Row>'
}


#Close out the XML
add-content $file1 (
' </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <Print>
    <ValidPrinterInfo/>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>6</ActiveRow>
     <ActiveCol>4</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>')

#Export report to CSV file (Disk Report)

Export-Csv -path "c:\belltech\DiskReports\DiskReport_$logDate.csv" -NoTypeInformation



#Send disk report using the exchange email module


Import-Module "\\Servername\ServerStorageReport\ExchangeModule\Exchange.ps1" -ErrorAction SilentlyContinue

# Attach and send CSV report (Most recent report will be attached)

$messageParameters = @{                        
                Subject = "Weekly Server Storage Report"                        
                Body = "Attached is Weekly Server Storage Report. The scipt has been amended to return only servers with free disk space less than or equal to 10%. All reports are located in \\Servername\ServerStorageReport\DiskReport\, but the most recent  is sent weekly"                   
                From = "Email name1 <storagereport@crowley.com>"                        
                To = "Email name1 <kroy@belltechlogix.com>"
                CC = "Email name2 <Email.name2@domainname.com>"
                Attachments = (Get-ChildItem c:\belltech\DiskReports\*.* | sort LastWriteTime | select -last 1)                   
                SmtpServer = "webmail.crowley.com"                        
            }   
Send-MailMessage @messageParameters -BodyAsHtml             

##Important note#

# The SendMail portion is not quite complete

#Convert to Excel
$objExcel = new-object -comobject excel.application
$UserWorkBook = $objExcel.Workbooks.Open($file1)
$UserWorkBook.SaveAs($file2,51)
$UserWorkBook.Close()
copy $file2 C:\belltech\

## END OF SCRIPT ### 

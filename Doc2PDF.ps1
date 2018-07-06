#############################################
# Doc2PDF                                   #
# Created: April 30, 2016                   #
# Last Modified: July 07, 2018              #
# Version: 1.0                              #
# Supported Office: 2010*, 2013, 2016       #
# Supported PowerShell: 4, 5                #
# Created By: Erick Scott Johnson           #
# Forked/Updated By: Thayumaanavan C R      # 
#############################################
#Input
$Input = $args[0]
#Thayu - added extra args & prefix
$xcode = $args[1]
$appName = $args[2]
$prefix = $xcode + "_" + $appName
#

#Define Office Formats
$Wrd_Array = '*.docx', '*.doc', '*.odt', '*.rtf', '*.txt', '*.wpd'
$Exl_Array = '*.xlsx', '*.xls', '*.ods', '*.csv'
$Pow_Array = '*.pptx', '*.ppt', '*.odp'
$Pub_Array = '*.pub'
$Vis_Array = '*.vsdx', '*.vsd', '*.vssx', '*.vss'
$Off_Array = $Wrd_Array + $Exl_Array + $Pow_Array + $Pub_Array + $Vis_Array
$ExtChk    = [System.IO.Path]::GetExtension($Input)

#Convert Word to PDF
Function Wrd-PDF($f, $p)
{
    $Wrd     = New-Object -ComObject Word.Application
    $Version = $Wrd.Version
    $Doc     = $Wrd.Documents.Open($f)

    #Check Version of Office Installed
    If ($Version -eq '16.0' -Or $Version -eq '15.0') {
        $Doc.SaveAs($p, 17) 
        $Doc.Close($False)
    }
    ElseIf ($Version -eq '14.0') {
        $Doc.SaveAs([ref] $p,[ref] 17)
        $Doc.Close([ref]$False)
    }
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    $Wrd.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Wrd)
    Remove-Variable Wrd
}

#Convert Excel to PDF
Function Exl-PDF($f, $p)
{
    $Exl = New-Object -ComObject Excel.Application
    $Doc = $Exl.Workbooks.Open($f)
    $Doc.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $p)
    $Doc.Close($False)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    $Exl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Exl)
    Remove-Variable Exl
}

#Convert PowerPoint to PDF
Function Pow-PDF($f, $p)
{
    $Pow = New-Object -ComObject PowerPoint.Application
    $Doc = $Pow.Presentations.Open($f, $True, $True, $False)
    $Doc.SaveAs($p, 32)
    $Doc.Close()
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    $Pow.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Pow)
    Remove-Variable Pow
}

#Convert Publisher to PDF
Function Pub-PDF($f, $p)
{
    $Pub = New-Object -ComObject Publisher.Application
    $Doc = $Pub.Open($f)
    $Doc.ExportAsFixedFormat([Microsoft.Office.Interop.Publisher.PbFixedFormatType]::pbFixedFormatTypePDF, $p)
    $Doc.Close()
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    $Pub.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Pub)
    Remove-Variable Pub
}

#Convert Visio to PDF
Function Vis-PDF($f, $p)
{
    $Vis = New-Object -ComObject Visio.Application
    $Doc = $Vis.Documents.Open($f)
    $Doc.ExportAsFixedFormat([Microsoft.Office.Interop.Visio.VisFixedFormatType]::xlTypePDF, $p)
    $Doc.Close()
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    $Vis.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Vis)
    Remove-Variable Vis
}

#Check for Word Formats
Function Wrd-Chk($f, $e, $p){
    $f = [string]$f
    For ($i = 0; $i -le $Wrd_Array.Length; $i++) {
        $Temp = [string]$Wrd_Array[$i]
        $Temp = $Temp.TrimStart('*')
        If ($e -eq $Temp) {
            Wrd-PDF $f $p
        }
    }
}

#Check for Excel Formats
Function Exl-Chk($f, $e, $p){
    $f = [string]$f
    For ($i = 0; $i -le $Exl_Array.Length; $i++) {
        $Temp = [string]$Exl_Array[$i]
        $Temp = $Temp.TrimStart('*')
        If ($e -eq $Temp) {
            Exl-PDF $f $p
        }
    }
}

#Check for PowerPoint Formats
Function Pow-Chk($f, $e, $p){
    $f = [string]$f
    For ($i = 0; $i -le $Pow_Array.Length; $i++) {
        $Temp = [string]$Pow_Array[$i]
        $Temp = $Temp.TrimStart('*')
        If ($e -eq $Temp) {
            Pow-PDF $f $p
        }
    }
}

#Check for Publisher Formats
Function Pub-Chk($f, $e, $p){
    $f = [string]$f
    For ($i = 0; $i -le $Pub_Array.Length; $i++) {
        $Temp = [string]$Pub_Array[$i]
        $Temp = $Temp.TrimStart('*')
        If ($e -eq $Temp) {
            Pub-PDF $f $p
        }
    }
}

#Check for Visio Formats
Function Vis-Chk($f, $e, $p){
    $f = [string]$f
    For ($i = 0; $i -le $Vis_Array.Length; $i++) {
        $Temp = [string]$Vis_Array[$i]
        $Temp = $Temp.TrimStart('*')
        If ($e -eq $Temp) {
            Vis-PDF $f $p
        }
    }
}

#Check if input is file or directory
If ($ExtChk -eq '')
{
    $Files = Get-ChildItem -path $Input -include $Off_Array -recurse
    ForEach ($File in $Files) {
        $Path     = [System.IO.Path]::GetDirectoryName($File)		
        $Filename = [System.IO.Path]::GetFileNameWithoutExtension($File)		
        $Ext      = [System.IO.Path]::GetExtension($File)
        $PDF      = $Path + '\' + $Filename + '.pdf'
        Wrd-Chk $File $Ext $PDF
        Exl-Chk $File $Ext $PDF
        Pow-Chk $File $Ext $PDF
        Pub-Chk $File $Ext $PDF
        Vis-Chk $File $Ext $PDF
		#Thayu - added rename
		$newName = $Path + '\' +$prefix + '_' + $Filename + '.pdf'
		Rename-Item $PDF $newName
		if(!(Test-Path $prefix)){
			$folder = New-Item -type directory -name $prefix
		}		
		Move-Item $newName -destination $folder.FullName
		#
    }
}
Else
{
    $File     = $Input
    $Path     = [System.IO.Path]::GetDirectoryName($File)
    $Filename = [System.IO.Path]::GetFileNameWithoutExtension($File)
    $Ext      = [System.IO.Path]::GetExtension($File)
    $PDF      = $Path + '\' + $Filename + '.pdf'
    Wrd-Chk $File $Ext $PDF
    Exl-Chk $File $Ext $PDF
    Pow-Chk $File $Ext $PDF
    Pub-Chk $File $Ext $PDF
    Vis-Chk $File $Ext $PDF
	#Thayu - added rename
	$newName = $Path + '\' +$prefix + ' ' + $Filename + '.pdf'
	Rename-Item $PDF $newName
	if(!(Test-Path $prefix)){
			$folder = New-Item -type directory -name $prefix
	}	
	Move-Item $newName -destination $folder.FullName
	#
}

#Thayu - zip the files
$FolderPath = $folder.FullName + '\*'
$ZipFileName = $prefix + '.zip'
Compress-Archive -Path $FolderPath -DestinationPath $ZipFileName

#Cleanup
Remove-Item Function:Wrd-PDF, Function:Wrd-Chk
Remove-Item Function:Exl-PDF, Function:Exl-Chk
Remove-Item Function:Pow-PDF, Function:Pow-Chk
Remove-Item Function:Pub-PDF, Function:Pub-Chk
Remove-Item Function:Vis-PDF, Function:Vis-Chk

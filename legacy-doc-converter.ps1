function Select-FileDialog
{
	param([string]$Description,[string]$Directory,[string]$Filter="All Files (*.*)|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	$objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    	$objForm.RootFolder = "Desktop"
    	$objForm.Description = $Description
	$Show = $objForm.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $objForm.SelectedPath
	}
	Else
	{
		Write-Error "Operation cancelled by user."
	}
}

$containingDir = Select-FileDialog -Description "Pick containing folder of legacy documents..."

[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
$wordInstance = new-object -ComObject Word.Application
$wordInstance.Visible = $False

$filesToConvert = Get-ChildItem $containingDir | where{$_.Extension -eq ".doc" } 

if($filesToConvert){
    write-host Found $filesToConvert.Count legacy doc(s) in: $containingDir -ForegroundColor Green
    forEach($file in $filesToConvert) {
            write-host "Converting:" $file.fullname -ForegroundColor Green 
            [ref]$name = Join-Path -Path $file.DirectoryName -ChildPath $($file.BaseName + ".docx")
            $opendoc = $wordInstance.documents.open($file.FullName)
            $opendoc.saveas([ref]$name.value, [ref]$saveFormat::wdFormatDocument)
            [ref]$saveFormat::wdFormatDocument
            $opendoc.close()
            $file = $null
        }

    $wordInstance.quit()

    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$wordInstance)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable wordInstance

write-host All Done! Closing in 15 seconds! -ForegroundColor Green 
sleep -seconds 15
}
else{
    write-host No legacy files detected in $containingDir
    sleep -seconds 15
}

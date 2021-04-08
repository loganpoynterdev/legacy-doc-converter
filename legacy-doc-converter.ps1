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

$containingDir = Select-FileDialog -Description "Pick A Folder"

[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
$1word = new-object -ComObject Word.Application
$1word.Visible = $False

$filesToConvert = Get-ChildItem $containingDir | where{$_.Extension -eq ".doc" } 
if($filesToConvert){
write-host Found $filesToConvert.Count word 97 docs in: $containingDir -ForegroundColor Green
forEach($EV in $filesToConvert) {
        write-host "Converting :" $EV.fullname -ForegroundColor Green 
        [ref]$name = Join-Path -Path $EV.DirectoryName -ChildPath $($EV.BaseName + ".docx")
        $opendoc = $1word.documents.open($EV.FullName)
        $opendoc.saveas([ref]$name.Value, [ref]$saveFormat::wdFormatDocument)
        $opendoc.saveas([ref]$name.Value, [ref]$SaveFormat::wdFormatDocument)
        [ref]$saveFormat::wdFormatDocument
        $opendoc.close()
        $EV = $null
    }

write-host Doing cleanup -ForegroundColor Green `n `n `n
get-childitem $containingDir | where{$_.extension -eq ".doc" } | remove-item
Set-location $containingDir
attrib +U -P /S

$1word.quit()

    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$1word)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable 1word

write-host All Done! -ForegroundColor Green 

write-host Closing in 15 seconds! -ForegroundColor Yellow
sleep -seconds 15
}
else{
    write-host No word 97 files detected in $containingDir
    sleep -Seconds 15
}

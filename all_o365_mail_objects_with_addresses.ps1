
<#
.SYNOPSIS
    Export all O365 mail-enabled objects with SMTP addresses to Excel.

.DESCRIPTION
    Connects to Exchange Online using certificate-based auth.
    Supports full export or single SMTP address search.
    Outputs to Excel using memory-efficient chunked writing.

.NOTES
    Author: Christopher Horting
    Requirements: ExchangeOnlineManagement module
    Optional: 	  Locally-installed Excel (will default to CSV if not)
                  App Registration with user.read Graph API permission
#>


$error.clear();cls
$ErrorActionPreference = 'Continue'


$appId = '<Your-AppId-Here>'
$thumbPrint = '<Your-App-Secret>'
$org = '<your-tenant>.onmicrosoft.com'

function Quit {
    Disconnect-ExchangeOnline -Confirm:$false
	if ($xlsx) {
        $xlsx.DisplayAlerts = $true
		$xlsx.Quit()
	}
    $global:error.clear()
    if(!($hostinvocation) -and !($script:MyInvocation.MyCommand.Path)) {
        $host.enternestedprompt()
    }
    Exit
}

[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

$addons = @('ExchangeOnlineManagement')
foreach ($addon in $addons) {
	if (Get-Module -ListAvailable -Name $addon) {
		if (!(Get-Module -Name $addon -ErrorAction Silentlycontinue)) {
			Import-Module -Name $addon
		}
	} else {
		[Microsoft.VisualBasic.Interaction]::MsgBox("$($addon) could not be imported. Script will now quit.",'OKOnly,Critical,MsgBoxSetForeground','Unable to Load Addon') | Out-Null
		Quit
	}
}

if ($appId -and $thumbprint -and $org) {
	Connect-ExchangeOnline -AppId $appId -CertificateThumbprint $thumbprint -Organization $org -ShowBanner:$false
} else {
	Connect-ExchangeOnline -ShowBanner:$false
}
if ($error) {
	[Microsoft.VisualBasic.Interaction]::MsgBox('Error Connecting to Exchange Online. Script will now quit.','OKOnly,Critical,MsgBoxSetForeground','Unable to Load Addon') | Out-Null
	Quit
}

do {
	$objForm = New-Object System.Windows.Forms.Form
	$objForm.Text = 'Action Type'
	$objForm.Size = New-Object System.Drawing.Size(250,160)
	$objForm.StartPosition = 'CenterScreen'
	$objForm.KeyPreview = $True
	$objForm.Add_KeyDown({if ($_.KeyCode -eq 'Enter')
	    {$script:actionType = $objDropDown.SelectedItem;$objForm.Close()}})
	$objForm.Add_KeyDown({if ($_.KeyCode -eq 'Escape')
	    {$objForm.Close()}})
	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size(42,90)
	$OKButton.Size = New-Object System.Drawing.Size(75,23)
	$OKButton.Text = 'OK'
	$OKButton.Add_Click({$script:actionType = $objDropDown.SelectedItem;$objForm.Close()})
	$objForm.Controls.Add($OKButton)
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(117,90)
	$CancelButton.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton.Text = 'Cancel'
	$CancelButton.Add_Click({$script:actionType = 'Cancel';$objForm.Close()})
	$objForm.Controls.Add($CancelButton)
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(15,20)
	$objLabel.Size = New-Object System.Drawing.Size(190,30)
	$objLabel.Text = 'Please choose the type of action you wish to take:'
	$objForm.Controls.Add($objLabel)
	$objDropDown = new-object System.Windows.Forms.ComboBox
	$choices = @('Single SMTP Search','All SMTP Addresses')
	foreach ($choice in $choices) {
		$objDropDown.Items.Add($choice) | Out-Null
	}
	$objDropDown.SelectedItem = $choices[0]
	$objDropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	$objDropDown.Location = new-object System.Drawing.Size(70,55)
	$objDropDown.Size = new-object System.Drawing.Size(134,30)
	$objForm.Controls.Add($objDropDown)
	$objForm.Topmost = $True
	$objForm.Add_Shown({$objForm.Activate()})
	[void] $objForm.ShowDialog()
	if (!$actionType) {
		$retry = [Microsoft.VisualBasic.Interaction]::MsgBox('Your input was empty. Do you want to try again?','YesNo,Question', 'Retry')
		if ($retry -eq 'no') {
			Quit
		}
	} elseif ($actionType -eq 'Cancel') {
		Quit
	} else {
		if ($actionType -eq 'Single SMTP Search') {
			$actionType = 1
		} elseif ($actionType -eq 'All SMTP Addresses') {
			$actionType = 2
		}
	}
} until ($actionType)


if ($actionType -eq 1) {
	do {
		$target = ([Microsoft.VisualBasic.Interaction]::InputBox('Which SMTP do you want to search for?','Address Entry','')) -replace '\s+', ''
		if (!$target) {
			$retry = [Microsoft.VisualBasic.Interaction]::MsgBox('Your input was empty. Do you want to try again?','YesNo,Question,MsgBoxSetForeground','Retry')
			if ($retry -eq 'no') {
				Quit
			}
        } elseif ($target -notlike '*@*.*') {
			$retry = [Microsoft.VisualBasic.Interaction]::MsgBox('Your entry does not match SMTP format. Do you want to try again?','YesNo,Question,MsgBoxSetForeground','Retry')
			if ($retry -eq 'no') {
				Quit
			} else {
                $target = $null
            }
		}
	} until ($target)
} elseif ($actionType -eq 2) {
    if($hostinvocation -ne $null) {
	    $dir = Split-Path $hostinvocation.MyCommand.path
    } else {
	    $dir = Split-Path $script:MyInvocation.MyCommand.Path
    }
    $xlsx = New-Object -ComObject 'Excel.Application'
    do {
	    if ($xlsx) {
		    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
		    $SaveDialog.filter = 'Excel Workbook (*.xls;*.xlsx)|*.xlsx'
		    $SaveDialog.Title = 'Select a destination Excel file'
		    $SaveDialog.FilterIndex = 0
		    if ($sourcedir) {
			    $SaveDialog.initialDirectory = $sourcedir
		    } else {
			    $SaveDialog.initialDirectory = $dir
		    }
		    $SaveDialog.RestoreDirectory = $true
		    $SaveDialog.ValidateNames = $true
		    $SaveDialog.ShowDialog() | Out-Null
		    $SaveDialog.ShowHelp = $true
		    if (!$SaveDialog.FileName) {
			    $retry = [Microsoft.VisualBasic.Interaction]::MsgBox('You must provide a file to save the results. Do you want to try again?','YesNo,Question','Retry')
			    if ($retry -eq 'no') {
				    Quit
			    }
		    } else {
			    $dest = $SaveDialog.FileName
		    }
	    } else {
		    $error.clear()
		    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
		    $SaveDialog.filter = 'CSV (Comma delimited) (*.csv)|*.csv'
		    $SaveDialog.Title = 'Save As'
		    $SaveDialog.FilterIndex = 0
		    if ($sourcedir) {
			    $SaveDialog.initialDirectory = $sourcedir
		    } else {
			    $SaveDialog.initialDirectory = $dir
		    }
		    $SaveDialog.RestoreDirectory = $true
		    $SaveDialog.ValidateNames = $true
		    $SaveDialog.ShowDialog() | Out-Null
		    $SaveDialog.ShowHelp = $true
		    if (!$SaveDialog.FileName) {
			    $retry = [Microsoft.VisualBasic.Interaction]::MsgBox('You must provide a file to save the results. Do you want to try again?','YesNo,Question','Retry')
			    if ($retry -eq 'no') {
				    Quit
			    }
		    } else {
			    $dest = $SaveDialog.FileName
		    }
	    }
    } until ($dest)
} else {
    Quit
}

$results = @()
$results += Get-Mailbox -ResultSize Unlimited -IncludeInactiveMailbox | Select @{Name='Type';Expression={'Mailbox'}},Alias,DisplayName,@{Name='InactiveMailbox';Expression={if ($_.IsInactiveMailbox) {'True'} else {'False'}}},PrimarySmtpAddress,@{Name='EmailAddresses';Expression={($_.EmailAddresses | ? {$_ -clike "smtp*"} | ForEach-Object {$_ -replace "smtp:",""}) -join ','}}
$results += Get-DistributionGroup -ResultSize Unlimited | Select @{Name='Type';Expression={'Distribution Group'}},Alias,DisplayName,@{Name='InactiveMailbox';Expression={'N/A'}},PrimarySmtpAddress,@{Name='EmailAddresses';Expression={($_.EmailAddresses | ? {$_ -clike "smtp*"} | ForEach-Object {$_ -replace "smtp:",""}) -join ','}}
$results += Get-DynamicDistributionGroup -ResultSize Unlimited | Select @{Name='Type';Expression={'Dynamic Distribution Group'}},Alias,DisplayName,@{Name='InactiveMailbox';Expression={'N/A'}},PrimarySmtpAddress,@{Name='EmailAddresses';Expression={($_.EmailAddresses | ? {$_ -clike "smtp*"} | ForEach-Object {$_ -replace "smtp:",""}) -join ','}}
$results += Get-UnifiedGroup -ResultSize Unlimited | Select @{Name='Type';Expression={'Unified Group'}},Alias,DisplayName,@{Name='InactiveMailbox';Expression={'N/A'}},PrimarySmtpAddress,@{Name='EmailAddresses';Expression={($_.EmailAddresses | ? {$_ -clike "smtp*"} | ForEach-Object {$_ -replace "smtp:",""}) -join ','}}
$results += Get-Contact -ResultSize Unlimited | ? {$_.RecipientType -ne 'Contact'} | Select @{Name='Type';Expression={if ($_.RecipientType -eq 'MailContact') {'Mail Contact'} else {'Unknown Contact'}}},DisplayName,@{Name='InactiveMailbox';Expression={'N/A'}},@{Name='PrimarySmtpAddress';Expression={$_.WindowsEmailAddress}}

if ($actionType -eq 1) {
    $results = $results | ? {$_.PrimarySmtpAddress -eq $target -or $_.EmailAddresses -contains $target}
}

function Get-ExcelColumnLetter {
    param ([int]$colNumber)
    $letter = ''
    while ($colNumber -gt 0) {
        $colNumber--
        $letter = [char](65 + ($colNumber % 26)) + $letter
        $colNumber = [math]::Floor($colNumber / 26)
    }
    return $letter
}

function Write-ChunkToExcel {
    param (
        [object]$worksheet,
        [array]$data,
        [int]$startRow,
        [int]$totalCols
    )
    $rowsInChunk = $data.Count
    $chunkArray = New-Object 'object[,]' $rowsInChunk, $totalCols
    for ($r = 0; $r -lt $rowsInChunk; $r++) {
        for ($c = 0; $c -lt $totalCols; $c++) {
            $chunkArray[$r, $c] = $data[$r][$c]
        }
    }
    $rangeEndRow = $startRow + $rowsInChunk - 1
    $rangeEndCol = Get-ExcelColumnLetter $totalCols
    $rangeAddress = "A$($startRow):$($rangeEndCol)$($rangeEndRow)"
    $range = $worksheet.Range($rangeAddress)
    if ($range -ne $null) {
        try {
            $range.Value2 = $chunkArray
        } catch {
            Add-Content -Path "$($dir)\excel_write_errors.log" -Value "Memory error writing range $($rangeAddress): $($_)"
        }
    } else {
        Add-Content -Path "$($dir)\excel_write_errors.log" -Value "Invalid range object for $($rangeAddress)"
    }
}


if (!$results) {
    if ($actionType -eq 1) {
        [Microsoft.VisualBasic.Interaction]::MsgBox('No Matches Found','OKOnly,Information,MsgBoxSetForeground','No Results') | Out-Null
    } elseif ($actionType -eq 2) {
        [Microsoft.VisualBasic.Interaction]::MsgBox('Error - No Results','OKOnly,Information,MsgBoxSetForeground','No Results') | Out-Null
    }
} else {
    if ($actionType -eq 1) {
        [Microsoft.VisualBasic.Interaction]::MsgBox("$($results.Alias) | $($results.DisplayName) ($($results.Type))",'OKOnly,Information,MsgBoxSetForeground','Results') | Out-Null
    } elseif ($actionType -eq 2) {
	    $results = $results | Sort Type,DisplayName
	    if ($xlsx) {
		    $xlsx.Visible = $false
		    $xlsx.DisplayAlerts = $false
		    $wb = $xlsx.Workbooks.Add()
            $sheets = ($wb.Sheets | Select Name).Name
		    $ws1 = $wb.Sheets.Add()
            $headers = 'Type','Alias','DisplayName','Inactive','PrimarySmtpAddress','EmailAddresses'
            $ws1.name = 'SMTP Addresses'
            for ($i = 0; $i -lt $headers.Count; $i++) {
                $ws1.Cells.Item(1, $i + 1) = $headers[$i]
                $ws1.Cells.Item(1, $i + 1).Font.Bold = $true
            }
            $chunkSize = 500
            $totalCols = $headers.Count
            $dataRows = @()
            foreach ($item in $results) {
                $dataRows += ,@(
                    [string]$item.Type,
                    [string]$item.Alias,
                    [string]$item.DisplayName,
                    [string]$item.InactiveMailbox,
                    [string]$item.PrimarySmtpAddress,
                    [string]$item.EmailAddresses
                )
            }
            for ($start = 0; $start -lt $dataRows.Count; $start += $chunkSize) {
                $chunk = $dataRows[$start..([math]::Min($start + $chunkSize - 1, $dataRows.Count - 1))]
                Write-ChunkToExcel -worksheet $ws1 -data $chunk -startRow ($start + 2) -totalCols $totalCols
            }
            $ws1.activate()
            $xlsx.Application.ActiveWindow.SplitColumn = 0
            $xlsx.Application.ActiveWindow.SplitRow = 1
            $xlsx.Application.ActiveWindow.FreezePanes = $true
            $ws1.Columns.Item("A:$(Get-ExcelColumnLetter $totalCols)").EntireColumn.AutoFit() | Out-Null
            $ws1.Rows.Item("1:$($totalRows + 1)").EntireRow.AutoFit() | Out-Null
            foreach ($ws in $wb.sheets | ? {$sheets -contains $_.Name}) {$ws.delete()}
            $ws1.Activate()
            $wb.SaveAs($dest)
            $wb.Close()
	    } else {
		    $results | Select Type,Alias,DisplayName,PrimarySmtpAddress,EmailAddresses | Export-Csv $dest -NoTypeInformation
	    }
	    [Microsoft.VisualBasic.Interaction]::MsgBox('Complete','OKOnly,Information,MsgBoxSetForeground','End of Script') | Out-Null
    }
}

Quit

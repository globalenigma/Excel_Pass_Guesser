#bypass execution policy
#Get-Content .Excel_PS_guess.ps1 | PowerShell.exe -noprofile - 
#PowerShell.exe -ExecutionPolicy Bypass -File .Excel_PS_guess.ps1


#Function to open files
Function Get-FileName([string]$file_filter,$initialDirectory){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = $file_filter
    $OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

#Get passwords from file
Function load_passwords{
    $global:passfile = Get-FileName("TXT (*.txt)| *.txt")
    $global:password = Get-Content -Path $passfile
}

#Load following Excel file
Function load_excel{
    $global:xl_file = Get-FileName("XLS (*.xls) or XLSX (*.xlsx)| *.xls*")
}

#Create new Object for Excel Document
Function create_xl_object{
    $global:xl = New-Object -com Excel.Application
    $xl.visible=$False
}

#Try to open Excel document with passwords provided
Function brute_force{
	foreach ($pass in $password){
		try{
			$xl.Workbooks.open($xl_file,1,$false,5,$pass)
			write-host -ForegroundColor Green [$pass] is the correct password for $xl_file
			break
		}catch [System.Runtime.InteropServices.COMException] {
			#find what to catch >> $Error[0].Exception.GetType().FullName
			write-host [$pass] is not the correct password for $xl_file
		}
	}
$decide_cleanup = Read-Host "Would you like to clean up now? [y/n]"
if ($decide_cleanup -eq 'y'){cleanup}
}

#write out current configs
Function write_configs{
    write-host -ForegroundColor Yellow "####################"
    write-host "Excel File Loaded:" $xl_file
    write-host -ForegroundColor Yellow "####################"
	write-host "Password File Loaded:" $passfile
    write-host -ForegroundColor Yellow "####################"
	write-host "Passwords Loaded:" $password
	pause
}

Function cleanup{
	#$xl.created_object.close()
	#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objectname) | Out-Null
	$xl.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
	#if user does not clean up after each run use this to clean  Get-Process EXCEL | Stop-Process
}

#Where the 
function write_menu{
    cls
    $menu_header = '
	            Excel Password Guesser
				                              v.01'
	$menu_selection_text = '
	Main Menu:
	1. Load Passwords from file (one per line)
	2. Choose Excel file to guess
	3. Show configs
	4. Execute Password Guessing
	5. Clean up started Excel processes
	6. Clean up all Excel processes (Caution: stops all Excel Processes)
	7. Exit the program
	'
	write-host -ForegroundColor Green $menu_header
	write-host $menu_selection_text
}

#The main menu worker
function menu{
	do{
		write_menu
		$input = Read-Host "Please make a selection"
		switch ($input){
			'1' {cls;load_passwords}
			'2' {cls;load_excel}
			'3' {cls;write_configs}
			'4' {cls;create_xl_object;brute_force}
			'5' {cls;cleanup}
			'6' {Get-Process EXCEL | Stop-Process}
			'7' { return }
		}
	}
	until ($input -eq '7')
}


menu

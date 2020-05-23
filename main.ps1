<#
┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ INVENTORY SCRIPT                                                                            │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤
│   AUTHOR      : Kyriakos Karpouzis                                                          │
│   VERSION     : 1.0                                                                         │
│   DATE        : 02/2020                                                                     │   
└─────────────────────────────────────────────────────────────────────────────────────────────┘ 
#>

Clear-Host

# Clear all errors that have been occured before the execution of the script.
$error.clear()

# This is needed so relative paths can be used
# $directoryPath holds the path the script is located
$invocation = (Get-Variable MyInvocation).Value
$directoryPath = Split-Path $invocation.MyCommand.Path

# Import logger script
. $directoryPath\logger.ps1

# Import appendtocsv script to deal with the missing "-Append" option in Powershell 2.0
# In PowerShell 3.0 and on, there is an -Append option in "Export-Csv" command
# Powershell 2.0 does not have this feature, so a custom script is imported to allow the pass of this option
. $directoryPath\appendtocsv.ps1

Write-Log "Started Application @ $env:computername using PowerShell $($Host.Version.Major)" "INFO"

# Try to get the info and save the report.
# If the script fails, catch the error, log it and stop the execution
try {
    # Load the list.csv
    # This is the file that contains all the computers the script has run in the format of:
    # ID, ComputerName
    $csvFilePath = $directoryPath + '\list.csv'
    $csvFile = Import-Csv $csvFilePath

    # Check if the script has run in this computer. If true, log the action & stop the execution
    if ($computerRecord = $csvFile | Where-Object  {$_.Computer -eq $env:computername}) {
        Write-Log "Record for $env:computername has been found with ID: $($computerRecord.ID). Exiting..." "INFO"
        Exit
    }

    # If list.csv is empty, assume the last ID is 0
    if (!$csvFile) {
        $lastReportId = 0
    }
    else {
        # Else get the ID of the last record in csv
        $lastReportId = [int]($csvFile[-1].ID)
    }
    
    
    # Set the id of this report to be $lastReportId + 1
    $reportId = $lastReportId + 1

    # Function to detect the type of the computer (Desktop or Laptop). This is just a declaration, the function call is occuring below.
    Function Detect-Type {
        $type = 'Desktop'
        #The chassis is the physical container that houses the components of a computer. Check if the machine’s chasis type is 9.Laptop 10.Notebook 14.Sub-Notebook
        if (Get-WmiObject -Class win32_systemenclosure | Where-Object { $_.chassistypes -eq 9 -or $_.chassistypes -eq 10 -or $_.chassistypes -eq 14 })
        { $type = 'Laptop' }
        return $type
    }

    # Get computer type (Custom function declared above)
    $type = Detect-Type

    # Get computer manufacturer & name
    $computerInfo = Get-WmiObject -Class Win32_ComputerSystem | Select-Object Manufacturer, Name

    # Get CPU
    $cpu = (Get-WmiObject -Class Win32_Processor).Name

    # Get Memory Object
    $memory = Get-WmiObject -Class Win32_PhysicalMemory

    # Get all physical drives. Checking for $_.mediatype -eq "Fixed hard disk media" will get ONLY the physical hard disks. Column 'SizeInGb' is added for each object with the calculation of the disk size in gigabytes.
    $drive = Get-WmiObject -Class Win32_DiskDrive | Where-Object { $_.mediatype -eq "Fixed hard disk media" } | Select-Object Model, SerialNumber, InterfaceType, @{n = 'SizeInGb'; e = { [int]($_.Size / 1GB) } }

    # Get Graphics Card. Column 'GpuRam' is added with the calculation of the card memory in gigabytes.
    $gpu = Get-WmiObject -Class Win32_VideoController | Select-Object Name, @{Expression = { $_.AdapterRAM / 1GB }; Label = "GpuRam" }

    # Get Operating System
    $os = (Get-WmiObject -Class Win32_OperatingSystem).Caption

    # Create Word object
    $Word = New-Object -ComObject word.application

    # Hide the word object (windowless mode)
    $Word.Visible = $False

    # Open the template document
    $templateDocPath = $directoryPath + '\templates\inventory_template.docx'
    $Doc = $Word.Documents.Add($templateDocPath)

    # Populate the form fields of the template
    # Forach form field $Control, based on the $Control.Title print the corresponding value
    ForEach ($Control in $Doc.ContentControls) {
        Switch ($Control.Title) {
            # Print report ID
            "id" { $Control.Range.Text = [String]$reportId }
            # Print Computer Type
            "type" { $Control.Range.Text = $type }
            # Print Computer Manufacturer
            "manufacturer" { $Control.Range.Text = $computerInfo.Manufacturer }
            # Print CPU
            "cpu" { $Control.Range.Text = [String]$cpu }
            # Print Memory in a format of: 
            # -Sum of the object sizes in gigabytes
            # -The channels used (e.g. 2-Channel)
            # -Print the minimum speed of memory
            "memory" { $Control.Range.Text = [String](($memory | Measure-Object -Property Capacity -Sum).Sum / 1GB) + "GB " + ($memory | Measure-Object).Count + "-Channel @ " + ($memory | Measure-Object -Property Speed -Minimum).Minimum + " MHz" }
            # Print the hard disks in a format of:
            # -Index Number
            # -Size in gigabytes
            # -Model
            # -Serial Number (trim() method is needed, because sometimes serial numbers will have leading spaces)
            # -Interface Types
            # In case there are more than one drives output them in new line
            "drive" {
                for ($i = 0; $i -le ($drive.length - 1); $i += 1) {
                    if ($i -eq 0) {
                        $Control.Range.Text = [String]($i + 1) + ": " + [String]$drive[$i].SizeInGb + "GB " + $drive[$i].Model + " #" + [String]($drive[$i].SerialNumber).trim() + " " + $drive[$i].InterfaceType
                    }
                    else {
                        $Control.Range.Text += "`n" + [String]($i + 1) + ": " + [String]$drive[$i].SizeInGb + "GB " + $drive[$i].Model + " #" + [String]($drive[$i].SerialNumber).trim() + " " + $drive[$i].InterfaceType 
                    }
                }
            }
            # Print the graphics card with its memory size in gigabytes
            "gpu" { $Control.Range.Text = $gpu.name + " " + [String]$gpu.GpuRam + "GB" }
            # Print the operating system
            "os" { $Control.Range.Text = $os }
            # Print the computer name
            "computer_name" { $Control.Range.Text = $computerInfo.Name }
        }
    }

    # Define the path the report will be saved along with its name. The format of the report is:
    # computerName.docx (e.g. CSS-GEP-01.docx)
    $saveDirAndName = $directoryPath + '\reports\' + $computerInfo.Name + '.docx'
    # Save the report
    # WARNING: *DO NOT* remove the [ref] as it is needed for the report to be saved in PowerShell 2.0
    $Doc.saveAs([ref]$saveDirAndName)
    # Close the document end exit MS Word
    $Doc.Close()
    $Word.Quit()
    
    # Append the record for this report to line.csv
    $lineToAppend = 
    New-Object PSObject -Property @{
        "ID"       = $reportId
        "Computer" = $computerInfo.Name
    }

    $lineToAppend | Export-Csv -Path $csvFilePath -Append -NoTypeInformation

    Write-Log "Document $env:computername.docx Saved with ID: $reportId!" "INFO"
}

# Oops! Error occured. Write it in the log.
catch { 
    Write-Log "$error Quitting..." "ERROR"
}
#
## Windows Printer Fix v4.0.0
## https://github.com/OllieB/WinPrinterFix
#

# Set window title
$host.ui.RawUI.WindowTitle = "WinPrinterFix"

# Start a counter for the 'Ne' number
$NeCounter = 50

# Loop through all PrintUI entries
ForEach($Connection in (Get-ChildItem -Path 'HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections')) {
    
    # Get printer name
    $PrinterName = $Connection.GetValue('Printer')
    
    Write-Host "# Working on $PrinterName (Using: Ne$NeCounter)"

    # Generate the correct naming syntax for the key
    $PrinterNameEscaped = $PrinterName.Replace('\', ',')
    $ServerName = $Connection.GetValue('Server')
    $LocalConnection = $Connection.GetValue('LocalConnection')

    #
    ## Do the connections
    #

    # Put the key path together
    $ConnectionPath = "HKCU:\Printers\Connections\$PrinterNameEscaped"

    # Only create it if it doesn't already exist
    if(-not (Test-Path $ConnectionPath)) {
        
        Write-Warning "Didn't find the connection key"

        # Create the key
        New-Item -Path $ConnectionPath -Force | Out-Null

        # Add the provider value
        New-ItemProperty -Path $ConnectionPath -Name 'Provider' -Value 'win32spl.dll' -Force

        # Add the server value
        New-ItemProperty -Path $ConnectionPath -Name 'Server' -Value $ServerName -Force

        # Add the LocalConnection value
        New-ItemProperty -Path $ConnectionPath -Name 'LocalConnection' -Value $LocalConnection -PropertyType DWORD -Force

        # Add a note
        New-ItemProperty -Path $ConnectionPath -Name 'CreatedByScript' -Value 'Yes' -Force

    } else {
		
        Write-Host "Found the connection key"
		
    }

        
    #
    ## Do the devices
    #

    $DevicesPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Devices"
	
	# Only create it if it doesn't already exist
    if(-not (Test-Path $DevicesPath)) {
        
        # Create the key
        New-Item -Path $DevicesPath -Force | Out-Null

    }
	
	# Only create it if it doesn't already exist
    if(-not (Get-Item -Path $DevicesPath).GetValue($PrinterName)) {
        
        Write-Warning "Didn't find the device property"

        # Add the property
        New-ItemProperty -Path $DevicesPath -Name $PrinterName -Value ('winspool,Ne' + $NeCounter + ':') -Force

    } else {
        
        Write-Host "Found the device property"

    }
    
    
    #
    ## Do the ports
    #

    $PortsPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"
    
	# Only create it if it doesn't already exist
    If(-not (Test-Path $PortsPath)) {
        
        # Create the key
        New-Item -Path $PortsPath -Force | Out-Null

    }
	
	# Only create it if it doesn't already exist
    if(-not (Get-Item -Path $PortsPath).GetValue($PrinterName)) {
        
        Write-Warning "Didn't find the port property"

        # Add the property
        New-ItemProperty -Path $PortsPath -Name $PrinterName -Value ('winspool,Ne' + $NeCounter + ':,15,45') -Force

    } else {
        
        Write-Host "Found the port property"

    }

    # Increment the counter
    $NeCounter++
}

Write-Host `n Microfix Demo `n
Write-Host `n This script was originally designed to reconnect PCs to Microsoft services such as Outlook. `n If the target is a desktop it will re-input the DNS server addresses and restart all network adapters. `n If the target is a laptop, this script will set the connection mode of the active wlan to auto `n and then restart all network adapters. `n
# Checks to see if running as admin. If not, runs script as admin
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
    }

# Automatically sets all hosts to trusted
Set-Item wsman:localhost\client\trustedhosts -Value * -Force

# Main loop for repeating the script
$Repeat = $True
While ($Repeat) {
    Write-Host `n This script reinputs DNS addresses `n and restarts network adapters for desktops and laptops. `n
    Write-Host `n     WARNING: The target PC will lose connection momentarily while network adapters are reset! `n

    # Asks user to input a PC name
    $userinput = Read-Host ('Enter Hostname')

    Write-Host ' '
    Try {
        
        #nslookup FQDN of $userinput for use in remote commands
        $hostname = (nslookup $userinput | Select-String Name | Where-Object LineNumber -eq 4).ToString().Split(' ')[-1]

        #nslookup IP address of $userinput to identify building and device type
        $hostaddress = (nslookup $userinput |Select-String Address | ? LineNumber -eq 5).ToString().Split(' ')[-1]
        
        #Grabs 1st octet of IP to differentiate wired/wireless devices
        $pcmatch = $hostaddress.Substring(0,3).TrimEnd('.')

        #Grabs 3rd octet of IP to determine site location
        $hostmatch = $hostaddress.Substring(8,3).TrimEnd('.')
        
        If ($pcmatch -eq '192') {

            #Matches $hostmatch to IP addresses of sites and sets $case to determine DNS server addresses order
            Switch -Wildcard ($hostmatch)
            {
                '10*' {$site = 'site1'; $case = '1'; Break}
                '11*' {$site = 'site2'; $case = '2'; Break}
                '12*' {$site = 'site3'; $case = '1'; Break}
                '13*' {$site = 'site4'; $case = '1'; Break}
                '14*' {$site = 'site5'; $case = '3'; Break}
                '15*' {$site = 'site6';  $case = '1';Break}
                '16*' {$site = 'site7'; $case = '1'; Break}
                '17*' {$site = 'site8'; $case = '1'; Break}
                '18*' {$site = 'site9'; $case = '1'; Break}
                '19*' {$site = 'site10'; $case = '1'; Break}
                '20*' {$site = 'site11'; $case = '1'; Break}
                '21*' {$site = 'site12'; $case = '1'; Break}
                default { $site = 'default'; Write-Host The IP address does not match any known site with desktop computers. Please try again.; Break }
                }

            #Sets $dnsservers order based on case
            Switch ($case) {
                '1' { $dnsservers = @(
                        '192.168.1.2'
                        '192.168.2.3'
                        '192.168.3.4'
                        '192.168.4.5'
                        ); Break}
                '2' {$dnsservers = @(
                        '192.168.2.3'
                        '192.168.1.2'
                        '192.168.4.5'
                        '192.168.3.4'
                        ); Break}
                '3' {$dnsservers = @(
                        '192.168.1.2'
                        '192.168.2.3'
                        '192.168.3.4'
                        '192.168.4.5'
                        ); Break}
                '4'{$dnsservers = @(
                        '192.168.2.3'
                        '192.168.1.2'
                        '192.168.4.5'
                        '192.168.3.4'
                        ); Break}
                default {Write-Host An error has occurred while matching the case to the DNS servers; Break}
                        }

            #Checks results of matching
            Write-Host Address is: $hostaddress
            Write-Host Match is: $hostmatch
            Write-Host Site is: $site `n

            # Gets service status for Windows Remote Management and starts it
		    Get-Service -Name WinRM -ComputerName $hostname | Start-Service -ErrorAction Stop
                    
            # Actual invoke commands 
            Invoke-Command -ComputerName $hostname -ArgumentList $hostname, $hostaddress, $dnsservers -ScriptBlock { 
                            
                #Passing local variables into remote commands
                param($hostname, $hostaddress, $dnsservers)

                #To confirm variables passed and dns servers are correct
                Write-Host Changing DNS server addresses to: $dnsservers
                Write-Host ' '

                #Sets DNS server addresses on network adapters where the IPV4 address matches a $siteIP address
                Get-NetIPAddress | ? IPAddress -like $hostaddress | Set-DnsClientServerAddress -ServerAddresses $dnsservers
                        
                #Writes the matched IP address out to confirm that the script identified a valid IP on a network adapter, otherwise, gives error message.
                $confirm = Get-NetIPAddress | ? IPAddress -like $hostaddress
                if ($confirm -ne $null) {
                    Write-Host DNS server addresses successfully changed on adapters with IP address: $confirm `n
                    }
                else {
                    Write-Host Error! The DNS server addresses were not changed!
                    }         
                } 

            # Warning message                 
            Write-Host `n Network adapters will be reset. You will lose connection for a moment. `n Please ask the user to restart problematic applications. `n You can close this window or wait for PowerShell to reconnect so you can target another PC.
    
            # Invokes restart of all network adapters on remote PC
            Invoke-Command -ComputerName $hostname -Scriptblock { Get-NetAdapter | Restart-NetAdapter }

            # Confirmation message               
            Write-Host `n Success! `n Connection re-established `n `n 
            }

        ElseIf ($pcmatch -eq '15') {
                        
            #Flushes DNS to prevent expired DNS resolution for laptops
            'ipconfig /flushdns' | cmd
            'exit /silent' | cmd
            cls
            
            #Fixes $hostmatch for shorter IP addresses posessed by wireless devices
            $hostmatch = $hostaddress.Substring(5,3).TrimEnd('.')

            #Matches $hostmatch to IP addresses of sites to impress whoever is running the script
            Switch -Wildcard ($hostmatch)
            {

                '10*' {$site = 'site1'; $case = '1'; Break}
                '11*' {$site = 'site2'; $case = '2'; Break}
                '12*' {$site = 'site3'; $case = '1'; Break}
                '13*' {$site = 'site4'; $case = '1'; Break}
                '14*' {$site = 'site5'; $case = '3'; Break}
                '15*' {$site = 'site6';  $case = '1';Break}
                '16*' {$site = 'site7'; $case = '1'; Break}
                '17*' {$site = 'site8'; $case = '1'; Break}
                '18*' {$site = 'site9'; $case = '1'; Break}
                '19*' {$site = 'site10'; $case = '1'; Break}
                '20*' {$site = 'site11'; $case = '1'; Break}
                '21*' {$site = 'site12'; $case = '1'; Break}
                default { $site = 'error'; Write-Host The IP address does not match any known building with laptop computers. Please try again.; Break }
                }
            
            #Starts WinRM which is required for invoking commands
            Get-Service -Name WinRM -ComputerName $hostname | Start-Service

            #Passing commands to remote target
            Invoke-Command -ComputerName $hostname -ArgumentList $site, $hostaddress -ScriptBlock {
                
                #Passing local variables into remote commands
                param($site, $hostaddress)
                
                #Uses netsh to grab active SSID
                $wlan = (netsh wlan show interfaces | Select-String SSID | ? LineNumber -eq 9).ToString().split(' ')[-1]
                Write-Host `n The PC appears to be connected to a wireless network.`n`n The active wlan is: $wlan `n Last known address is: $hostaddress `n Last known site is: $site 
                
                #Sets active SSID to "Connect automatically"
                netsh wlan set profileparameter name = $wlan connectionmode = auto
                Write-Host $wlan set to auto connection mode. `n `n Network adapters will be reset. `n Please ask the user to restart problematic applications. `n You can close this window or wait for PowerShell to reconnect so you can target another PC.
                
                #Restarts network adapters
                #Get-NetAdapter | Restart-NetAdapter

                Write-Host `n Success! `n Connection re-established `n `n 
                }
            }     
        
        #Catches any unforeseen IP parameters that don't start with 192 or 10
        Else {
            Write-Host The IP address of the device is unrecognized. Please try again.
            }    
        
        # Asks the user if they would like to continue
        $continue = Read-Host ('Would you like to configure another PC? [Y] Yes [N] No (Default is "Y")')
                
        # Any input other than "N" will result in continuing
        if ($continue -ne 'n') {
            $Repeat = $True
            }
                
        # If anser is "N" then exits script
        else {
            Write-Host Exiting Script
            $Repeat = $False
            }
        }

    #Catches errors and restarts script
    catch {
    Write-Host An Error Has Occurred `n
    Write-Host Please Try Again
    $Repeat = $True
    }        
}
Write-Host Microfix Demo                    			    	
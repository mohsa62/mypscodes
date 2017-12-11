Function Test-RegistryValue {
    param(
        [Alias("PSPath")]
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$Path
        ,
        [Parameter(Position = 1, Mandatory = $true)]
        [String]$Name
        ,
        [Switch]$PassThru
        ,
        [Switch]$Set
        ,
        $Value
    ) 

    process {
        $MySubKey = $global:regKey.OpenSubKey($Path,$true)
        if ($MySubKey -ne $null) {
            #$Key = Get-Item -LiteralPath $Path
            if ($MySubKey.GetValue($Name) -ne $null) {
                if ($PassThru) {
                    $MySubKey.GetValue($Name)
                } else {
                    $true
                }
                if ($Set) {
                    $MySubKey.SetValue($Name,$Value)
                }
            } else {
                $false
            }
        } else {
            $false
        }
    }
}

Function Send-Msg {
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$ComputerName
        ,
        [Parameter(Position = 1, Mandatory = $false)]
        [string]$message
    )
    # test remote registry running status
    Write-Host -Object "Checking Remote Registry Service..." -BackgroundColor Blue -ForegroundColor Yellow
    $RemoteRegistryService = Get-Service -Name RemoteRegistry -ComputerName $ComputerName
    if(($RemoteRegistryService).Status -ne "Running")
    {
        Write-Host -Object "not running. starting..." -BackgroundColor Blue -ForegroundColor Yellow
        # starting remote registry service
        $RemoteRegistryService | Start-Service -ErrorAction SilentlyContinue | Out-Null
        if (-not $?)
        {
            Write-Host -Object "Failed to start service." -BackgroundColor Red -ForegroundColor Yellow
            Write-Host -Object "Changing Startup type to automatic..." -BackgroundColor Blue -ForegroundColor Yellow
            # when failed change startup
            $result = (Get-WmiObject -Class win32_service -computername $ComputerName -filter "name='RemoteRegistry'").ChangeStartMode("Automatic")
            Write-Host -Object "Trying to start again..." -BackgroundColor Blue -ForegroundColor Yellow 
            $RemoteRegistryService | Start-Service -ErrorAction SilentlyContinue | Out-Null
        }
        $RemoteRegistryService = Get-Service -Name RemoteRegistry -ComputerName $ComputerName
        if(($RemoteRegistryService).Status -eq "Running")
        {
            Write-Host -Object "success" -BackgroundColor Green -ForegroundColor Blue
        }
        else
        {
            Write-Host -Object "Permanatly Failed" -BackgroundColor Red -ForegroundColor Yellow
        }
    }
    else
    {
        Write-Host -Object "is running" -BackgroundColor Green -ForegroundColor Blue
    }
    Write-Host -Object "Creating Powershell Remote Session..."
    try 
    {
        $result = Invoke-Command -ComputerName $ComputerName { 1 } -ErrorAction Stop
        #New-PSSession -ComputerName $ComputerName -Name session -ErrorAction SilentlyContinue -ErrorVariable SessionError
    } 
    catch 
    {
        if ($SessionError.FullyQualifiedErrorId -like "*WinRMOperationTimeout*") 
        {
            Write-Host -Object "WinRM is not Enabled or Firewall is Active. Checking Firewall Status..." -BackgroundColor Blue -ForegroundColor Yellow
            $keystd = "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile"
            $type = [Microsoft.Win32.RegistryHive]::LocalMachine
            $global:regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $ComputerName)
            $thekeystd = $global:regKey.OpenSubKey($keystd,$true)
            if ($thekeystd.GetValue("EnableFirewall") -eq 0)
            {
                Write-Host -Object "Firewall in Standard Profile is OFF" -BackgroundColor Green -ForegroundColor blue
            }
            else
            {
                Write-Host -Object "Firewall in Standard Profile is ON" -BackgroundColor Green -ForegroundColor blue
            }
            $keydom = "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile"
            $keypub = "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\PublicProfile"
            $thekeypub = $global:regKey.OpenSubKey($keypub,$true)
            if ($thekeypub.GetValue("EnableFirewall") -eq 0)
            {
                Write-Host -Object "Firewall in Public Profile is OFF" -BackgroundColor Green -ForegroundColor blue
            }
            else
            {
                Write-Host -Object "Firewall in Public Profile is ON" -BackgroundColor Green -ForegroundColor blue
            }
            $thekeydom = $global:regKey.OpenSubKey($keypub,$true)
            if ($thekeydom.GetValue("EnableFirewall") -eq 0)
            {
                Write-Host -Object "Firewall in Domain Profile is OFF" -BackgroundColor Green -ForegroundColor blue
            }
            else
            {
                Write-Host -Object "Firewall in Domain Profile is ON" -BackgroundColor Green -ForegroundColor blue
            }
        }
        else
        {
            Write-Host -Object "Creating Powershell Session Failed" -BackgroundColor Red -ForegroundColor Yellow
        }
    }
    $type = [Microsoft.Win32.RegistryHive]::LocalMachine
    $global:regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $ComputerName)
    Write-Host -Object "Checking registry..." -BackgroundColor Blue -ForegroundColor Yellow
    if (Test-RegistryValue -Path "SYSTEM\CurrentControlSet\Control\Terminal Server" -Name AllowRemoteRPC) {
    Write-Host -Object "The Key & Property Exists." -BackgroundColor Green -ForegroundColor blue
    if(Test-RegistryValue -Path "SYSTEM\CurrentControlSet\Control\Terminal Server" -Name AllowRemoteRPC -PassThru) {
        Write-Host -Object "Msg is Enabled." -BackgroundColor Green -ForegroundColor blue
    }
    else
    {
        Write-Host -Object "Setting Registery Property to Enable Msg..." -BackgroundColor Blue -ForegroundColor Yellow
        #New-ItemProperty -Path "SYSTEM\CurrentControlSet\Control\Terminal Server" -Name AllowRemoteRPC -Value 1 -ErrorAction SilentlyContinue -Force | Out-Null
        Test-RegistryValue -Path "SYSTEM\CurrentControlSet\Control\Terminal Server" -Name AllowRemoteRPC -set -Value 1
        if ($?){
            Write-Host -Object "Success Setting the Registry Value." -BackgroundColor Green -ForegroundColor blue
        }
        else
        {
            Write-Host -Object "Failed to Set the Registry Property" -BackgroundColor Red -ForegroundColor Yellow
        }
    }
    }
    else
                                            {
    Write-Host -Object "the AllowRemoteRPC Property not found." -BackgroundColor Magenta -ForegroundColor Yellow
    Write-Host -Object "Creating Registery Prperty to Enable Msg..." -BackgroundColor Blue -ForegroundColor Yellow
    New-ItemProperty -Path "SYSTEM\CurrentControlSet\Control\Terminal Server" -Name AllowRemoteRPC -Value 1 -ErrorAction SilentlyContinue -Force | Out-Null
    if ($?){
        Write-Host -Object "Success Setting the Registry Value." -BackgroundColor Green -ForegroundColor blue
    }
    else
    {
        Write-Host -Object "Failed to Set the Registry Property" -BackgroundColor Red -ForegroundColor Yellow
    }
    }
    
    Write-Host -Object "Sending Message..." -BackgroundColor Blue -ForegroundColor Yellow
    get-loggedonuser -computername $ComputerName | Format-Table -AutoSize
    $recipient = Read-Host -Prompt "Enter UserName or SessionID"
    msg $recipient /server:$ComputerName /V /W $message
}
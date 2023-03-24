#########################################################
#														#
#	 Get-ClientIPAddresses by Jean Louw					#
#	 Blog http://powershellneedfulthings.blogspot.com/	#
#														#
#########################################################

function Get-ComputerSite ($ip){
Write-Host "Current IP:" $ip
$site = $null
$computer = [System.Net.Dns]::gethostentry($ip) 
$site = nltest /server:$($computer.hostname) /dsgetsite
Return $site[0]
}

$ADSiteWMI = @{Name="ADSite";expression={Get-ComputerSite $($_.ClientIP)}}
$ADSite = @{Name="ADSite";expression={Get-ComputerSite $($_.ClientIPAddress)}}

foreach ($server in get-mailboxserver){
write-host "Current server: " $server
$filename = ".\" + $server + ".csv"
$LogonStats = Get-LogonStatistics -server $server | sort UserName -Unique 
$LogonStats | select UserName, ClientIPAddress, $ADSite | Export-Csv $filename 
}

foreach ($server in (Get-ExchangeServer | Where {$_.IsExchange2007OrLater -eq $false})){
write-host "Current server: " $server
$filename = ".\" + $server + ".csv"
$LogonStats = Get-Wmiobject -namespace root\MicrosoftExchangeV2 -class Exchange_Logon -Computer $server | sort MailboxDisplayName -Unique
$LogonStats | select MailboxDisplayName, ClientIP, $ADSiteWMI | Export-Csv $filename
}

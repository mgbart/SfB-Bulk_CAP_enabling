<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.178
	 Created on:   	8/25/2020 1:38 PM
	 Created by:   	Matías Bernhardt
	 Organization: 	Matías Bernhardt
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>



###############################################################################################################################################################################################
#                                                                                                                                                                                             #
# Prepare CSV File as following example:                                                                                                                                                      #
#                                                                                                                                                                                             #                       
# DisplayNumber,E164,Type,Account,DisplayName,VoicePolicy,DialPlan,Pool,ClientPolicy,PIN                                                                                                      #
# + 1 222 333 4444, +12223334444, CAP, PHONE-AMER-US-NYC-12223334444, USA New York Conference Room 1, AMER-US-NYC-International, AMER-US-NYC, pool01.contoso.com, AMER-US-STD-CAP, 111111     #
#                                                                                                                                                                                             #
#                                                                                                                                                                                             #
###############################################################################################################################################################################################

$Date = Get-Date -Format yyyy-MM-dd___hh-mm

# General Config

$xml = [xml](Get-Content $PSScriptRoot\config.xml)

$SipDomain = $xml.Config.SipDomain
$OU = $xml.Config.OU


$CSVFile 	= "$PSScriptRoot\import\" + $xml.Config.ImportFileName
$NBSuccess 	= "$PSScriptRoot\logs\" + $Date + "____Success.csv"
$NBFailed 	= "$PSScriptRoot\logs\" + $Date + "____Failed.csv"


$Pool = $null
$ClientPolicy = $null
$PIN = $null

$CSV = import-csv $CSVFile

"Account,DisplayName,E164,Pool" | Out-File $NBSuccess
"Account,DisplayName,E164,Pool" | Out-File $NBFailed


Foreach ($Phone in $CSV) {
	$Account = $Phone.Account.Trim()
	$DisplayName = $Phone.DisplayName.Trim()
	$DisplayNumber = $Phone.DisplayNumber.Trim()
	$E164 = $Phone.E164.Trim()
	$Pool = $Phone.Pool.Trim()
	$SIP = "sip:" + $Account + "@" + $SipDomain
	
	If ($E164 -like "+*") {
		$LineURI = "tel:" + $E164 + ";ext=" + $E164.Substring($E164.length - 4, 4)
		$LineURI = $LineURI.Trim()
		$DN = "cn=" + $Account + "," + $OU
		$DN = $DN.Trim()
		
		# Create the AD Object
		New-ADObject -Type contact -Name $Account -DisplayName $DisplayName -Path $OU
		Start-Sleep -Seconds 20
		Get-ADObject -Filter 'name -eq $Account' -SearchBase $OU | Set-ADObject -add @{ telephoneNumber = $DisplayNumber }
		
		# Create the SfB Object	
		New-CsCommonAreaPhone -LineUri $LineURI -SipAddress $SIP -RegistrarPool $Pool -DN $DN -DisplayNumber $DisplayNumber
		
		Write-host  $E164 "   " $DisplayName "   ", $LineURI, "   ", $SIP, "     ", $pool
		
		$Account + "," + $DisplayName + "," + $E164 + "," + $Pool | Out-File $NBSuccess -Append
	}
	Else {
		$Account + "," + $DisplayName + "," + $E164 + "," + $Pool | Out-File $NBFailed -Append
	}
	
}

Start-Sleep -Seconds 180

foreach ($Phone in $CSV) {
	
	$DisplayName = $Phone.DisplayName.Trim()
	$PIN = $Phone.PIN.Trim()
	$VoicePolicy = $Phone.VoicePolicy.Trim()
	$ClientPolicy = $Phone.ClientPolicy.Trim()
	$DialPlan = $Phone.DialPlan.Trim()
	
	Get-CsCommonAreaPhone -Identity $DisplayName | Set-CsClientPin -Pin $PIN 
	Get-CsCommonAreaPhone -Identity $DisplayName | Grant-CsClientPolicy -PolicyName $ClientPolicy
	Get-CsCommonAreaPhone -Identity $DisplayName | Grant-CsDialPlan -PolicyName $DialPlan
	Get-CsCommonAreaPhone -Identity $DisplayName | Grant-CsVoicePolicy -PolicyName $VoicePolicy 
}

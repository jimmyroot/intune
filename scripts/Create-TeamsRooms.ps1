# A quick and dirty script to create Teams Rooms, configure some preferences, and license them
#
# Create a class for a simple Teams Room Object, we can use this to create an 
# array of "rooms" to iterate over
#
# Before running this script you'll need to import/install and authenticate to the following
# modules:
#
# MicrosoftTeams
# ExchangeOnline
# MSOnline
# Microsoft.Graph

class TeamsRoom {
	[string]$Identifier
	[string]$DisplayName
	
	TeamsRoom( [string]$Identifier, [string]$DisplayName) {
		$this.Identifier = $Identifier
		$this.DisplayName = $DisplayName
	}
}

# Create an array list
$roomsArray = [System.Collections.ArrayList]::new()

# Populate it with our rooms
$roomsArray.Add([TeamsRoom]::new('lux.andy', 'Luxembourg Andy')) | Out-Null
# $roomsArray.Add([TeamsRoom]::new('lux.cafe', 'Luxembourg Cafe')) | Out-Null
# $roomsArray.Add([TeamsRoom]::new('lux.reuter', 'Luxembourg Reuter')) | Out-Null
# $roomsArray.Add([TeamsRoom]::new('lux.hamilius', 'Luxembourg Hamilius')) | Out-Null
# $roomsArray.Add([TeamsRoom]::new('lux.monterey', 'Luxembourg Monterey')) | Out-Null

# Loop through the array list and do some magic
$roomsArray | ForEach-Object {

    $Identifier = $_.Identifier
    $DisplayName = $_.DisplayName

    # Get the unique ObjectId for the MgGraph commands
    $uid = (Get-MSOLUser -UserPrincipalName "$Identifier@ikpartners.com" | Select ObjectId).ObjectId
	
    # Create room mailbox
    New-Mailbox -Name "$DisplayName" -Alias "$Identifier" -MicrosoftOnlineServicesID "$Identifier@ikpartners.com" -Room -EnableRoomMailboxAccount $true -RoomMailboxPassword (ConvertTo-SecureString -String "IKinvest2014!" -AsPlainText -Force)

    # Allow some time to pass before we try to access the new room
    Start-Sleep -Seconds 15

    # Configure calendar processing preferences
    Set-CalendarProcessing -Identity "$Identifier" -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -BookingWindowInDays 366 -ProcessExternalMeetingMessages $true -AddAdditionalResponse $true -AdditionalResponse "Your meeting has been scheduled. You can connect to the meeting by tapping the Join button on the touch panel in the room."

    # Set usage location (required for license)
    Set-MsolUser -UserPrincipalName "$Identifier@ikpartners.com" -PasswordNeverExpires $true -UsageLocation "GB"

    # Add license using new MgGraph cmdlets
    $TeamsProSKU = Get-MgSubscribedSku -All | where SkuPartNumber -eq "Microsoft_Teams_Rooms_Pro"
    Set-MgUserLicense -UserId $uid -AddLicenses @{SkuId = $TeamsProSKU.SkuId} -RemoveLicenses @()
}


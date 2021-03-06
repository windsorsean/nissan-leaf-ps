﻿<#
    Original code by Sean Hart (@seanhart) - 2018-06-14
    [2021-04-15] Updates:
        - Fix door unlock endpoint.
    [2018-10-30] Updates:
        - Added ability to request door lock/unlock.
        - Added AccountID to properties (needed for door lock/unlock).
        - Some code tidying (I'm still learning how functions work in PowerShell).
    [2018-09-19] Updates:
        - Added param to set country (defaults to 'CA' for Canada, 'US' can also be used.)
        - Changed language field from "en_CA" to "en_US" for no particular reason.
    [2018-08-27] Updates:
        - Pass car object to functions instead of global variable.
        - Bug with -update resolved.
    [2018-07-17] Updates:
        - Change code to use functions
        - Location; object returned has a Location property with an object instead of lat/long properties.

    Thanks to Ben Woodford (https://gist.github.com/BenWoodford) for documenting the Nissan Connect API
#>

param (
    [string]$username = $( Read-Host "Input username"),
    [string]$password = $( Read-Host "Input password" ),
    [string][ValidateSet("CA","US")]$country = "CA",  #Other country codes may work here?
    [switch]$update = $false,
    [switch]$climate_on = $false,
    [int]$set_temp = 0,
    [switch]$climate_off = $false,
    [switch]$charge_on = $false,
    [switch]$locate = $false,
    [switch]$last_location = $false,
    [switch]$no_map = $false,
    [switch]$door_lock = $false,
    [switch]$door_unlock = $false,
    [string]$pin_code = ""
)

# --- FUNCTION SECTION ---

# Show script usage info
function Show-Usage {
    Write-Host "`nUsage:"
    Write-Host "  -username <NissanConnect username>" -nonewline
	    Write-Host " [required]" -ForegroundColor darkgray
    Write-Host "  -password <password>" -nonewline
	    Write-Host " [required]" -ForegroundColor darkgray
    Write-Host "  -country <CA/US>" -nonewline
	    Write-Host " ['CA' for Canada (default), 'US' for United States]" -ForegroundColor darkgray
    Write-Host "  -update" -nonewline
	    Write-Host " [use to refresh data]" -ForegroundColor darkgray
    Write-Host "  -climate_on" -nonewline
	    Write-Host " [use to turn on climate control, can't be used with climate_off]" -ForegroundColor darkgray
    Write-Host "  -set_temp <integer in C>" -nonewline
	    Write-Host " [optional, but requires climate_on]" -ForegroundColor darkgray
    Write-Host "  -climate_off" -nonewline
	    Write-Host " [use to turn off climate control, can't be used with climate_on]" -ForegroundColor darkgray
    Write-Host "  -charge_on" -nonewline
	    Write-Host " [use to start charging]" -ForegroundColor darkgray
    Write-Host "  -locate" -nonewline
	    Write-Host " [use to initiate a location request, will open Google Maps in default browser]" -ForegroundColor darkgray
    Write-Host "  -last_location" -nonewline
	    Write-Host " [use to get the last known location, will open Google Maps in default browser]" -ForegroundColor darkgray
    Write-Host "  -no_map" -nonewline
	    Write-Host " [use to stop Google Maps from opening on location request]" -ForegroundColor darkgray
    Write-Host "  -door_lock" -nonewline
	    Write-Host " [use to send a remote door lock, requires -pin_code]" -ForegroundColor darkgray
    Write-Host "  -door_unlock" -nonewline
	    Write-Host " [use to send a remote door unlock, requires -pin_code]" -ForegroundColor darkgray
    Write-Host "  -pin_code <4 digit number>" -nonewline
	    Write-Host " [required for door lock/unlock]`n" -ForegroundColor darkgray
}


# Initiate connection to Nissan Connect and return current car information
function Connect-Nissan {
    $payload = @{
        'authenticate' = @{
          'brand-s'= 'N';
          'country'= $country;
          'language-s'= 'en_US';
          'userid' = $username;
          'password' = $password
        }
      }

    $data = Invoke-RestMethod -Uri ($baseUrl + 'auth/authenticationForAAS') -Method Post -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'

    if ($data.authToken -eq $null) {
        Write-Host 'Error:' -ForegroundColor red
        Write-Output $data
        Exit
    }

    #$data | ConvertTo-Json | Write-Output

    $header['Authorization'] = $data.authToken

    $car = [PSCustomObject]@{
        'AccountID' = $data.accountID;
        'VIN' = $data.vehicles[0].uvi;
        'Year' = $data.vehicles[0].modelyear;
        'Nickname' = $data.vehicles[0].nickname;
        'BatteryCharge' = $data.vehicles[0].batteryRecords.batteryStatus.soc.value;
        'RangeClimateOn' = ($data.vehicles[0].batteryRecords.cruisingRangeAcOn / 1000);
        'RangeClimateOff' = ($data.vehicles[0].batteryRecords.cruisingRangeAcOff / 1000);
        'PlugState' = $data.vehicles[0].batteryRecords.pluginState;
        'Charging' = $data.vehicles[0].batteryRecords.batteryStatus.batteryChargingStatus;
        'InsideTemp' = $data.vehicles[0].interiorTempRecords.inc_temp;
        'LastUpdate' = Get-Date $data.vehicles[0].batteryRecords.lastUpdatedDateAndTime;
    }

    return $car
}


# Request a refresh from the car
function Update-Leaf($car) {
    $data = Invoke-RestMethod -Uri ($baseUrl + 'battery/vehicles/' + $car.VIN + '/getChargingStatusRequest') -Method Get -Headers $header -WebSession $mysession -ContentType 'application/json'
    if ($data.batteryRecords -ne $null) {
        $car.BatteryCharge = $data.batteryRecords.batteryStatus.soc.value
        $car.RangeClimateOn = ($data.batteryRecords.cruisingRangeAcOn / 1000)
        $car.RangeClimateOff = ($data.batteryRecords.cruisingRangeAcOff / 1000)
        $car.PlugState = $data.batteryRecords.pluginState
        $car.Charging = $data.batteryRecords.batteryStatus.batteryChargingStatus
        $car.InsideTemp = $data.temperatureRecords.inc_temp
        $car.LastUpdate = Get-Date $data.batteryRecords.lastUpdatedDateAndTime
        return $car
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
        return $data
    }
}


# Turn on climate control
function Send-LeafClimateOn($car) {
    $payload = @{
        'executionTime' = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
      }

    if ($set_temp -gt 0) {
        $payload['preACtemp'] = $set_temp
        $payload['preACunit'] = 'C'
    }

    $data = Invoke-RestMethod -Uri ($baseUrl + 'hvac/vehicles/' + $car.VIN + '/activateHVAC') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.messageDeliveryStatus -eq 'Success') {
        return $true
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
        return $false
    }

}


# Turn off climate control
function Send-LeafClimateOff($car) {
    $payload = @{
        'executionTime' = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
      }

    $data = Invoke-RestMethod -Uri ($baseUrl + 'hvac/vehicles/' + $car.VIN + '/deactivateHVAC') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.messageDeliveryStatus -eq 'Success') {
        return $true
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
        return $false
    }
}


# Start charge
function Send-LeafChargeStart($car) {
    $payload = @{
        'executionTime' = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
      }

    $data = Invoke-RestMethod -Uri ($baseUrl + 'battery/vehicles/' + $car.VIN + '/remoteChargingRequest') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.messageDeliveryStatus -eq 'Success') {
        return $true
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
        return $false
    }
}


# Location refresh
function Update-LeafLocation($car) {
    $payload = @{
        'searchPeriod' = (Get-Date -Month ((Get-Date).Month - 1)).ToString("yyyyMMdd") + "," + (Get-Date).ToString("yyyyMMdd");
        'acquiredDataUpperLimit' = '1';
        'serviceName' = 'MyCarFinderResult'
      }

    $data = Invoke-RestMethod -Uri ($baseUrl + 'vehicleLocator/vehicles/' + $car.VIN + '/refreshVehicleLocator') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.sandsNotificationEvent -ne $null) {
        $location = [PSCustomObject]@{
            'Updated' = Get-Date $data.sandsNotificationEvent.sandsNotificationEvent.head.receivedDate;
            'Latitude' = $data.sandsNotificationEvent.sandsNotificationEvent.body.location.latitudeDMS;
            'Longitude' = $data.sandsNotificationEvent.sandsNotificationEvent.body.location.longitudeDMS
        }
        $location | Add-Member -MemberType NoteProperty -Name Link -Value ('https://www.google.com/maps?q=' + $location.Latitude + ',' + $location.Longitude)

        return $location
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
        return $data
    }

}


# Get last location
function Get-LeafLocation($car) {
    $payload = @{
        'searchPeriod' = (Get-Date -Month ((Get-Date).Month - 1)).ToString("yyyyMMdd") + "," + (Get-Date).ToString("yyyyMMdd");
        'acquiredDataUpperLimit' = '1';
        'serviceName' = 'MyCarFinderResult'
      }

    $data = Invoke-RestMethod -Uri ($baseUrl + 'vehicleLocator/vehicles/' + $car.VIN + '/getNotificationHistory') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.sandsNotificationEvent -ne $null) {
        $location = [PSCustomObject]@{
            'Updated' = Get-Date $data.sandsNotificationEvent.sandsNotificationEvent.head.receivedDate;
            'Latitude' = $data.sandsNotificationEvent.sandsNotificationEvent.body.location.latitudeDMS;
            'Longitude' = $data.sandsNotificationEvent.sandsNotificationEvent.body.location.longitudeDMS
        }
        $location | Add-Member -MemberType NoteProperty -Name Link -Value ('https://www.google.com/maps?q=' + $location.Latitude + ',' + $location.Longitude)

        return $location
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data

        return $data
    }
}


# Send door lock command
function Send-DoorLock($car, $pinCode) {
    $payload = @{
        'remoteRequest' = @{
            'authorizationKey' = $pinCode
        }
      }

    try {
        $data = Invoke-RestMethod -Uri ($baseUrl + 'remote/vehicles/' + $car.VIN + '/accounts/' + $car.AccountID + '/rdl/createRDL') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
        if ($data.messageDeliveryStatus -eq 'Success') {
            return $true
        } else {
            Write-Host 'Failed.' -ForegroundColor red
            Write-Output $data
            return $false
        }
    } catch {
        Write-Host 'Failed, check pin code.' -ForegroundColor red
        Write-Output $_
        return $false
    }
}


# Send door unlock command
function Send-DoorUnLock($car, $pinCode) {
    $payload = @{
        'remoteRequest' = @{
            'authorizationKey' = $pinCode
        }
      }

    try {
        $data = Invoke-RestMethod -Uri ($baseUrl + 'remote/vehicles/' + $car.VIN + '/accounts/' + $car.AccountID + '/rdul/createRDUL') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
        if ($data.messageDeliveryStatus -eq 'Success') {
            return $true
        } else {
            Write-Host 'Failed.' -ForegroundColor red
            Write-Output $data
            return $false
        }
    } catch {
        Write-Host 'Failed, check pin code.' -ForegroundColor red
        Write-Output $_
        return $false
    }
}


# --- CODE SECTION ---

# Global variables
$header = @{'API-Key' = 'f950a00e-73a5-11e7-8cf7-a6006ad3dba0'}
$baseUrl = 'https://icm.infinitiusa.com/NissanLeafProd/rest/'
$mysession = New-Object Microsoft.PowerShell.Commands.WebRequestSession

Show-Usage

if (($username -eq '') -or ($password -eq '')) { 
    Write-Host "Username and password are required."
    Exit 
}

$cardata = Connect-Nissan
Write-Host 'Current Leaf Data:' -NoNewline
Write-Host ($cardata | Format-List | Out-String)

# Process update request
if ($update) {
    Write-Host "`nRequesting data refresh..."

    # If we don't get updated data in the first request, wait 5 seconds and re-connect
    $LastUpdate = $cardata.LastUpdate
    $cardata = Update-Leaf $cardata
    if ($LastUpdate -eq $cardata.LastUpdate) {
        Start-Sleep -Seconds 5
        $cardata = $null
        $cardata = Connect-Nissan
    }
    Write-Host "`nUpdated Leaf Data:" -NoNewline
    Write-Host ($cardata | Format-List | Out-String)
}

# Process climate on
if ($climate_on -and (-not $climate_off)) {
    Write-Host "`nSending climate on request..."
    if ((Send-LeafClimateOn $cardata) -eq $true) { Write-Host "`nClimate on request sent successfully." }
}

# Process climate off
if ($climate_off -and (-not $climate_on)) {
    Write-Host "`nSending climate off request..."
    if ((Send-LeafClimateOff $cardata) -eq $true) { Write-Host "`nClimate off request sent successfully." }
}

# Process charge on
if ($charge_on) {
    Write-Host "`nSending charge start request..."
    if ((Send-LeafChargeStart $cardata) -eq $true) { Write-Host "`nStart charge request sent successfully." }
}

# Process last location
if ($last_location) {
    Write-Host "`nRequesting last location..."
    $loc = Get-LeafLocation $cardata
    if ($loc.errorCode -eq $null) {
        $cardata | Add-Member -MemberType NoteProperty -Name Location -Value $loc
        Write-Host 'Location data:' -NoNewline
        Write-Host ($loc | Format-List | Out-String)
        if (-not $no_map) { Start-Process $loc.Link }
    }
}

# Process location refresh
if ($locate) {
    Write-Host "`nRequesting location update..."
    $loc = Update-LeafLocation $cardata
    if ($loc.errorCode -eq $null) {
        $cardata | Add-Member -MemberType NoteProperty -Name Location -Value $loc
        Write-Host 'Location data:' -NoNewline
        Write-Host ($loc | Format-List | Out-String)
        if (-not $no_map) { Start-Process $loc.Link }
    } else {
        Write-Host "`nUnable to refresh location, getting last known location..."
        $loc = Get-LeafLocation($cardata)
        if ($loc.errorCode -eq $null) {
            $cardata | Add-Member -MemberType NoteProperty -Name Location -Value $loc
            Write-Host 'Location data:' -NoNewline
            Write-Host ($loc | Format-List | Out-String)
            if (-not $no_map) { Start-Process $loc.Link }
        }
    }
}

# Process door lock
if ($door_lock) {
    if ($pin_code -eq "") {
        Write-Host "`npin_code required for door lock."
    } else {
        Write-Host "`nSending door lock request..."
        if ((Send-DoorLock $cardata $pin_code) -eq $true) { Write-Host "`nDoor lock request sent successfully." }
    }
}

# Process door unlock
if ($door_unlock) {
    if ($pin_code -eq "") {
        Write-Host "`npin_code required for door unlock."
    } else {
        Write-Host "`nSending door unlock request..."
        if ((Send-DoorUnLock $cardata $pin_code) -eq $true) { Write-Host "`nDoor unlock request sent successfully." }
    }
}

# output object
Write-Output $cardata

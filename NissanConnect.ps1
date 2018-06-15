# Original code by Sean Hart (@seanhart) - 2018-06-14
# Thanks to Ben Woodford (https://gist.github.com/BenWoodford) for documenting the Nissan Connect API

param (
    [string]$username = $( Read-Host "Input username"),
    [string]$password = $( Read-Host "Input password" ),
    [switch]$update = $false,
    [switch]$climate_on = $false,
    [int]$set_temp = 0,
    [switch]$climate_off = $false,
    [switch]$charge_on = $false,
    [switch]$locate = $false,
    [switch]$last_location = $false,
    [switch]$no_map = $false
)

Write-Host "`nUsage:"
Write-Host "  -username <NissanConnect username>" -nonewline
	Write-Host " [required]" -ForegroundColor darkgray
Write-Host "  -password <password>" -nonewline
	Write-Host " [required]" -ForegroundColor darkgray
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
	Write-Host " [use to stop Google Maps from opening on location request]`n" -ForegroundColor darkgray

if (($username -eq '') -or ($password -eq '')) { Exit }

$payload = @{
    'authenticate' = @{
      'brand-s'= 'N';
      'country'= 'CA';
      'language-s'= 'en_CA';
      'userid' = $username;
      'password' = $password
    }
  }

$header = @{'API-Key' = 'f950a00e-73a5-11e7-8cf7-a6006ad3dba0'}

$baseUrl = 'https://icm.infinitiusa.com/NissanLeafProd/rest/'

$data = Invoke-RestMethod -Uri ($baseUrl + 'auth/authenticationForAAS') -Method Post -Headers $header -Body ($payload | ConvertTo-Json) -SessionVariable mysession -ContentType 'application/json'

if ($data.authToken -eq $null) {
    Write-Host 'Error:' -ForegroundColor red
    Write-Output $data
    Exit
}

$header['Authorization'] = $data.authToken

$cardata = [PSCustomObject]@{
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

Write-Host 'Current Leaf Data:' -NoNewline
Write-Host ($cardata | Format-List | Out-String)


# Refresh data
if ($update) {
    Write-Host "`nRequesting data refresh..."
    $data = Invoke-RestMethod -Uri ($baseUrl + 'battery/vehicles/' + $cardata.VIN + '/getChargingStatusRequest') -Method Get -Headers $header -WebSession $mysession -ContentType 'application/json'
    if ($data.batteryRecords -ne $null) {
        $cardata.BatteryCharge = $data.batteryRecords.batteryStatus.soc.value
        $cardata.RangeClimateOn = ($data.batteryRecords.cruisingRangeAcOn / 1000)
        $cardata.RangeClimateOff = ($data.batteryRecords.cruisingRangeAcOff / 1000)
        $cardata.PlugState = $data.batteryRecords.pluginState
        $cardata.Charging = $data.batteryRecords.batteryStatus.batteryChargingStatus
        $cardata.InsideTemp = $data.temperatureRecords.inc_temp
        $cardata.LastUpdate = Get-Date $data.batteryRecords.lastUpdatedDateAndTime
        Write-Host "`nUpdated Leaf Data:" -NoNewline
        Write-Host ($cardata | Format-List | Out-String)
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
    }
}


# Turn on climate control
if ($climate_on -and (-not $climate_off)) {
    $payload = @{
        'executionTime' = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
      }

    if ($set_temp -gt 0) {
        $payload['preACtemp'] = $set_temp
        $payload['preACunit'] = 'C'
    }

    $data = Invoke-RestMethod -Uri ($baseUrl + 'hvac/vehicles/' + $cardata.VIN + '/activateHVAC') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.messageDeliveryStatus -eq 'Success') {
        Write-Host "`nClimate on request sent successfully."
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
    }

}


# Turn off climate control
if ($climate_off -and (-not $climate_on)) {
    $payload = @{
        'executionTime' = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
      }

    $data = Invoke-RestMethod -Uri ($baseUrl + 'hvac/vehicles/' + $cardata.VIN + '/deactivateHVAC') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.messageDeliveryStatus -eq 'Success') {
        Write-Host "`nClimate off request sent successfully."
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
    }

}


# Start charge
if ($charge_on) {
    $payload = @{
        'executionTime' = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
      }

    $data = Invoke-RestMethod -Uri ($baseUrl + 'battery/vehicles/' + $cardata.VIN + '/remoteChargingRequest') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.messageDeliveryStatus -eq 'Success') {
        Write-Host "`nStart charge request sent successfully."
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
    }

}


# Location refresh
if ($locate) {
    $payload = @{
        'searchPeriod' = (Get-Date -Month ((Get-Date).Month - 1)).ToString("yyyyMMdd") + "," + (Get-Date).ToString("yyyyMMdd");
        'acquiredDataUpperLimit' = '1';
        'serviceName' = 'MyCarFinderResult'
      }

    Write-Host "`nRequesting location update..."

    $data = Invoke-RestMethod -Uri ($baseUrl + 'vehicleLocator/vehicles/' + $cardata.VIN + '/refreshVehicleLocator') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.sandsNotificationEvent -ne $null) {
        $location = [PSCustomObject]@{
            'Updated' = Get-Date $data.sandsNotificationEvent.sandsNotificationEvent.head.receivedDate;
            'Latitude' = $data.sandsNotificationEvent.sandsNotificationEvent.body.location.latitudeDMS;
            'Longitude' = $data.sandsNotificationEvent.sandsNotificationEvent.body.location.longitudeDMS
        }
        $location | Add-Member -MemberType NoteProperty -Name Link -Value ('https://www.google.com/maps?q=' + $location.Latitude + ',' + $location.Longitude)
        $cardata | Add-Member -MemberType NoteProperty -Name Latitude -Value $location.Latitude
        $cardata | Add-Member -MemberType NoteProperty -Name Longitude -Value $location.Longitude

        Write-Host 'Location data:' -NoNewline
        Write-Host ($location | Format-List | Out-String)
        if (-not $no_map) { Start-Process $location.Link }
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
    }

}


# Get last location
if ($last_location) {
    $payload = @{
        'searchPeriod' = (Get-Date -Month ((Get-Date).Month - 1)).ToString("yyyyMMdd") + "," + (Get-Date).ToString("yyyyMMdd");
        'acquiredDataUpperLimit' = '1';
        'serviceName' = 'MyCarFinderResult'
      }

    Write-Host "`nRequesting last location..."

    $data = Invoke-RestMethod -Uri ($baseUrl + 'vehicleLocator/vehicles/' + $cardata.VIN + '/getNotificationHistory') -Method POST -Headers $header -Body ($payload | ConvertTo-Json) -WebSession $mysession -ContentType 'application/json'
    if ($data.sandsNotificationEvent -ne $null) {
        $location = [PSCustomObject]@{
            'Updated' = Get-Date $data.sandsNotificationEvent.sandsNotificationEvent.head.receivedDate;
            'Latitude' = $data.sandsNotificationEvent.sandsNotificationEvent.body.location.latitudeDMS;
            'Longitude' = $data.sandsNotificationEvent.sandsNotificationEvent.body.location.longitudeDMS
        }
        $location | Add-Member -MemberType NoteProperty -Name Link -Value ('https://www.google.com/maps?q=' + $location.Latitude + ',' + $location.Longitude)
        $cardata | Add-Member -MemberType NoteProperty -Name Latitude -Value $location.Latitude
        $cardata | Add-Member -MemberType NoteProperty -Name Longitude -Value $location.Longitude

        Write-Host 'Location data:' -NoNewline
        Write-Host ($location | Format-List | Out-String)
        if (-not $no_map) { Start-Process $location.Link }
    } else {
        Write-Host 'Failed.' -ForegroundColor red
        Write-Output $data
    }
}

# output object
Write-Output $cardata
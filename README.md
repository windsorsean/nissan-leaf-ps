# NissanConnect.ps1
PowerShell cmdlet for the Nissan Leaf using NissanConnect EV APIs.

## Usage
```
.\NissanConnect.ps1 -username <NissanConnect username> -password <password> -country <CA|US> -update -climate_on -set_temp <integer in C> -climate_off -charge_on -locate -last_location -no_map -door_lock -door_unlock -pin_code
```

## Parameters
```
-username       : NissanConnect username [string]
-password       : NissanConnect password [string]
-country        : CA / US (CA is default) [string]
-update         : Refresh data [switch]
-climate_on     : Turn climate control on [switch]
-set_temp       : Set temperature for climate control (optional but requires -climate_on) [int]
-climate_off    : Turn climate control off [switch]
-charge_on      : Start charge [switch]
-locate         : Request location refresh [switch]
-last_location  : Get last recorded location [switch]
-no_map         : Do not open Google Maps on location request [switch]
-door_lock      : Request door lock (requires -pin_code) [switch]
-door_unlock    : Request door unlock (requires -pin_code) [switch]
-pin_code       : 4 digit pin code for door lock/unlock [string]
```

## Notes
**-climate_on** and **-climate_off** cannot be used at the same time (no action will be taken). **-locate** or **-last_location** will result in the default browser opening Google Maps to the location unless **-no_map** is specified. **-set_temp** is not required, however if used it only applies to **-climate_on**.

Cmdlet will return an object with the following properties:
- Account ID
- VIN
- Year
- Nickname
- BatteryCharge
- RangeClimateOn
- RangeClimateOff
- PlugState
- Charging
- InsideTemp
- LastUpdate
- Location (only if location option used)
  - Updated
  - Latitude
  - Longitude
  - Link


## Acknowledgements
Thanks to [Ben Woodford](https://gist.github.com/BenWoodford) for his initial documentation of the NissanConnect APIs.

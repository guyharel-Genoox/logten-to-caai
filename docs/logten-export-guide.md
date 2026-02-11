# How to Export from LogTen Pro

This guide explains how to create the tab-delimited export file that this tool requires.

## Steps

### LogTen Pro for Mac

1. Open LogTen Pro
2. Go to **File > Export**
3. Select **Tab-delimited** as the format
4. Select **All Flights** (or the date range you need)
5. Save the file (default name: `Export Flights (Tab) - YYYY-MM-DD HH-MM-SS.txt`)
6. Place the file in your project directory

### LogTen Pro for iOS/iPad

1. Open LogTen Pro
2. Tap the **More** menu
3. Tap **Export**
4. Select **Tab-delimited**
5. Share/save the file to your computer

## Required Fields

The export must include these LogTen Pro field names:

### Flight Data
- `flight_flightDate` - Flight date
- `flight_from` - Departure airport (ICAO)
- `flight_to` - Arrival airport (ICAO)
- `flight_route` - Route of flight
- `flight_totalTime` - Total flight time
- `flight_pic` - PIC time
- `flight_sic` - SIC time
- `flight_night` - Night time
- `flight_crossCountry` - Cross-country time
- `flight_actualInstrument` - Actual instrument time
- `flight_simulatedInstrument` - Simulated instrument time
- `flight_dualReceived` - Dual received time
- `flight_dualGiven` - Dual given time
- `flight_solo` - Solo time
- `flight_simulator` - Simulator time
- `flight_multiPilot` - Multi-pilot time
- `flight_distance` - Distance (if available)

### Landings & Approaches
- `flight_dayLandings`, `flight_dayTakeoffs`
- `flight_nightLandings`, `flight_nightTakeoffs`
- `flight_selectedApproach1` through `flight_selectedApproach4`
- `flight_holds` - Holding patterns
- `flight_ifr` - IFR time
- `flight_goArounds` - Go-arounds

### Crew
- `flight_selectedCrewPIC` - PIC name
- `flight_selectedCrewInstructor` - Instructor name
- `flight_selectedCrewStudent` - Student name
- `flight_selectedCrewObserver` - Observer/Safety Pilot

### Aircraft
- `aircraft_aircraftID` - Registration
- `aircraftType_type` - Type code (e.g., C172, A319)
- `aircraftType_make` - Manufacturer
- `aircraftType_model` - Model name
- `aircraftType_selectedEngineType` - Engine type
- `aircraftType_selectedCategory` - Category
- `aircraftType_selectedAircraftClass` - Class

### Other
- `flight_remarks` - Remarks (important: "safety pilot" detection)
- `aircraft_complex`, `aircraft_highPerformance`
- `aircraft_efis`, `aircraft_undercarriageRetractable`, `aircraft_pressurized`

## Tips

- Make sure **all flights** are exported, including simulator sessions
- The tool handles multiline remarks automatically
- Time values can be in H:MM format (e.g., "1:30") - they will be converted to decimal (1.50)
- Airport codes should be in ICAO format (4-letter codes like KFPR, LLBG)

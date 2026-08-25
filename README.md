# Amtrak Southwest Chief #3 Status Panel

A lightweight Windows desktop status panel for tracking **Amtrak Southwest Chief #3** as it approaches **Fort Madison, IA (FMD)** and **La Plata, MO (LAP)**.

The script retrieves current train information from the RailRat station pages and displays arrival estimates, delays, and countdowns in a small always-on-top Windows panel.

Designed to work with:

* Windows 7
* Windows PowerShell 2.0+
* Later Windows / PowerShell versions should also work
* No additional PowerShell modules required

## What It Looks For

The panel tracks westbound **Southwest Chief #3** at:

| Code | Station            |
| ---- | ------------------ |
| FMD  | Fort Madison, Iowa |
| LAP  | La Plata, Missouri |

It uses the individual RailRat station pages rather than the train-number page.

## Features

* Displays current Southwest Chief #3 status for FMD and LAP
* Shows estimated or actual arrival time
* Shows departure time when available
* Shows whether the train is early, on time, or late
* Shows a countdown when the train is approaching
* Automatically refreshes every 60 seconds
* Always-on-top compact desktop window
* Borderless window that can be dragged anywhere on the desktop
* Right-click menu with **Refresh** and **Exit**
* Press **Esc** to close the panel
* Falls back from .NET `WebClient` to Windows WinHTTP if necessary
* Designed to remain compatible with Windows PowerShell 2.0

## Arrival Warning Colors

As the estimated station arrival gets closer, the station card changes color:

| Time Until Arrival   | Color            |
| -------------------- | ---------------- |
| More than 60 minutes | Normal dark gray |
| 60 minutes or less   | Yellow           |
| 30 minutes or less   | Orange           |
| 10 minutes or less   | Red              |

When the train enters a new warning range, that station card flashes between the warning color and the normal background for approximately **15 seconds**.

The initial status does not flash.

## Running the Script

Download or clone the repository and run the `.ps1` file with Windows PowerShell.

For example:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Amtrak3-Status.ps1
```

Replace `Amtrak3-Status.ps1` with the actual filename if it is different.

### From PowerShell

You can also open PowerShell, change to the directory containing the script, and run:

```powershell
.\Amtrak3-Status.ps1
```

If PowerShell prevents local scripts from running because of the execution policy, you can launch it with:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Amtrak3-Status.ps1
```

This bypass applies only to that PowerShell process.

## Using the Panel

The panel opens near the upper-right corner of the primary display and stays above normal windows.

### Move it

Hold the **left mouse button** anywhere on the panel and drag it.

### Refresh it manually

**Right-click → Refresh**

### Exit

Either:

* **Right-click → Exit**
* Press **Esc**

## Information Displayed

An approaching train may look conceptually like:

```text
Southwest Chief #3

FMD  Fort Madison
     Est. arrival 21:42 CT (+38m)
     24m late - 8/24

LAP  La Plata
     Est. arrival 23:07 CT (+2h3m)
     18m late - 8/24
```

After an actual arrival has been reported, the panel switches from estimated arrival information to actual arrival/departure information.

For example:

```text
Arrived 21:47 CT
Departed 21:51 - 29m late
```

## Data Source

Train information is obtained from:

```text
https://railrat.net/stations/
```

The script reads the RailRat station pages for FMD and LAP and extracts the Southwest Chief #3 entries.

This project is not affiliated with or endorsed by Amtrak or RailRat.

Because the script depends on the current structure and contents of the RailRat pages, changes to that website may require changes to the parser.

## Network Requirements

The computer must have Internet access to RailRat.

The script attempts to enable TLS 1.2 where supported and uses two methods for retrieving the station pages:

1. .NET `System.Net.WebClient`
2. Windows `WinHttp.WinHttpRequest.5.1` as a fallback

This fallback behavior is particularly useful on older Windows installations.

## Requirements

The script uses standard Windows/.NET components:

```text
System.Windows.Forms
System.Drawing
System.Web
```

No third-party PowerShell modules are required.

## Configuration

Near the top of the script are the main configurable settings:

```powershell
$RefreshSeconds = 60
$StationBaseUrl = 'https://railrat.net/stations/'
```

To change the automatic refresh interval, modify:

```powershell
$RefreshSeconds = 60
```

The value is in seconds.

## Notes

The displayed station times are **Central Time**, because both Fort Madison and La Plata are in the U.S. Central Time Zone.

The countdown logic also evaluates the reported station arrival in Central Time rather than assuming the local timezone of the computer.

If the station page contains no usable Southwest Chief #3 entry, the panel displays:

```text
No #3 listed
Station page has no Southwest Chief #3
```

If a network or parsing error occurs, the affected station card displays:

```text
Unable to update
```

along with the underlying error information when available.

## Compatibility

The script was intentionally written using syntax and .NET functionality compatible with older PowerShell environments, including **Windows PowerShell 2.0 on Windows 7**.

That means it avoids dependencies on newer PowerShell-only features.

## Project Scope

This is deliberately a small utility rather than a general Amtrak tracking application.

Its job is simple:

> Keep Southwest Chief #3's progress through Fort Madison and La Plata visible on the desktop without having to repeatedly check a browser.

## License

Add the license you want to use for this repository here.

For a small open-source utility, the MIT License is a common choice.

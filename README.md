# NetDiscover - NIC Diagnostic Report (PowerShell)

This repository contains `Networking/NetDiscover.ps1`, an interactive PowerShell script that generates NIC (network interface) diagnostic reports for selected network adapters on Windows.

The script collects adapter properties, IP addresses, MTU, driver information, advanced adapter properties (for example, Jumbo Frame settings), and interface statistics (packets, errors, discards). It writes the results to CSV, TXT, and HTML files and provides an interactive menu to view the reports.

## Quick summary

- Script: `Networking/NetDiscover.ps1`
- Purpose: Generate human- and machine-readable NIC diagnostic reports for a selected network adapter.
- Platform: Windows (PowerShell). The script uses Windows networking cmdlets (`Get-NetAdapter`, `Get-NetIPAddress`, `Get-NetAdapterStatistics`, etc.) and calls some Win32 APIs to enable ANSI in the console.

## Features

- Interactive adapter selection menu showing each adapter name, description and status.
- Collects and reports:
	- Adapter name and description
	- Administrative status and link speed
	- MAC address
	- IPv4 and IPv6 addresses
	- MTU
	- Driver version and driver info
	- Advanced adapter properties (including Jumbo Frame settings when available)
	- Adapter statistics: RX/TX packets, errors and discards
- Outputs:
	- CSV file with a structured object suitable for automation or import into spreadsheets
	- Plain text report (readable summary)
	- Simple HTML report (preformatted text)
- Colorized console output with ANSI sequences and Windows console colors for clearer diagnostics.
- Simple 'open report' menu that attempts to open the generated CSV/TXT/HTML in common apps (Excel, Notepad++, Notepad, default browser) and the output folder in Explorer.
- Output files are timestamped and written to a configurable output folder.

## Usage

1. Open PowerShell on Windows (recommended: run with sufficient privileges if you expect to read driver/advanced properties).
2. Run the script:

```powershell
PS> .\Networking\NetDiscover.ps1
```

3. Choose an adapter from the presented list by entering its number.
4. After the report is generated you will see the paths to the CSV, TXT and HTML report files and an interactive prompt to open one or more of the files.

When finished you can re-run the report for the same adapter, pick another adapter, or quit.

## Configuration / Customization

- Default output folder is set at the top of the script as:

```powershell
$OutFolder = "C:\Users\Thoma\OneDrive\Documents\!_DIAGNOSTICS"
```

Change this path to a location appropriate for your environment before running the script.

- The script attempts to enable ANSI sequences for richer output and uses colors; consoles that don't support ANSI will still display the textual report but without the ANSI styling.

## Requirements

- Windows with PowerShell (script uses Windows-only networking cmdlets and Win32 calls).
- Recommended: PowerShell 7 (Core) or greater — the script works with Windows PowerShell but works best on PowerShell v7+ for consistent cross-version behavior.
- Terminal recommendation: use Windows Terminal (or any terminal that supports ANSI/VT sequences) for the intended ASCII/ANSI formatting and colorized output.
- The script uses the following cmdlets which are typically available on modern Windows versions:
	- `Get-NetAdapter`, `Get-NetIPAddress`, `Get-NetIPInterface`, `Get-NetAdapterStatistics`, `Get-NetAdapterAdvancedProperty`.
- When opening reports, the script tries application shortcuts such as `excel.exe`, Notepad++ and `notepad.exe`. If those aren't available, the attempt to open will fail silently or fall back where coded.

## Notes & Troubleshooting

- If the script prints "No network adapters found", confirm you are on a Windows host and that the networking modules are present.
- Reading advanced adapter properties may require elevated privileges; if the script cannot read advanced properties it will indicate that no advanced properties were found or access was denied.
- The script's Start-Process calls to open CSV/TXT/HTML files may fail silently if the target application is not installed. The TXT open operation attempts Notepad++ first, then Notepad.
- The script hardcodes the default output directory—change `\\$OutFolder` near the top if needed.
- ANSI/console color behavior may vary by terminal (Windows Terminal, conhost, PowerShell ISE, or remote sessions).

## Output files

- CSV: useful for automated parsing or importing into spreadsheets (Excel).
- TXT: human-readable summary with the same content as the HTML.
- HTML: minimal HTML wrapper around the text report; can be opened in a browser for quick viewing.

## Example workflow

1. Run `Networking/NetDiscover.ps1`.
2. Choose adapter 1 (for example).
3. After generation, open the CSV in Excel or open the TXT in Notepad++ to inspect driver properties.

## Extending the script

- You can modify the script to change the output folder, add more fields to the CSV object, or improve HTML formatting. Consider adding logging, or an option to run non-interactively by passing an interface index or name.

## License & Attribution

This repository does not include an explicit license file. If you plan to reuse or redistribute this script, add an appropriate license (for example, MIT) and ensure any internal documentation and identifiers are updated.

## Contact / Maintainer

Script location: `Networking/NetDiscover.ps1` — for questions or improvements, edit the script or open an issue in the repo.

---

Generated from reading the script in the `Networking` folder. If you want, I can also:

- Add a small non-interactive wrapper to run the report from CI or a scheduled task.
- Add a `-OutFolder` parameter to the main loop so the output directory doesn't need to be edited inline.

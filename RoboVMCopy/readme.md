# RoboVMCopy

Windows PowerShell GUI wrapper for RoboCopy.

RoboVMCopy lets you:

- Pick any source folder and destination folder on Windows
- Browse with a dialog that can access local and mapped drives
- Toggle common RoboCopy options in a GUI
- Run and stop copy jobs with live per-file output
- Show the current file activity in a dedicated live status area
- Show a visual overall progress bar tied to the live copy metrics
- Show elapsed time, estimated completion time, and throughput in the status bar
- Save each run output to a timestamped log file in a Logs sub-folder beside the script
- Save copy jobs as quick-launch buttons
- Edit or delete saved jobs from the context menu
- Re-run saved jobs from left-click confirmation or context menu

## Files

- `RoboVMCopy.ps1` - GUI application
- `RoboVMCopy_Jobs.json` - saved jobs (created automatically after first save)

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1+
- `robocopy.exe` (included with Windows)
- Script must run in STA mode

## Run

From this folder:

```powershell
powershell -STA -ExecutionPolicy Bypass -File .\RoboVMCopy.ps1
```

If you run from another location, pass the full path:

```powershell
powershell -STA -ExecutionPolicy Bypass -File "C:\Path\To\RoboVMCopy.ps1"
```

## How To Use

1. Click **Browse...** for Source.
2. Click **Browse...** for Destination.
3. Select RoboCopy options.
4. Click **Run RoboCopy**.
5. Watch progress in **Output Log**.

The **Current File** area shows the latest file activity during the run.

The **Overall Progress** bar uses the same metric model as the status bar, so percentage, elapsed time, ETA, and throughput stay aligned.

The status bar shows elapsed time, estimated completion time, and throughput in Mbps while the copy is running.

To stop a running copy, click **Stop**.

The application stays open after a job finishes. It only closes when you click **Exit** and confirm.

## Save Jobs To Buttons

1. Set Source and Destination.
2. Set options and extra flags.
3. Enter a name in **Save as button**.
4. Click **Save Job**.

Saved jobs appear in the **Saved Jobs** panel.

- Left click saved button: loads the job and prompts "Run job now?"
- Right click saved button: load paths, edit, run, or delete job

To edit a saved job:

1. Right click the saved job button.
2. Click **Edit job**.
3. Update values in the main form.
4. Click **Save Job** to overwrite.

Delete prompts for confirmation before removal.

## Run Logs

- Every run writes output to `Logs/RoboCopy_yyyyMMdd_HHmmss.log` under the same folder as `RoboVMCopy.ps1`
- The full path is shown in the output log when a run starts and when it finishes

## Built-In Option Controls

Checkbox options:

- `/MIR` - mirror source to destination
- `/E` - include empty directories
- `/COPYALL` - copy all file metadata
- `/Z` - restartable mode
- `/XA:H` - exclude hidden files
- `/PURGE` - remove destination items not in source

Numeric options:

- `/MT:n` - thread count
- `/R:n` - retries
- `/W:n` - wait time between retries (seconds)

Extra flags:

- Add any additional RoboCopy flags in the **Extra flags** box
- Example: `/XF *.tmp /XD Temp Cache`

## RoboCopy Exit Code Notes

- `0-7` are generally success or informational outcomes
- `8+` indicates copy errors

The status bar and output log show success and error states.

Per-file RoboCopy output is always shown in the GUI, similar to the CLI view.

The dedicated **Current File** panel highlights the most recent file activity without needing to scan the full log.

The overall progress bar provides a quick visual estimate of how far the copy has advanced.

The elapsed time and ETA are estimates based on the copy activity seen so far.

## Troubleshooting

### Parser errors with strange characters (for example: `aEUR"`, `aEURo`, `aEUR`)

Cause:

- The script was saved with the wrong text encoding and Unicode characters were corrupted.

Fix:

- Use the ASCII-safe script in this folder
- Save script files as UTF-8 if you edit them
- Re-copy the latest `RoboVMCopy.ps1` if you previously used a corrupted copy

### Script says it must run in STA mode

Run with:

```powershell
powershell -STA -ExecutionPolicy Bypass -File .\RoboVMCopy.ps1
```

### Source path does not exist

- Confirm the source folder still exists
- Re-browse and select the folder again

### Access denied or file locked errors

- Run PowerShell as Administrator if needed
- Close apps using those files
- Add suitable RoboCopy retry flags

### COPYALL fails with "Manage Auditing user right"

Cause:

- `/COPYALL` includes auditing information (`U`) and needs the Manage Auditing user right.

Fix:

- When prompted, choose **Yes** to run with `/COPY:DATS` for that run.
- Or edit your saved job and remove `/COPYALL` if auditing copy is not required.

## Safety Notes

- `/MIR` and `/PURGE` can delete files from destination.
- Test with a non-critical destination first.
- Review the output log before using jobs on production data.

## Suggested Enhancements

- Add "Test Run" support with `/L`
- Add per-job include/exclude templates
- Add export/import for saved jobs
- Add scheduled-task integration for unattended runs

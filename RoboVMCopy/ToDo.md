# To Do

- [x] Add delete and edit saved jobs
- [x] Show robocopy output and save a log with date and time appended to filename
- [x] Stop the application from exiting after copy
- [x] Find a way to show mapped drives on local PC when browsing
- [x] Gui exits after starting copy from saved copy process
  - [x] I should only exit the gui when I click exit.
    - [x] when clicking exit it should prompt if I really want to exit.
- [x] The hard copy of the log should be saved in the same folder as this powerscript executable
- [x] an extra blank saved copy process button is created when I save one
- [x] I want to file copying process in the gui for each file.. just like is seen in cli when runniong robocopy
- [x] Send logs to the 'Logs' sub-folder
- [x] Update throughput ever30 seconds
  - the number is displayed in hundredths, 0.01, change to whole number to tenths.
  - For example, currently a value of 0.02 would show as 20.0
- [x] ETA keep changing so much it is useless. Is this in time of day? Should be in 12-hr format
  - maybe it should only be displayed as an average over every minute.
- [x] Overall Progress bar is not working
- [x] Remove the words "The app stays open until you click Exit."
- [x] the Saved Job name field is overlapping Save Job button
- [x] When a saved job is clicked, load job information
  - The prompt and ask "Run job now?"
- [x] Make Output Log box a collapsible section. Default open for 10 seconds after run, then collapse.
- [x] In the output log box, instead of displaying percentages complete, just refresh the percentage number in place and not put it on a new line? Like it shows up in terminal.
- In the summary at the end:
  - [x] print the filenames of any files Skipped, Mismatch, FAILED, Extras
- [x] Add hover-over tool tips explaining each Robocopy Option
- [x] Also in summary, what are the values in Times row?
<!-- EOF -->


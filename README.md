# Extract IPR

This script extracts the IPR for all producer wells from a PRT file, then calculates and compiles the phase PIs into a spreadsheet.

## How to use this script

1. Specify the wells that you want to have their IPRs reported using the WellIPRReport function (example: Well_IPR_example.ixf)
2. Run the model to get a PRT file containing all the IPRs
3. In Extract_IPR.py, set the variable prt_file to the directory of the PRT file
4. Run the script

Warning: Running this script may take a long time - with a run that reports monthly data point for 15 years, it takes around 1 minute for one well, but more than 3 hours for 300 wells.
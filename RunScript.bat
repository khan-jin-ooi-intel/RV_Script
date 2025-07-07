@echo OFF
setlocal

choice /c yn /m "Run Script"
:: Set input arguments
set "DIRECTORY=C:\PythonScripts"
set "INPUT_FILE=TestData\vmax.csv"
set "OUTPUT_FILE=TestData\vmax_processed.xlsx"
set "VID_LIST=U5W56U6400101"

\\gar.corp.intel.com\ec\proj\ba\aqua\AquaHbase\AquaCMDClient\Client\AquaCmdLine.exe -aquaServer GAR -reportPath "khanjino\BMG_ByLot_VMIN_pythonscript" -outputFileName "%DIRECTORY%\%INPUT_FILE%" -visualIds "%VID_LIST%" -segregateRunWithFailures -sendmail

:: Run the Python script with arguments
python RV_prelimauto_v2.4.py --inputfile "%INPUT_FILE%" --outputfile "%OUTPUT_FILE%" --vid "%VID_LIST%"

endlocal

pause

:: =================== INSTRUCTIONS =====================
:: 1) Create empty folder and point "DIRECTORY" to folder
:: 2) Specify "INPUT_FILE" name *ends with .csv*
:: 3) Specify "OUTPUT_FILE" name *ends with .xlsx*
:: 4) Specify VID/VIDs to pull ex.123,111,333
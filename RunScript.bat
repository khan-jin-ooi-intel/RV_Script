@echo OFF
setlocal
cls
:: =================== INSTRUCTIONS =====================
:: 1) Point "DIRECTORY" to Output folder (one will be created if does not exist)
:: 2) Specify "INPUT_FILE" name *ends with .csv*
:: 3) Specify "OUTPUT_FILE" name *ends with .xlsx*
:: 4) Specify VID/VIDs to pull ex.123,111,333

echo ************** Welcome to KJ's RV Script ************************
echo(

:: Set input arguments
set "DIRECTORY=C:\PythonScripts"
set "FORMAT=datatopull_v2.xlsx"
set "AQUA_REPORT=\TestData\w28test.csv"
set "OUTPUT_FILE=\TestData\w28test_processed.xlsx"
set "VID_LIST=U5W56U6400101,U5W56U6400102"
set "LOCN_LIST=119325,6261,6212"

if not exist "%DIRECTORY%\" (
	echo Folder not found. Creating Folder...
	mkdir "%DIRECTORY%"
)

choice /c yn /n /m "Is Your Aqua Report Available (Y/N)?"

if errorlevel 2 goto no_section
if errorlevel 1 goto yes_section

:no_section
setlocal enabledelayedexpansion

<nul set /p=Pulling Aqua Report
::printing the dots 
for /L %%i in (1,1,3) do (
	<nul set /p=.
	timeout /t 1 >nul
)
setlocal disabledelayedexpansion

\\gar.corp.intel.com\ec\proj\ba\aqua\AquaHbase\AquaCMDClient\Client\AquaCmdLine.exe -aquaServer GAR -reportPath "khanjino\BMG_ByLot_VMIN_pythonscript" -outputFileName "%DIRECTORY%\%AQUA_REPORT%" -visualIds "%VID_LIST%" -segregateRunWithFailures -sendmail

echo ....................
echo Aqua Report Pulled!
goto end

:yes_section
goto end

:end
:: Run the Python script with arguments
python RV_prelim_v2.4.1_auto.py --directory "%DIRECTORY%" --aquareport "%AQUA_REPORT%" --outputfile "%OUTPUT_FILE%" --format %FORMAT% --vid "%VID_LIST%" --locn "%LOCN_LIST%"

endlocal

pause

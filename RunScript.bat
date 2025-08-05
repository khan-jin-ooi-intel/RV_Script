@echo OFF
setlocal
cls
:: =================== INSTRUCTIONS =====================
:: 1) Specify path to "INPUT_FILE" [*ends with .csv]
:: 2) Specify path to "OUTPUT_FILE" [*ends with .xlsx] (optional, will default to input file path if not defined)
:: 3) Specify VID/VIDs to pull ex.123,111,333
:: 4) Specify LOCATION(S) to pull for [current ver. only works for 119325(sort),6261(classhot),6212(classcold),6242(qahot),6243(qacold)]

type welcome_art.txt
echo.
echo.
:: Set input arguments
set "INPUT_FILE=C:\PythonScripts\Test\w31_facr.csv"
set "OUTPUT_FILE=C:\PythonScripts\Test\w31_facr_processed.xlsx"
set "FORMAT=C:\PythonScripts\format.xlsx"
set "VID_LIST=U4JQ199702965"
set "LOCN_LIST=119325,6261,6212,5242,5243"

:: Prerequisites Check for Python and Pandas
python --version >nul 2>&1
if errorlevel 1 (
	echo Python not installed or defined in enviroment path
	goto end
)

python -c "import pandas" >nul 2>&1 
if errorlevel 1 (
	echo Pandas not found. Installing...
	pip install pandas
)else (
	echo Pandas found. Skipping installation...
) 

:: Find parent folder defined in INPUT_FILE and create folder if not exist
for %%F in ("%INPUT_FILE%") do set "PARENT=%%~dpF"
if not exist "%PARENT%" (
	echo Folder not found. Creating Folder...
	mkdir %PARENT%
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

\\gar.corp.intel.com\ec\proj\ba\aqua\AquaHbase\AquaCMDClient\Client\AquaCmdLine.exe -aquaServer GAR -reportPath "khanjino\BMG_ByLot_VMIN_pythonscript" -outputFileName "%INPUT_FILE%" -visualIds "%VID_LIST%" -segregateRunWithFailures -sendmail
echo ....................
echo Aqua Report Pulled!
goto yes_section

:yes_section
:: Run the Python script with arguments (--dump)
python RV_prelimauto_v2.4.1.py --inputfile "%INPUT_FILE%" --outputfile "%OUTPUT_FILE%" --format %FORMAT% --vid "%VID_LIST%" --locn "%LOCN_LIST%"
goto end

:end
endlocal
pause

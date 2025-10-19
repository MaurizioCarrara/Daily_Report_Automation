@echo off
cls
color 2
echo Benvenuto!
echo Questo script avvia il workflow ETL di Knime e aggiorna i file Excel.
echo Assicurati che KNIME sia installato nel tuo spazio personale! :)
echo.

:: Verifica accesso all'unità di rete
if not exist "R:\" (
    echo Errore: Impossibile accedere all'unità R:.
    echo Verifica che l'unità di rete sia mappata e accessibile.
    pause
    exit /b 1
)

:: Verifica delle impostazioni di sicurezza di Excel
set VBS_CHECK="%TEMP%\check_excel_security.vbs"
echo On Error Resume Next > %VBS_CHECK%
echo Set objExcel = CreateObject("Excel.Application") >> %VBS_CHECK%

echo If Err.Number ^<^> 0 Then >> %VBS_CHECK%
echo     WScript.Echo "ERRORE_EXCEL" >> %VBS_CHECK%
echo     WScript.Quit 1 >> %VBS_CHECK%
echo End If >> %VBS_CHECK%

echo objExcel.Visible = False >> %VBS_CHECK%
echo Set objVBProject = objExcel.VBE.ActiveVBProject >> %VBS_CHECK%

echo If Err.Number ^<^> 0 Then >> %VBS_CHECK%
echo     WScript.Echo "ERRORE_TRUST" >> %VBS_CHECK%
echo     objExcel.Quit >> %VBS_CHECK%
echo     WScript.Quit 2 >> %VBS_CHECK%
echo End If >> %VBS_CHECK%

echo objExcel.Quit >> %VBS_CHECK%
echo WScript.Echo "OK" >> %VBS_CHECK%

:: Esegui il controllo
for /f "delims=" %%i in ('cscript //nologo %VBS_CHECK%') do set EXCEL_CHECK=%%i
del %VBS_CHECK%

if "%EXCEL_CHECK%"=="ERRORE_EXCEL" (
    echo Errore: Impossibile avviare Excel.
    echo Verifica che Excel sia installato correttamente.
    pause
    exit /b 1
)

if "%EXCEL_CHECK%"=="ERRORE_TRUST" (
    echo Errore: Le impostazioni di sicurezza di Excel non consentono l'accesso al VBA.
    echo Per risolvere:
    echo 1. Apri Excel
    echo 2. Vai su File ^> Opzioni ^> Centro protezione ^> Impostazioni Centro protezione
    echo 3. Seleziona "Impostazioni macro"
    echo 4. Spunta "Considera attendibile l'accesso al modello a oggetti dei progetti VBA"
    echo 5. Clicca OK e riavvia Excel
    pause
    exit /b 1
)

:: Trova il percorso di KNIME
set KNIME_PATH=%USERPROFILE%\AppData\Local\Programs\KNIME\knime.exe
if not exist "%KNIME_PATH%" (
    echo Errore: KNIME non trovato in %KNIME_PATH%.
    echo Controlla che KNIME sia installato nel percorso corretto.
    pause
    exit /b 1
)

:: Verifica esistenza workflow KNIME
set WORKFLOW_PATH="R:\AutomationFolder\KNIME\DataProcessing.knwf"
if not exist %WORKFLOW_PATH% (
    echo Errore: Workflow non trovato in %WORKFLOW_PATH%.
    echo Controlla il nome e il percorso del workflow.
    pause
    exit /b 1
)

:: Esegui il workflow KNIME
echo Avvio di KNIME con workflow 'DataProcessing'...

"%KNIME_PATH%" -reset -nosplash -application org.knime.product.KNIME_BATCH_APPLICATION -workflowFile=%WORKFLOW_PATH%
if errorlevel 1 (
    echo Errore durante l'esecuzione del workflow KNIME.
    pause
    exit /b 1
)

echo Workflow DataProcessing completato correttamente!

:: Ottieni anno, mese e giorno correnti
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set datetime=%%I
set AA=%datetime:~2,2%
set MM=%datetime:~4,2%
set GG=%datetime:~6,2%

:: Calcola il mese e anno successivo
::set /a MM_NEXT=%MM%+1
::set /a AA_NEXT=%AA%
::if %MM_NEXT% gtr 12 (
::    set MM_NEXT=01
::    set /a AA_NEXT=%AA_NEXT%+1
::)
:: Aggiungi lo zero iniziale se necessario per il mese successivo
::if %MM_NEXT% lss 10 set MM_NEXT=0%MM_NEXT%

:: Calcola mese e anno successivo in modo sicuro (formati yy e MM)
for /f "tokens=1,2 delims=;" %%A in ('
   powershell -command "(Get-Date).AddMonths(1).ToString('yy;MM')"
') do (
   set AA_NEXT=%%A
   set MM_NEXT=%%B
)

:: Imposta i percorsi
set BASE_PATH=R:\AutomationFolder
set EXCEL_FILE_1=%BASE_PATH%\DataComparison.xlsx
set EXCEL_FILE_2=%BASE_PATH%\DataComparisonNext.xlsx
set BAS_FILE=%BASE_PATH%\VBA\DataProcessing.bas

:: Verifica esistenza file necessari
if not exist "%EXCEL_FILE_1%" (
    echo Errore: File Excel DataComparison non trovato in %EXCEL_FILE_1%
    pause
    exit /b 1
)

if not exist "%EXCEL_FILE_2%" (
    echo Errore: File Excel DataComparisonNext non trovato in %EXCEL_FILE_2%
    pause
    exit /b 1
)

if not exist "%BAS_FILE%" (
    echo Errore: File BAS non trovato in %BAS_FILE%
    pause
    exit /b 1
)



:: Crea e esegui lo script VBS per il primo file
set VBA_SCRIPT="%TEMP%\temp_vba_script.vbs"

echo Option Explicit > %VBA_SCRIPT%
echo On Error Resume Next >> %VBA_SCRIPT%
echo Dim objExcel, objWorkbook, FSO, errNum, errDesc >> %VBA_SCRIPT%

:: Inizializzazione FSO
echo Set FSO = CreateObject("Scripting.FileSystemObject") >> %VBA_SCRIPT%
echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore: Impossibile creare FileSystemObject" >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

:: Inizializzazione Excel
echo Set objExcel = CreateObject("Excel.Application") >> %VBA_SCRIPT%
echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore: Impossibile avviare Excel." >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo objExcel.Visible = False >> %VBA_SCRIPT%
echo objExcel.DisplayAlerts = False >> %VBA_SCRIPT%

:: Primo file
echo WScript.Echo "Elaborazione primo file..." >> %VBA_SCRIPT%
echo Set objWorkbook = objExcel.Workbooks.Open("%EXCEL_FILE_1%") >> %VBA_SCRIPT%

echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore nell'apertura del primo file Excel" >> %VBA_SCRIPT%
echo     objExcel.Quit >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo objExcel.VBE.ActiveVBProject.VBComponents.Import "%BAS_FILE%" >> %VBA_SCRIPT%
echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore durante l'importazione del modulo VBA" >> %VBA_SCRIPT%
echo     objWorkbook.Close False >> %VBA_SCRIPT%
echo     objExcel.Quit >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo objWorkbook.Application.Run "DataComparison.AVVIO" >> %VBA_SCRIPT%
echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore durante l'esecuzione della macro" >> %VBA_SCRIPT%
echo     objWorkbook.Close False >> %VBA_SCRIPT%
echo     objExcel.Quit >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo WScript.Echo "Salvataggio primo file..." >> %VBA_SCRIPT%
echo Dim newFile1 >> %VBA_SCRIPT%
echo newFile1 = "%BASE_PATH%\%AA%%MM% - Report Giornaliero.xlsx" >> %VBA_SCRIPT%
echo If FSO.FileExists(newFile1) Then >> %VBA_SCRIPT%
echo     FSO.DeleteFile(newFile1) >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo objWorkbook.SaveAs newFile1, 51 >> %VBA_SCRIPT%
echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore durante il salvataggio del primo file" >> %VBA_SCRIPT%
echo     objWorkbook.Close False >> %VBA_SCRIPT%
echo     objExcel.Quit >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo objWorkbook.Close >> %VBA_SCRIPT%

:: Secondo file
echo WScript.Echo "Elaborazione secondo file..." >> %VBA_SCRIPT%
echo Set objWorkbook = objExcel.Workbooks.Open("%EXCEL_FILE_2%") >> %VBA_SCRIPT%
echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore nell'apertura del secondo file Excel" >> %VBA_SCRIPT%
echo     objExcel.Quit >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo objExcel.VBE.ActiveVBProject.VBComponents.Import "%BAS_FILE%" >> %VBA_SCRIPT%
echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore durante l'importazione del modulo VBA" >> %VBA_SCRIPT%
echo     objWorkbook.Close False >> %VBA_SCRIPT%
echo     objExcel.Quit >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo objWorkbook.Application.Run "DataComparison.AVVIO" >> %VBA_SCRIPT%
echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore durante l'esecuzione della macro" >> %VBA_SCRIPT%
echo     objWorkbook.Close False >> %VBA_SCRIPT%
echo     objExcel.Quit >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo WScript.Echo "Salvataggio secondo file..." >> %VBA_SCRIPT%
echo Dim newFile2 >> %VBA_SCRIPT%
echo newFile2 = "%BASE_PATH%\%AA_NEXT%%MM_NEXT% - Report Giornaliero.xlsx" >> %VBA_SCRIPT%
echo If FSO.FileExists(newFile2) Then >> %VBA_SCRIPT%
echo     FSO.DeleteFile(newFile2) >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo objWorkbook.SaveAs newFile2, 51 >> %VBA_SCRIPT%
echo If Err.Number ^<^> 0 Then >> %VBA_SCRIPT%
echo     WScript.Echo "Errore durante il salvataggio del secondo file" >> %VBA_SCRIPT%
echo     objWorkbook.Close False >> %VBA_SCRIPT%
echo     objExcel.Quit >> %VBA_SCRIPT%
echo     WScript.Quit 1 >> %VBA_SCRIPT%
echo End If >> %VBA_SCRIPT%

echo objWorkbook.Close >> %VBA_SCRIPT%
echo objExcel.Quit >> %VBA_SCRIPT%
echo Set objWorkbook = Nothing >> %VBA_SCRIPT%
echo Set objExcel = Nothing >> %VBA_SCRIPT%
echo Set FSO = Nothing >> %VBA_SCRIPT%

echo WScript.Echo "OK" >> %VBA_SCRIPT%

:: Esecuzione dello script VBA
echo Esecuzione dello script VBA...
cscript //nologo %VBA_SCRIPT%
if errorlevel 1 (
    echo Errore durante l'esecuzione del file VBScript.
    pause
    exit /b 1
)

:: Pulizia file temporaneo
del %VBA_SCRIPT%

echo.
echo Files salvati come:
echo %AA%%MM% - Report Giornaliero.xlsx
echo %AA_NEXT%%MM_NEXT% - Report Giornaliero.xlsx
echo Operazione completata con successo!

pause

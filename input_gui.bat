@echo off
setlocal enabledelayedexpansion

:: === Configuration ===
set "output_folder=C:\dpool\in"
set "temp_folder=%temp%"

:: Create output folder if it doesn't exist
if not exist "%output_folder%" (
    mkdir "%output_folder%"
    echo Created output folder: %output_folder%
)

:: === Main Menu ===
:MAIN_MENU
cls
echo ================================================
echo    DOCUMENT GENERATOR
echo ================================================
echo.
echo 1. Generate Invoice
echo 2. Generate Credit Note
echo 3. Exit
echo.
set /p "menu_choice=Please select an option (1-3): "

if "%menu_choice%"=="1" goto :INVOICE_INPUT
if "%menu_choice%"=="2" goto :CREDITNOTE_INPUT
if "%menu_choice%"=="3" goto :EXIT
echo Invalid choice. Please try again.
pause
goto :MAIN_MENU

:: === INVOICE SECTION ===
:INVOICE_INPUT
cls
echo ================================================
echo    INVOICE GENERATOR
echo ================================================
echo.
echo Creating invoice input form...

:: Create simple VBS script for invoice input
set "invoice_vbs=%temp_folder%\invoice_input.vbs"
echo Dim customer, hscode, product, qty, amount, tax > "%invoice_vbs%"
echo Dim timestamp, filename, fso, file >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo ' Get customer name >> "%invoice_vbs%"
echo customer = InputBox("Enter Customer Name:", "Invoice - Customer") >> "%invoice_vbs%"
echo If customer = "" Then >> "%invoice_vbs%"
echo     WScript.Quit >> "%invoice_vbs%"
echo End If >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo ' Get HS Code (optional) >> "%invoice_vbs%"
echo hscode = InputBox("Enter HS Code (optional - leave blank if none):", "Invoice - HS Code") >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo ' Get product name >> "%invoice_vbs%"
echo product = InputBox("Enter Product Name:", "Invoice - Product") >> "%invoice_vbs%"
echo If product = "" Then >> "%invoice_vbs%"
echo     WScript.Quit >> "%invoice_vbs%"
echo End If >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo ' Get quantity >> "%invoice_vbs%"
echo qty = InputBox("Enter Quantity:", "Invoice - Quantity", "1") >> "%invoice_vbs%"
echo If qty = "" Or Not IsNumeric(qty) Or qty ^<= 0 Then >> "%invoice_vbs%"
echo     WScript.Quit >> "%invoice_vbs%"
echo End If >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo ' Get amount >> "%invoice_vbs%"
echo amount = InputBox("Enter Unit Amount:", "Invoice - Amount", "0.00") >> "%invoice_vbs%"
echo If amount = "" Or Not IsNumeric(amount) Or amount ^< 0 Then >> "%invoice_vbs%"
echo     WScript.Quit >> "%invoice_vbs%"
echo End If >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo ' Get tax class >> "%invoice_vbs%"
echo tax = InputBox("Enter Tax Class (e.g., v1, v3, exempt):", "Invoice - Tax", "v3") >> "%invoice_vbs%"
echo If tax = "" Then >> "%invoice_vbs%"
echo     WScript.Quit >> "%invoice_vbs%"
echo End If >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo ' Generate timestamp >> "%invoice_vbs%"
echo timestamp = Year(Now) ^& Right("0" ^& Month(Now), 2) ^& Right("0" ^& Day(Now), 2) ^& "_" ^& Right("0" ^& Hour(Now), 2) ^& Right("0" ^& Minute(Now), 2) >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo ' Create output file >> "%invoice_vbs%"
echo Set fso = CreateObject("Scripting.FileSystemObject") >> "%invoice_vbs%"
echo If Not fso.FolderExists("C:\dpool\in") Then >> "%invoice_vbs%"
echo     fso.CreateFolder "C:\dpool\in" >> "%invoice_vbs%"
echo End If >> "%invoice_vbs%"
echo filename = "C:\dpool\in\invoice_" ^& timestamp ^& ".txt" >> "%invoice_vbs%"
echo Set file = fso.CreateTextFile(filename, True) >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo ' Write invoice data >> "%invoice_vbs%"
echo file.WriteLine "=== INVOICE ===" >> "%invoice_vbs%"
echo file.WriteLine "Generated: " ^& Now >> "%invoice_vbs%"
echo file.WriteLine "Customer: " ^& customer >> "%invoice_vbs%"
echo If hscode ^<^> "" Then >> "%invoice_vbs%"
echo     file.WriteLine "r_trp """ ^& hscode ^& " " ^& product ^& """ " ^& qty ^& " * " ^& amount ^& " " ^& tax >> "%invoice_vbs%"
echo Else >> "%invoice_vbs%"
echo     file.WriteLine "r_trp """ ^& product ^& """ " ^& qty ^& " * " ^& amount ^& " " ^& tax >> "%invoice_vbs%"
echo End If >> "%invoice_vbs%"
echo file.WriteLine "Total: " ^& (CDbl(qty) * CDbl(amount)) >> "%invoice_vbs%"
echo file.Close >> "%invoice_vbs%"
echo. >> "%invoice_vbs%"
echo MsgBox "Invoice saved successfully!" ^& vbCrLf ^& "File: " ^& filename, vbInformation, "Success" >> "%invoice_vbs%"

echo Starting invoice input dialog...
cscript //nologo "%invoice_vbs%"

if errorlevel 1 (
    echo Invoice generation was cancelled.
) else (
    echo Invoice generated successfully!
)

echo.
echo Press any key to return to main menu...
pause >nul
goto :MAIN_MENU

:: === CREDIT NOTE SECTION ===
:CREDITNOTE_INPUT
cls
echo ================================================
echo    CREDIT NOTE GENERATOR
echo ================================================
echo.
echo Creating credit note input form...

:: Create simple VBS script for credit note input
set "creditnote_vbs=%temp_folder%\creditnote_input.vbs"
echo Dim dcn, qty, amount, tax > "%creditnote_vbs%"
echo Dim timestamp, filename, fso, file >> "%creditnote_vbs%"
echo. >> "%creditnote_vbs%"
echo ' Get DCN number >> "%creditnote_vbs%"
echo dcn = InputBox("Enter Credit note Number:", "Credit Note - DCN") >> "%creditnote_vbs%"
echo If dcn = "" Then >> "%creditnote_vbs%"
echo     WScript.Quit >> "%creditnote_vbs%"
echo End If >> "%creditnote_vbs%"
echo. >> "%creditnote_vbs%"
echo ' Get quantity >> "%creditnote_vbs%"
echo qty = InputBox("Enter Quantity:", "Credit Note - Quantity", "1") >> "%creditnote_vbs%"
echo If qty = "" Or Not IsNumeric(qty) Or qty ^<= 0 Then >> "%creditnote_vbs%"
echo     WScript.Quit >> "%creditnote_vbs%"
echo End If >> "%creditnote_vbs%"
echo. >> "%creditnote_vbs%"
echo ' Get amount >> "%creditnote_vbs%"
echo amount = InputBox("Enter Credit Amount:", "Credit Note - Amount", "0.00") >> "%creditnote_vbs%"
echo If amount = "" Or Not IsNumeric(amount) Or amount ^< 0 Then >> "%creditnote_vbs%"
echo     WScript.Quit >> "%creditnote_vbs%"
echo End If >> "%creditnote_vbs%"
echo. >> "%creditnote_vbs%"
echo ' Get tax class >> "%creditnote_vbs%"
echo tax = InputBox("Enter Tax Class (e.g., v1, v3, exempt):", "Credit Note - Tax", "v1") >> "%creditnote_vbs%"
echo If tax = "" Then >> "%creditnote_vbs%"
echo     WScript.Quit >> "%creditnote_vbs%"
echo End If >> "%creditnote_vbs%"
echo. >> "%creditnote_vbs%"
echo ' Generate timestamp >> "%creditnote_vbs%"
echo timestamp = Year(Now) ^& Right("0" ^& Month(Now), 2) ^& Right("0" ^& Day(Now), 2) ^& "_" ^& Right("0" ^& Hour(Now), 2) ^& Right("0" ^& Minute(Now), 2) >> "%creditnote_vbs%"
echo. >> "%creditnote_vbs%"
echo ' Create output file >> "%creditnote_vbs%"
echo Set fso = CreateObject("Scripting.FileSystemObject") >> "%creditnote_vbs%"
echo If Not fso.FolderExists("C:\dpool\in") Then >> "%creditnote_vbs%"
echo     fso.CreateFolder "C:\dpool\in" >> "%creditnote_vbs%"
echo End If >> "%creditnote_vbs%"
echo filename = "C:\dpool\in\creditnote_" ^& timestamp ^& ".txt" >> "%creditnote_vbs%"
echo Set file = fso.CreateTextFile(filename, True) >> "%creditnote_vbs%"
echo. >> "%creditnote_vbs%"
echo ' Write credit note data >> "%creditnote_vbs%"
echo file.WriteLine "r_dcn """ ^& dcn ^& """" >> "%creditnote_vbs%"
echo file.WriteLine "r_trp ""total"" " ^& qty ^& " * " ^& amount ^& " " ^& tax >> "%creditnote_vbs%"
echo file.Close >> "%creditnote_vbs%"
echo. >> "%creditnote_vbs%"
echo MsgBox "Credit Note saved successfully!" ^& vbCrLf ^& "File: " ^& filename, vbInformation, "Success" >> "%creditnote_vbs%"

echo Starting credit note input dialog...
cscript //nologo "%creditnote_vbs%"

if errorlevel 1 (
    echo Credit Note generation was cancelled.
) else (
    echo Credit Note generated successfully!
)

echo.
echo Press any key to return to main menu...
pause >nul
goto :MAIN_MENU

:EXIT
echo.
echo Cleaning up temporary files...
if exist "%temp_folder%\invoice_input.vbs" del "%temp_folder%\invoice_input.vbs"
if exist "%temp_folder%\creditnote_input.vbs" del "%temp_folder%\creditnote_input.vbs"
echo.
echo Thank you for using Document Generator!
pause
exit /b 0
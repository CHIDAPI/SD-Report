@echo off
 
:: Execute the PS1 file with the same name as this batch file.
set filename=EmailMove.ps1
 
if exist "%filename%" (
  PowerShell.exe -NoProfile -NonInteractive -ExecutionPolicy unrestricted -File "%filename%" 
 
  :: Collect the exit code from the PowerShell script.
  set err=%errorlevel%
) else (
  echo File not found.
  echo %filename%
 
  :: Set our exit code.
  set err=1
)
 
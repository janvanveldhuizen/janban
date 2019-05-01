@echo off
if not exist %USERPROFILE%\janban (mkdir %USERPROFILE%\janban)
if not exist kanban.html (
      echo The install script is running in the wrong folder
      pause
      goto :EOF
)
robocopy /mir . %USERPROFILE%\janban 
set offver=0
for /l %%x in (12, 1, 19) do (
      reg export HKCU\Software\Microsoft\Office\%%x.0\Outlook\Today %%x /y
	  if exist %%x set offver=%%x
)
if %offver%==0 (
      echo  The install script could not detect your Office version. Please install manually.
      pause
      goto :EOF
)
if not %offver%==0 (
      del %offver%
      reg add HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today /v Stamp /t REG_DWORD /d 1 /f
      reg add HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today /v UserDefinedUrl /t REG_SZ /d %USERPROFILE%\janban\kanban.html /f
      goto :EOF
)

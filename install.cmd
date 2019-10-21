@echo off
set offver=0
for /l %%x in (12, 1, 19) do (
      reg export HKCU\Software\Microsoft\Office\%%x.0\Outlook\Today %%x /y
	  if exist %%x set offver=%%x
)
cls
if %offver%==0 (
      echo  The install script could not detect your Office version. Please install manually.
      pause
      goto :EOF
)
if not %offver%==0 (
      del %offver%
      reg add HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today /v Stamp /t REG_DWORD /d 1 /f
      reg add HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today /v UserDefinedUrl /t REG_SZ /d "https://janware.nl/janban" /f
      cls
      echo  JanBan successfully set up. Have fun.
      pause
      goto :EOF
)

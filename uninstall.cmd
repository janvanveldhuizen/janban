@echo off
if exist %USERPROFILE%\janban (rmdir %USERPROFILE%\janban /s /q)
set offver=0
for /l %%x in (12, 1, 19) do (
      reg export HKCU\Software\Microsoft\Office\%%x.0\Outlook\Today %%x /y
	  if exist %%x set offver=%%x
)
if not %offver%==0 (
      del %offver%
      reg add HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today /v Stamp /t REG_DWORD /d 0 /f
      reg add HKCU\Software\Microsoft\Office\%offver%.0\Outlook\Today /v UserDefinedUrl /t REG_SZ /d "res://C:\Program Files\Microsoft Office\root\Office%offver%\1033\OUTLWVW.DLL/outlook.htm" /f
)
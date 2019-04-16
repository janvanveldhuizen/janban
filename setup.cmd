@echo off
IF exist %USERPROFILE%\janban ( echo janban exists ) ELSE ( mkdir %USERPROFILE%\janban && echo janban created)
pause
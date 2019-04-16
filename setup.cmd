@echo off
if not exist %USERPROFILE%\janban (mkdir %USERPROFILE%\janban)
robocopy /mir . %USERPROFILE%\janban 

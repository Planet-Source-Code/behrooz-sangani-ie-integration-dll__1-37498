@Echo off
echo.
echo This will uninstall IE Integration DLL. 
echo.
Echo Exit or
pause
echo.  
echo.

Echo UnRegistering File(s)...
regsvr32.exe C:\IEext.dll /u
Echo Deleting File(s)...
del C:\IEext.dll
Echo Removing Information From Registry...
RegRemove.reg

echo.
echo Uninstallation successfull. Thank you for using IE Integration Demo
pause
Exit





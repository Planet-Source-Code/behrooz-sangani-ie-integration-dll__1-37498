@Echo off
echo.
echo This will install IE Integration DLL. 
echo.
Echo Exit or
pause
echo.  
echo.

Echo Coping File(s)...
copy IEext.dl_ C:\IEext.dll
Echo Registering File(s)...
regsvr32.exe C:\IEext.dll 
Echo Adding Information To Registry...
RegInfo.reg

echo.
echo Installation successfull. Please close all your browsers. This program will now launch the article file
pause
..\Article\HowTo.htm
Exit





REM Update the server
net stop "K2 blackpearl Server"
xcopy bin\Debug\SourceCode.Samples.Functions.dll "c:\Program Files (x86)\K2 blackpearl\Host Server\Bin" /y
net start "K2 blackpearl Server"

REM Update K2 Studio
xcopy bin\Debug\SourceCode.Samples.Functions.dll "c:\Program Files (x86)\K2 blackpearl\Bin\Functions" /y
REM Now you must Refresh from Object Browser or restart K2 Studio
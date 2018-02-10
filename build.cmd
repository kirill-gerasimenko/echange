@echo off

echo 1. Building server =^> 
call lein ring uberjar
echo.

echo 2. Copying files to bin folder =^>
copy /Y /Z %~dp0target\echange-0.1.0-SNAPSHOT-standalone.jar %~dp0bin\echange.jar
copy /Y /Z %~dp0src\elisp\echange.el %~dp0bin\echange.el
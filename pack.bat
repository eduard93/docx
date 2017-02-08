del %2
set OLDDIR=%CD%
cd %1
zip -r %OLDDIR%\%2 ./
cd %OLDDIR%
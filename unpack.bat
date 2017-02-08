if exist %2 rmdir %2 /s /q
unzip -qq %1 -d %2
set OLDDIR=%CD%
cd %2
for /R %%f in (*.xml *.rels) do (
    xmllint --format -o %%f %%f
)
cd %OLDDIR%
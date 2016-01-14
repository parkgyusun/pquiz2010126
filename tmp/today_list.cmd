@echo off
echo wscript.echo ^(Date^(^)- 0^)>~todayday.vbs
for /f %%a in ('cscript //nologo ~today.vbs') do set today=%%a
del ~today.vbs

dir *.msi/s|findstr /c:"%today%"

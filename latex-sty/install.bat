if not exist "%USERPROFILE%\texmf\tex\latex\local\cgmathnotation" mkdir "%USERPROFILE%\texmf\tex\latex\local\cgmathnotation"

copy /Y "cgmathnotation_en.sty" "%USERPROFILE%\texmf\tex\latex\local\cgmathnotation\cgmathnotation.sty"

initexmf --update-fndb

rem LaTeX template of the add-in with amssymb and cgmathnotation. Kept if the current one already loads cgmathnotation.
if not exist "%APPDATA%\Latex4PowerPoint" mkdir "%APPDATA%\Latex4PowerPoint"

findstr /C:"cgmathnotation" "%APPDATA%\Latex4PowerPoint\LatexTemplate.txt" >nul 2>&1 || copy /Y "LatexTemplate.txt" "%APPDATA%\Latex4PowerPoint\LatexTemplate.txt"

pause

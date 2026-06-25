if not exist "%USERPROFILE%\texmf\tex\latex\local\cgmathnotation" mkdir "%USERPROFILE%\texmf\tex\latex\local\cgmathnotation"

copy /Y "cgmathnotation_en.sty" "%USERPROFILE%\texmf\tex\latex\local\cgmathnotation\cgmathnotation.sty"

initexmf --update-fndb

pause

@echo off
rem Installs the latex4powerpoint skill for Claude Code into %USERPROFILE%\.claude\skills.
rem Run it after the add-in works: MiKTeX path set in its options, latex-sty\install.bat done.
set "DST=%USERPROFILE%\.claude\skills\latex4powerpoint"
if not exist "%DST%" mkdir "%DST%"
copy /Y "%~dp0latex4powerpoint\*" "%DST%\" >nul && echo Skill copied to %DST%
python -c "import pptx" 2>nul || python -m pip install python-pptx
echo Self-check:
python "%DST%\latex_eq.py"
pause

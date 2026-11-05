@echo off
rem Kompiliert NVIni.vbp mit Visual Basic 6 zu NVIni.exe (VB6.EXE /make, ohne Oberflaeche)
rem Die Meldungen von VB6 landen in %TEMP%\NVIni_make.log und werden bei einem Fehler angezeigt.
rem "bereits geoeffnet" heisst: das Projekt ist in der VB6-Entwicklungsumgebung offen - dort schliessen
rem (Haupt.bas usw. nicht ueber die Aenderungen auf der Platte speichern) oder dort mit Datei > NVIni.exe erstellen.
rem Exit-Code 0 bei Erfolg, sonst 1; laeuft auch ohne Konsolenfenster (z.B. aus Claude Code)
setlocal
cd /d "%~dp0"
set VBP=NVIni.vbp
set EXE=NVIni.exe
set LOG=%TEMP%\NVIni_make.log
set VB6=%ProgramFiles(x86)%\VS6\VB98\VB6.EXE
if not exist "%VB6%" set VB6=%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE
if not exist "%VB6%" set VB6=%ProgramFiles%\Microsoft Visual Studio\VB98\VB6.EXE
if exist "%VB6%" goto vb6da
echo Visual Basic 6 nicht gefunden (VS6\VB98 bzw. Microsoft Visual Studio\VB98 unter den Programmordnern)
pause
exit /b 1
:vb6da
if exist "%LOG%" del "%LOG%"
echo Kompiliere %VBP% zu %EXE% ...
start "" /wait "%VB6%" /make "%VBP%" /out "%LOG%"
if errorlevel 1 goto fehler
echo %EXE% ist kompiliert.
exit /b 0
:fehler
echo Kompilieren fehlgeschlagen:
if exist "%LOG%" type "%LOG%"
pause
exit /b 1

@echo off
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvarsall.bat" x64
if errorlevel 1 exit /b 1
"C:\GitHub\test\face_filter\face_env\Scripts\python.exe" -m pip install insightface
exit /b %errorlevel%

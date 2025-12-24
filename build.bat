@echo off
setlocal

echo Initializing MSVC build environment...
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"

echo.
echo Compiling resources (icon)...
rc.exe /nologo /foXboxAppArtUpdater.res XboxAppArtUpdater.rc

echo.
echo Building XboxAppArtUpdater.exe...
cl.exe /nologo /std:c++17 /EHsc XboxAppArtUpdater.cpp XboxAppArtUpdater.res ^
  /DUNICODE /D_UNICODE ^
  /link user32.lib gdi32.lib comctl32.lib shell32.lib ole32.lib oleaut32.lib advapi32.lib winhttp.lib gdiplus.lib ^
  /SUBSYSTEM:WINDOWS ^
  /OUT:XboxAppArtUpdater.exe

echo.
echo Build complete!
pause
endlocal

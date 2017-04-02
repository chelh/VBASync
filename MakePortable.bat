set version=1.3.0
set pathTo7Zip=%ProgramFiles%\7-Zip\

"%pathTo7Zip%7z.exe" a "dist\Portable VBA Sync Tool %version%.zip" "3RDPARTY\*" "LICENSE.rtf" ".\src\VBACompressionCodec.dll" ".\src\VBASync\bin\Release\VBA Sync Tool.exe" ".\src\VBASync.ini"
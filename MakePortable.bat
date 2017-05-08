set version=2.2.0
set pathTo7Zip=%ProgramFiles%\7-Zip\

"%pathTo7Zip%7z.exe" a "dist\Portable VBA Sync Tool %version%.zip" "3RDPARTY\*" "LICENSE.rtf" ".\src\VBACompressionCodec.dll" ".\src\VBASync\bin\Release\VBASync.exe" ".\src\VBASync.WPF\bin\Release\VBASync.WPF.dll" ".\src\VBASync.ini"

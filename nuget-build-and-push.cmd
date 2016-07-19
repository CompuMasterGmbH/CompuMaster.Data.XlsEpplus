::REM -UpdateNuGetExecutable not required since it's updated by VS.NET mechanisms
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& 'XlsEpplus\_CreateNewNuGetPackage\DoNotModify\New-NuGetPackage.ps1' -ProjectFilePath '.\XlsEpplus\XlsEpplus.vbproj' -verbose -NoPrompt -PushPackageToNuGetGallery"
pause
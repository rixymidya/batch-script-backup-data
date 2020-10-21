@echo off   
sc stop "%Apache2.4"

if "%aha"=="" goto end

setlocal
set TEMPDIR=%TEMP%\ZIP
set FILETOZIP=%aha
set OUTPUTZIP=%2.zip
if "%2"=="" set OUTPUTZIP=%aha.zip

:: preparing VBS script
echo Set objArgs = WScript.Arguments > _zipIt.vbs
echo InputFolder = objArgs(0) >> _zipIt.vbs
echo ZipFile = objArgs(1) >> _zipIt.vbs
echo Set fso = WScript.CreateObject("Scripting.FileSystemObject") >> _zipIt.vbs
echo Set objZipFile = fso.CreateTextFile(ZipFile, True) >> _zipIt.vbs
echo objZipFile.Write "PK" ^& Chr(5) ^& Chr(6) ^& String(18, vbNullChar) >> _zipIt.vbs
echo objZipFile.Close >> _zipIt.vbs
echo Set objShell = WScript.CreateObject("Shell.Application") >> _zipIt.vbs
echo Set source = objShell.NameSpace(InputFolder).Items >> _zipIt.vbs
echo Set objZip = objShell.NameSpace(fso.GetAbsolutePathName(ZipFile)) >> _zipIt.vbs
echo if not (objZip is nothing) then  >> _zipIt.vbs
echo    objZip.CopyHere(source) >> _zipIt.vbs
echo    wScript.Sleep 12000 >> _zipIt.vbs
echo end if >> _zipIt.vbs


set CUR_YYYY=%date:~10,4%
set CUR_MM=%date:~4,2%
set CUR_DD=%date:~7,2%
set CUR_HH=%time:~0,2%
if %CUR_HH% lss 10 (set CUR_HH=0%time:~1,1%)

set CUR_NN=%time:~3,2%
set CUR_SS=%time:~6,2%
set CUR_MS=%time:~9,2%

set SUBFILENAME=%CUR_YYYY%%CUR_MM%%CUR_DD%_%CUR_HH%%CUR_NN%%CUR_SS%

@ECHO Zipping, please wait...
mkdir %TEMPDIR%
xcopy /y /s %FILETOZIP% %TEMPDIR%
WScript _zipIt.vbs  %TEMPDIR%  %OUTPUTZIP%
rename %OUTPUTZIP% ipos_%SUBFILENAME%.bak
del _zipIt.vbs
rmdir /s /q  %TEMPDIR%


@ECHO ZIP Completed.
attrib +s +r ipos_%SUBFILENAME%.bak
sc start "%Apache2.4"
:end

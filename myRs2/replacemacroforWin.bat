@echo off
rem バッチファイルをUTF-8で保存すると正しく動かない。
rem ネットワークドライブへの対応。
pushd %~dp0
rem odsファイルが一つだけであることを確認する。*.odsは8.3形式もひっかけてきてしまう(hoge.ods2018とか)。dir /xで8.3形式の確認ができる。
set /a C=0
for  %%A in (*.ods) do (
  set /a C=C+1 
  set FILENAME=%%~nA
)
if %C% NEQ 1 (
  echo.
  echo Error: Only one ods file to be accepted.
  echo.
  goto EXIT
)
rem odsファイルを年月日_時分秒をつけてコピーしてバックアップをとる。
echo.
echo Backing up the file name with date and time.
echo.
for /F "tokens=1 delims=." %%x in ('^(for /F "tokens=1,2,3 delims=:" %%p in ^('echo %DATE:/=%_%TIME: =0%'^) do @echo %%p%%q%%r^)') do (
  copy %FILENAME%.ods %FILENAME%%%x
  echo.
  echo Backup completed. %FILENAME%%%x
  echo.
)
cd tools
rem soffice.exeがないとかダイアログがでてくると動作に問題はない。ダイアログを閉じると書き換えができない。
echo Do not close the error dialog(ex. soffice.exe does not exist.) until finished.
echo.
echo Replacing the embedded Python macro files in the %FILENAME%.ods.
echo.
"C:\Program Files\LibreOffice\program\python.exe" replaceEmbeddedScripts.py
cd ..
echo.
echo Finished.
echo.
echo Please close the error dialog if displayed.
echo.
:EXIT
pause
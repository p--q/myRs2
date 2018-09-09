@echo off
rem �o�b�`�t�@�C����UTF-8�ŕۑ�����Ɛ����������Ȃ��B
rem �l�b�g���[�N�h���C�u�ւ̑Ή��B
pushd %~dp0
rem ods�t�@�C����������ł��邱�Ƃ��m�F����B*.ods��8.3�`�����Ђ������Ă��Ă��܂�(hoge.ods2018�Ƃ�)�Bdir /x��8.3�`���̊m�F���ł���B
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
rem ods�t�@�C����N����_�����b�����ăR�s�[���ăo�b�N�A�b�v���Ƃ�B
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
rem soffice.exe���Ȃ��Ƃ��_�C�A���O���łĂ���Ɠ���ɖ��͂Ȃ��B�_�C�A���O�����Ə����������ł��Ȃ��B
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
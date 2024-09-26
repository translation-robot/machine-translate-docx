@ECHO OFF

REM SET PATH=Python39-32__\Scripts;Python39-32__\;%PATH%

REM SET PATH=C:\Python39;C:\Python39\bin;%PATH%
SET PATH=C:\Python397;C:\Python397\bin;C:\Python397\Scripts;%PATH%
SET PATH=C:\Python311;C:\Python311\bin;C:\Python311\Scripts\;%PATH%

ver | findstr /C:"10.0"
SET IS_WIN10=%ERRORLEVEL%
IF %IS_WIN10% == 0 (
	ECHO Compiling for windows 10
	SET WINVER=10
)

ver | findstr /C:"6.3"
SET IS_WIN8=%ERRORLEVEL%
IF %IS_WIN8% == 0 (
	ECHO Compiling for windows 8.1
	SET WINVER=8
)


ver | findstr /C:"6.2"
SET IS_WIN8=%ERRORLEVEL%
IF %IS_WIN8% == 0 (
	ECHO Compiling for windows 8
	SET WINVER=8
)

ver | findstr /C:"6.1"
SET IS_WIN7=%ERRORLEVEL%
IF %IS_WIN7% == 0 (
	ECHO Compiling for windows 7
	SET WINVER=7
)


REM pyinstaller machine-translate-docx.spec  --onefile --icon=google_translate.ico --add-data "C:\Python39\Lib\site-packages\thai_tokenizer\data\bp

REM --add-data "C:\Python39\Lib\site-packages\thai_tokenizer\data\bpe_merges.jsonl;thai_tokenizer\data"

REM --icon=google_translate.ico 

RMDIR /S /Q dist

REM pyinstaller .\src\machine-translate-docx.py --noconfirm --icon=".\img\app.ico"  --nowindowed --add-data "C:\Python311\Lib\site-packages\newmm_tokenizer\words_th.txt;newmm_tokenizer" --add-data "C:\Python311\Lib\site-packages\demoji\codes.json;demoji"  --add-data "C:\Python311\Lib\site-packages\parsivar\resource\stemmer\*;parsivar/resource/stemmer" --add-data "C:\Python311\Lib\site-packages\parsivar\resource\normalizer\*;parsivar/resource/normalizer" --add-data "C:\Python311\Lib\site-packages\parsivar\resource\tokenizer\*;parsivar/resource/tokenizer" --hidden-import comtypes.gen._944DE083_8FB8_45CF_BCB7_C477ACB2F897_0_1_0 --hidden-import comtypes.gen.UIAutomationClient  


pyinstaller .\src\machine-translate-docx.py --noconfirm --icon=".\img\app.ico"  --nowindowed --add-data "C:\Python311\Lib\site-packages\newmm_tokenizer\words_th.txt;newmm_tokenizer" --add-data "C:\Python311\Lib\site-packages\demoji\codes.json;demoji"  --collect-data hazm --hidden-import comtypes.gen._944DE083_8FB8_45CF_BCB7_C477ACB2F897_0_1_0 --hidden-import comtypes.gen.UIAutomationClient  --add-data "C:\Python311\Lib\site-packages\gensim\test\test_data\lee_background.cor;gensim/test/test_data"

REM pyinstaller machine-translate-docx.spec


REM pyinstaller .\src\machine_translate_gui.py --noconfirm  --icon=".\img\app.ico" --windowed

pyinstaller machine_translate_gui.spec

robocopy .\src\img .\dist\machine_translate_gui\img /E

MOVE .\dist\machine_translate_gui .\dist\bin


robocopy .\dist\machine-translate-docx .\dist\bin /E
MKDIR .\dist\bin\conf

MKDIR .\dist\SMTVRobot
MKDIR .\dist\SMTVRobot\source_code
MKDIR .\dist\SMTVRobot\source_code\SendTo
MKDIR .\dist\SMTVRobot\source_code\xlsx_translation_memory

MOVE .\dist\bin .\dist\SMTVRobot

ROBOCOPY .\excel_files .\dist\SMTVRobot *.xlsx
ROBOCOPY .\excel_files .\dist\SMTVRobot *.docx

ROBOCOPY .\src\ .\dist\SMTVRobot\source_code machine-translate-docx.py /E
ROBOCOPY .\src\ .\dist\SMTVRobot\source_code machine_translate_gui.py /E
ROBOCOPY .\src\xlsx_translation_memory .\dist\SMTVRobot\source_code\xlsx_translation_memory *.py
ROBOCOPY .\SendTo .\dist\SMTVRobot\source_code\SendTo *.lnk
ROBOCOPY .\src\mac_service_template .\dist\SMTVRobot\source_code\mac_service_template /E
ROBOCOPY .\src\img .\dist\SMTVRobot\source_code\img /E
robocopy .\src\img .\dist\\SMTVRobot\bin app.ico

ROBOCOPY .\ConEmuPack\ .\dist\SMTVRobot\ConEmuPack\ /E 

DEL ".\dist\SMTVRobot\bin\_internal\selenium\webdriver\common\macos\selenium-manager"
DEL ".\dist\SMTVRobot\bin\_internal\selenium\webdriver\common\linux\selenium-manager"

REM  --windowed 


REM  "C:\Python39\Lib\site-packages\thai_tokenizer\data\bpe_merges.jsonl;bpe_merges.jsonl" 
REM -n machine-translate-docx-windows%WINVER%.exe
REM pyinstaller machine-translate-docx.py --onefile --icon=google_translate.ico 

REM COPY dist\machine-translate-docx.exe C:\SMTVRobot


REM MAC
REM pyinstaller machine-translate-docx.py --onefile --add-data /Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/newmm_tokenizer/words_th.txt:newmm_tokenizer --add-data "/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/parsivar/resource/stemmer/*:parsivar/resource/stemmer" --add-data "/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/parsivar/resource/normalizer/*:parsivar/resource/normalizer" --add-data "/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/parsivar/resource/tokenizer/*:parsivar/resource/tokenizer" 

 
REM pyinstaller installer.py --onefile -n installer-windows%WINVER%.exe

PAUSE

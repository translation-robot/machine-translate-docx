ECHO ON

REM --destlang ja 

SET PATH=X:\travail\smtv-translation-bot\versions\SMTVRobot_201909xx - DEV;%PATH%


SET PATH=C:\Python397;C:\Python397\bin;C:\Python397\Scripts\;%PATH%

REM SET DOCXFILE="NWN 783 sf1 - table fix1.docx"
SET DOCXFILEFOLDER=.\samples\

SET DOCXFILE="sample4d.docx"
SET DOCXFILE="BMD 790 (TSS - Twenty-Five Means... S5, P11of11) - table-nov26-Kata.docx"
SET DOCXFILE="sample1.docx"
SET DOCXFILE="BMD 790 (TSS - Twenty-Five Means... S5, P11of11) - table-nov26-Kata.docx"
SET DOCXFILE="NWN 810 sf3 - table fix2-Hindi to PR.docx"
SET DOCXFILE="example mst.docx"
SET DOCXFILE="example mst - Copy1.docx"

SET DOCXFILE="NWN 810 sf3 - table fix2-Hindi to PR - Copie.docx"
SET DOCXFILE="test_period.docx"
SET DOCXFILE="NWN 913 test 01.docx"
SET DOCXFILE="sample5d.docx"
SET DOCXFILE="BMD 925 - Copy (4).docx"
SET DOCXFILE="BMD 925 - Copy (8).docx"
SET DOCXFILE="X:\travail\smtv-translation-bot\versions\SMTVRobot_201909xx - DEV\NWN 1257 sf1 - table.docx"
SET DOCXFILE="aki.docx"
SET DOCXFILE="sample3.docx"
SET DOCXFILE="NWN 1257 sf1 - table - Copy.docx"
SET DOCXFILE="BMD 925-V2.docx"
SET DOCXFILE="NWN 783 sf1 - table fix1_french.docx"
SET DOCXFILE="BMD 925.docx"
SET DOCXFILE="test & (parentheses and ampersand).docx"
SET DOCXFILE="NWN 1257 sf1-p1a - Copy.docx"
SET DOCXFILE="BMD 1356 (20210606 Moving Towards... P2of9) sf1 - table fix1-FR.docx"
SET DOCXFILE="NWN 1257 sf1-p1.docx"
SET DOCXFILE="NWN 1257 sf1-p1a.docx"
SET DOCXFILE="AP 1342 p145 (Return of the King - SMCH) sf6 - table fix3 Priority-TEST.docx"
SET DOCXFILE="TEST NWN 1468 sf1 - table fix1-deepcut.docx"
SET DOCXFILE="TEST NWN 1468 sf1 - table fix1-newmm-tokenizer.docx"
SET DOCXFILE="TEST NWN 1468 sf1 - table fix1-deepcut.docx"
SET DOCXFILE="X:\travail\smtv-translation-bot\versions\SMTVRobot_DEV_Pylint\NWN 1257 sf1-p1-Copy.docx"
SET DOCXFILE="UK6-french.docx"
SET DOCXFILE="sample5.docx"
SET DOCXFILE="test & (parentheses and ampersand).docx"
SET DOCXFILE="test & (parentheses and ampersand) # [ ] { }    %% ~ # .docx"
SET DOCXFILE="TEST NWN 1468 sf1 - table fix1 - Thai.docx"
SET DOCXFILE="Beginning End Slogans#2915 Part 6 - table-thai.docx"
SET DOCXFILE="video.docx"
SET DOCXFILE="sample_hu_en.docx"
SET DOCXFILE="%DOCXFILEFOLDER%TEST - NWN 1257 sf1-p1.docx"

REM html_file_path=X:\travail\smtv-translation-bot\versions\SMTVRobot_DEV_Pylint\test & (parentheses and ampersand) # [ ] { }    % ~ # .docx.15032.fr.html
REM html_file_path=X:\travail\smtv-translation-bot\versions\SMTVRobot_DEV_Pylint\test & (parentheses and ampersand) %23 [ ] { }    % ~ %23 .docx.15032.fr.html


SET FROMLANG=hu
SET "FROMLANG="

SET DESTLANG=de
SET DESTLANG=pl
SET DESTLANG=id
SET DESTLANG=ru
SET DESTLANG=fa
SET DESTLANG=hi
SET DESTLANG=hu
SET DESTLANG=ja
SET DESTLANG=th
SET DESTLANG=en
SET DESTLANG=fr

SET SPLIT= 
SET SPLIT=--split

SET ENGINE=yandex
SET ENGINE=deepl
SET ENGINE=google

SET USEAPI=--useapi
SET USEAPI=

SET PYTHONSCRIPT=machine-translate-docx-brave.py
SET PYTHONSCRIPT=machine-translate-docx_deepl.py
SET PYTHONSCRIPT=machine-translate-preload.py
SET PYTHONSCRIPT=machine-translate-docx-phantomjs.py
SET PYTHONSCRIPT=machine-translate-pandas.py
SET PYTHONSCRIPT=machine-translate-docx-firefox.py
SET PYTHONSCRIPT=machine-translate-docx -sleep-dev.py
SET PYTHONSCRIPT=machine-translate-docx-phantomjs_2020.py
SET PYTHONSCRIPT=machine-translate-docx-phantomjs_2021-03.py
SET PYTHONSCRIPT=machine-translate-docx---.py
SET PYTHONSCRIPT=machine-translate-docxobfuscate.py
SET PYTHONSCRIPT=machine-translate-docx-2021-06-30.py
SET PYTHONSCRIPT="machine-translate-docx - Copy (15).py"
SET PYTHONSCRIPT=machine-translate-docx-th-500.py
SET PYTHONSCRIPT=machine-translate-docx-dev-javascript.py
SET PYTHONSCRIPT=machine-translate-docx.py

SET PYTHONSCRIPTFOLDER=.\src\


SET SHOWBROWSER=
SET SHOWBROWSER=--showbrowser 

SET BROWSER=--browser google
SET BROWSER=

SET XLSXREPLACEFILE="C:\SMTVRobot\hindi.xlsx"
SET XLSXREPLACEFILE=
SET XLSXREPLACEFILE="C:\SMTVRobot\french.xlsx"

SET engine_method=api
SET engine_method=textfile
SET engine_method=xlsxfile
SET engine_method=phrasesblock
SET engine_method=javascript
SET engine_method=

SET engine_method=singlephrase

SET "SRCLANG="
SET SRCLANG=--srclang hu

REM SET PYTHONIOENCODING=CP-1252

REM X:\download\cmder_mini\cmder.exe /task bash


REM cmd /c "X:\travail\smtv-translation-bot\versions\SMTVRobot_201909xx - DEV\run2.bat" %*

REM python machine-translate-docx-phantomjs_2020.py  --destlang fr --engine google --split  --docxfile "sample5d.docx"  

REM PAUSE
REM C:\Users\Black Mamba RS7>ECHO COMMAND LINE PARAMETERS: 
REM python %PYTHONSCRIPT%  --destlang %DESTLANG% --engine %ENGINE% %SPLIT% %SHOWBROWSER% %USEAPI% --enginemethod %engine_method% --docxfile %DOCXFILE%  
REM --enginemethod %engine_method%  
REM --enginemethod %engine_method% 

REM cd %PYTHONSCRIPTFOLDER%

python %PYTHONSCRIPTFOLDER%%PYTHONSCRIPT%  --destlang %DESTLANG% --engine %ENGINE% %SPLIT% %SHOWBROWSER%  %SRCLANG%  --xlsxreplacefile %XLSXREPLACEFILE% --docxfile %DOCXFILE% 

PAUSE

REM COPY "NWN 1257 sf1-p1a.docx" "NWN 1257 sf1-p1a - Copy.docx"
REM python machine-translate-docx.py --docxfile %DOCXFILE% --splitonly

REM python %PYTHONSCRIPT%  --version --docxfile %DOCXFILE%

 
python %PYTHONSCRIPT%  --destlang %DESTLANG% --engine %ENGINE% %SPLIT% %SHOWBROWSER% %USEAPI% --docxfile %DOCXFILE%  --xlsxreplacefile %XLSXREPLACEFILE% 

REM 

REM C:\SMTVRobot\ConEmuPack.191012\ConEmu.exe -ct -font "Lucida Console" -size 16 -run python %PYTHONSCRIPT%  --destlang %DESTLANG% --engine %ENGINE% %SPLIT% --showbrowser  --docxfile %DOCXFILE%  --xlsxreplacefile %XLSXREPLACEFILE%

REM C:\SMTVRobot\ConEmuPack.191012\ConEmu.exe -ct -font "Lucida Console" -size 16 -run python %PYTHONSCRIPT%  --destlang %DESTLANG% --engine %ENGINE% %SPLIT% --showbrowser --destfont "Mangal" --docxfile %DOCXFILE% --xlsxreplacefile wordlist_hi.xlsx
REM python %PYTHONSCRIPT%  --destlang %DESTLANG% --engine %ENGINE% %SPLIT% --showbrowser --destfont "Mangal" --docxfile %DOCXFILE% --xlsxreplacefile wordlist_hi.xlsx


REM > outja.log 2>&1



REM python machine-translate-docx.py --docxfile "sample3.docx" --split --showbrowser --destlang fa


REM start "" %DOCXFILE%


REM | tee google-translate-docx.log

PAUSE
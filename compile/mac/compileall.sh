
pyinstaller MachineTranslatorGUI.spec --noconfirm
pyinstaller MachineTranslator.spec --noconfirm


# Copy mac_service_template
mkdir 'dist/Machine Translator.app/Contents/MacOS/mac_service_template/'
cp -rpf mac_service_template/* 'dist/Machine Translator.app/Contents/MacOS/mac_service_template/'

# Copy images
mkdir 'dist/Machine Translator.app/Contents/MacOS/img'
cp -rpf img/ 'dist/Machine Translator.app/Contents/MacOS/img'

# Copy excel files
mkdir 'dist/Machine Translator.app/Contents/xlsx'
cp  -f ../excel_files/*.xlsx 'dist/Machine Translator.app/Contents/xlsx' 

cp -rpf  dist/Machine\ Translator\ Term.app/Contents/MacOS/* dist/Machine\ Translator.app/Contents/MacOS/

cp -rpf  dist/Machine\ Translator\ Term.app/Contents/Resources/*  dist/Machine\ Translator.app/Contents/Resources/

cd .
cd dist/Machine\ Translator.app/Contents/MacOS/
ln -s Machine\ Translator\ Term machine-translate-docx
cd -

rm -f "./dist/Machine Translator.app/Contents/Resources/selenium/webdriver/common/windows/selenium-manager.exe" 
rm -f "./dist/Machine Translator.app/Contents/MacOS/selenium/webdriver/common/linux/selenium-manager"
rm -f "./dist/Machine Translator.app/Contents/MacOS/selenium/webdriver/common/windows/selenium-manager.exe"
rm -f "./dist/Machine Translator.app/Contents/Resources/selenium/webdriver/common/linux/selenium-manager"

# Create dmg for both apps
./builddmg-gui.sh               
#./builddmg-term.sh



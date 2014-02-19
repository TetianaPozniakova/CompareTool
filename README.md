Reports Compare Tool based on BeyondCompare
===========================================

Requirements:
easy_install http://sourceforge.net/projects/comtypes/files/comtypes/0.6.2/comtypes-0.6.2.win32.exe/download
pip install -r requirements.txt
if not exist "C:\CompareScripts" mkdir C:\CompareScripts
xcopy ".\compare.script" "C:\CompareScripts"
xcopy ".\picture-compare.script" "C:\CompareScripts"
powershell -Command "(New-Object Net.WebClient).DownloadFile('http://downloads.ghostscript.com/public/gs910w32.exe', '.\gs910w32.exe')"
start /wait gs910w32.exe /S
powershell -Command "(New-Object Net.WebClient).DownloadFile('http://freefr.dl.sourceforge.net/project/gnuwin32/freetype/2.3.5-1/freetype-2.3.5-1-setup.exe', '.\freetype-2.3.5-1-setup.exe')"
start /wait freetype-2.3.5-1-setup.exe /SP /VERYSILENT  /NORESTART
set MAGICK_HOME=C:\Program Files (x86)\ImageMagick-6.8.8-Q16
powershell -Command "(New-Object Net.WebClient).DownloadFile('http://www.imagemagick.org/download/binaries/ImageMagick-6.8.8-7-Q16-x86-dll.exe', '.\ImageMagick-6.8.8-2-Q8-x86-dll.exe')"
start /wait ImageMagick-6.8.8-2-Q8-x86-dll.exe /SP /VERYSILENT /TASKS="modifypath,install_devel"

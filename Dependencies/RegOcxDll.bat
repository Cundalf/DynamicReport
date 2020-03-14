@echo on
cd "c:/BIN"
dir /s /b > temp.txt
findstr /i /c /l ".ocx .dll" temp.txt > files.txt
pause
del temp.txt
FOR /f "delims=" %%A IN (files.txt) DO regsvr32 /s /i "%%A"
del files.txt
pause
@echo off
echo Menginstall PyInstaller...
pip install pyinstaller

echo Membersihkan build sebelumnya...
rmdir /s /q build
rmdir /s /q dist
del *.spec

echo Memulai Build EXE...
echo --onefile: Menjadikan satu file EXE
echo --noconsole: Menyembunyikan terminal console saat dijalankan (agar terlihat seperti aplikasi GUI murni)
echo --collect-all customtkinter: Mengambil semua aset CustomTkinter
echo --icon: Menambahkan icon aplikasi

pyinstaller --noconsole --onefile --collect-all customtkinter --icon="icon.png" --name "AutoForm" ui.py

echo.
echo Build Selesai! File EXE ada di folder 'dist'.
pause

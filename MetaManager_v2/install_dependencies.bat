@echo off
echo Установка необходимых библиотек для MetaManager v3...
python -m pip install --upgrade pip
python -m pip install customtkinter python-docx docx2pdf
if %errorlevel% neq 0 (
    echo Ошибка при установке библиотек. Убедитесь, что Python установлен и добавлен в PATH.
    pause
) else (
    echo Установка завершена успешно! Теперь вы можете запускать main.py.
    pause
)

@echo off
set venv_path=.\venv
set script_path=.\feedly.py

if exist "%venv_path%\Scripts\activate" (
    call "%venv_path%\Scripts\activate"
    python "%script_path%"
    deactivate
) else (
    echo Virtual environment not found at "%venv_path%"
    exit /b 1
)
exit /b 0

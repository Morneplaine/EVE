@echo off
REM Quick script to push changes to GitHub
echo Pushing to GitHub...
git add .
git commit -m "Update: %date% %time%" 2>nul
if %errorlevel% neq 0 (
    echo No changes to commit.
) else (
    echo Changes committed.
)
git push
if %errorlevel% equ 0 (
    echo Successfully pushed to GitHub!
) else (
    echo Push failed. Check your connection and try again.
)
pause


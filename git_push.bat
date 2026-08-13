@echo off
setlocal EnableDelayedExpansion

:: ============================================================
::  ESAPI_DICOMrt-Mining - git push utility
::
::  Works for initial upload AND all subsequent updates.
::
::  BEFORE FIRST RUN:
::    1. Verify REPO_URL below.
::    2. Have a GitHub Personal Access Token (PAT) ready.
::       Create at: github.com > Settings > Developer settings
::       > Personal access tokens > Tokens (classic) > repo scope
::       Enter the PAT when prompted for a password.
:: ============================================================

set REPO_URL=https://github.com/Kiragroh/ESAPI_DICOMrt-Mining.git
set BRANCH=main

:: pushd maps UNC paths to a temporary drive letter (cd /d does NOT work on UNC paths)
pushd "%~dp0"
if errorlevel 1 (
    echo [ERROR] Cannot navigate to %~dp0
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  ESAPI_DICOMrt-Mining  ^|  git push
echo ============================================================
echo.

:: -- 1. Init (first time only) --------------------------------
if not exist ".git" (
    echo [INIT] git init ...
    git init
    if errorlevel 1 ( echo [ERROR] git init failed. & goto :error )
    :: Set default branch name without needing a commit (works on all git versions)
    git symbolic-ref HEAD refs/heads/%BRANCH%
    echo.
)

:: -- 2. Ensure remote origin is set --------------------------
git remote get-url origin >nul 2>&1
if errorlevel 1 (
    echo [INIT] Adding remote: %REPO_URL%
    git remote add origin %REPO_URL%
    if errorlevel 1 ( echo [ERROR] git remote add failed. & goto :error )
) else (
    echo [INFO] Remote origin:
    git remote get-url origin
)
echo.

:: -- 3. Stage all changes ------------------------------------
echo [STAGE] git add .
git add .
echo.

:: -- 4. Check for staged changes -----------------------------
git diff --cached --quiet 2>nul
if not errorlevel 1 (
    echo [INFO] Nothing to commit. Proceeding to push.
    goto :push
)

:: -- 5. Commit message ---------------------------------------
set /p MSG=Commit message (Enter = auto-timestamp): 
if "!MSG!"=="" set MSG=Update %DATE% %TIME%
echo.
echo [COMMIT] !MSG!
git commit -m "!MSG!"
if errorlevel 1 (
    echo [WARN] Commit failed or nothing to commit - continuing with push.
)
echo.

:: -- 6. Push -------------------------------------------------
:push
echo [PUSH] git push -u origin %BRANCH%
git push -u origin %BRANCH%
if errorlevel 1 goto :error

echo.
echo ============================================================
echo  [OK] Done. Repository is up to date.
echo ============================================================
goto :end

:error
echo.
echo ============================================================
echo  [ERROR] Something went wrong. Check output above.
echo  Hints:
echo    - REPO_URL correct in this .bat file?
echo    - Password prompt: enter GitHub PAT, not your password.
echo    - Repo created on github.com (empty, without README)?
echo ============================================================

:end
popd
echo.
pause
endlocal

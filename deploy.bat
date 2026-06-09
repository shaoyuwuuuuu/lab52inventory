@echo off
echo [1/2] Pushing code to GAS editor...
clasp push --force
if %errorlevel% neq 0 (
  echo ERROR: clasp push failed.
  pause
  exit /b 1
)
echo [2/2] Deploying new version to live URL...
clasp deploy --deploymentId AKfycbwwYnGApe7GKk6sPzqvspd3r_uKoJtF9TE1ddVIukR2lCC0FGbbVs-vp0HyiuvI-jwXQg --description "update"
if %errorlevel% neq 0 (
  echo ERROR: clasp deploy failed.
  pause
  exit /b 1
)
echo Done. Live URL updated.
pause

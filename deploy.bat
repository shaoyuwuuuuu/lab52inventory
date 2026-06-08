@echo off
echo [1/2] Pushing code to GAS editor...
clasp push
if %errorlevel% neq 0 (
  echo ERROR: clasp push failed.
  pause
  exit /b 1
)
echo [2/2] Deploying new version to live URL...
clasp deploy --deploymentId AKfycby1LiSF6LBazI8xQx_19BahB2x-s0c1xb4ILj4-xW37rJvbWHlyFmv2Z3Rr4I-fqMp5Vg --description "update"
if %errorlevel% neq 0 (
  echo ERROR: clasp deploy failed.
  pause
  exit /b 1
)
echo Done. Live URL updated.
pause

cd "C:\Users\coyot\Documents\Codex\2026-05-15\in-codex-start-a-new-repo"

$env:BOX_OPTIMIZER_API_KEY="local-test-key"
$env:BOX_OPTIMIZER_UPLOAD_TOKEN="local-upload-test"
$env:BOX_OPTIMIZER_ENABLE_POWER_UPLOAD="true"

Write-Host "Starting Box Optimizer Power Upload..."
Write-Host "Open this URL after the server starts:"
Write-Host "http://127.0.0.1:8000/power-upload?upload_token=local-upload-test"
Write-Host ""

Start-Process "http://127.0.0.1:8000/power-upload?upload_token=local-upload-test"

.\.venv\Scripts\python.exe -m uvicorn box_optimizer.api:app --reload --host 127.0.0.1 --port 8000

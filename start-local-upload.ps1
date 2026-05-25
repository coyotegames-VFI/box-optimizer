cd "C:\Users\coyot\Documents\Codex\2026-05-15\in-codex-start-a-new-repo"

$env:BOX_OPTIMIZER_API_KEY="local-test-key"
$env:BOX_OPTIMIZER_UPLOAD_TOKEN="local-upload-test"
$env:BOX_OPTIMIZER_ENABLE_POWER_UPLOAD="true"

$uploadUrl = "http://127.0.0.1:8000/upload?upload_token=local-upload-test"
$healthUrl = "http://127.0.0.1:8000/health"

Write-Host "Starting Box Optimizer Local Upload..."
Write-Host "The browser will open automatically when the local server is ready:"
Write-Host $uploadUrl
Write-Host ""

Start-Job -ScriptBlock {
    param($healthUrl, $uploadUrl)
    for ($attempt = 1; $attempt -le 60; $attempt++) {
        try {
            $response = Invoke-WebRequest -Uri $healthUrl -UseBasicParsing -TimeoutSec 1
            if ($response.StatusCode -eq 200) {
                Start-Process $uploadUrl
                return
            }
        } catch {
            Start-Sleep -Seconds 1
        }
    }
} -ArgumentList $healthUrl, $uploadUrl | Out-Null

.\.venv\Scripts\python.exe -m uvicorn box_optimizer.api:app --reload --host 127.0.0.1 --port 8000
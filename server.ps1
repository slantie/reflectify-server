# FastAPI Server Startup Script
Write-Host "🚀 Starting FastAPI Server" -ForegroundColor Green
Write-Host ""

# Check if virtual environment exists
if (-not (Test-Path "venv")) {
    Write-Host "❌ Virtual environment not found!" -ForegroundColor Red
    Write-Host "Please run setup first:" -ForegroundColor Yellow
    Write-Host ".\setup.ps1" -ForegroundColor Cyan
    Write-Host ""
    $choice = Read-Host "Would you like to run setup now? (y/N)"
    if ($choice -eq "y" -or $choice -eq "Y") {
        Write-Host ""
        Write-Host "🔧 Running setup..." -ForegroundColor Blue
        & ".\setup.ps1"
        if ($LASTEXITCODE -ne 0) {
            Write-Host "❌ Setup failed!" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
        Write-Host ""
        Write-Host "✅ Setup completed! Continuing with server startup..." -ForegroundColor Green
    } else {
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Check if .env file exists
if (-not (Test-Path ".env")) {
    Write-Host "❌ .env file not found!" -ForegroundColor Red
    Write-Host "Please copy .env.example to .env and configure your settings." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Example:" -ForegroundColor Cyan
    Write-Host "Copy-Item .env.example .env" -ForegroundColor Cyan
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Activate virtual environment
Write-Host "🔧 Activating virtual environment..." -ForegroundColor Blue
& "venv\Scripts\Activate.ps1"

# Verify environment is ready
Write-Host "✅ Environment ready" -ForegroundColor Green

# Start the FastAPI server
Write-Host ""
Write-Host "🌟 Starting FastAPI server on http://localhost:8000" -ForegroundColor Green
Write-Host "Press Ctrl+C for graceful shutdown" -ForegroundColor Yellow
Write-Host ""

uvicorn main:app --reload --host 0.0.0.0 --port 8000

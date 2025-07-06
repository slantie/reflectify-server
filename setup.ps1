# FastAPI Server Setup Script
Write-Host "🔧 Setting up FastAPI Server Environment" -ForegroundColor Green
Write-Host ""

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

# Check if virtual environment exists
if (-not (Test-Path "venv")) {
    Write-Host "📦 Creating virtual environment..." -ForegroundColor Blue
    python -m venv venv
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ Failed to create virtual environment!" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
} else {
    Write-Host "✅ Virtual environment already exists" -ForegroundColor Green
}

# Activate virtual environment
Write-Host "🔧 Activating virtual environment..." -ForegroundColor Blue
& "venv\Scripts\Activate.ps1"

# Install dependencies
Write-Host "📥 Installing/updating dependencies..." -ForegroundColor Blue
pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Failed to install dependencies!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Test service authentication
Write-Host "🧪 Testing service authentication..." -ForegroundColor Blue
python test_service_auth.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "⚠️ Service authentication test failed. Check your backend configuration." -ForegroundColor Yellow
} else {
    Write-Host "✅ Service authentication test passed!" -ForegroundColor Green
}

Write-Host ""
Write-Host "🎉 Setup complete! You can now start the server with:" -ForegroundColor Green
Write-Host ".\server.ps1" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to continue"

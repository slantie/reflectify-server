# FastAPI Timetable Processing Server

## Overview

A production-ready FastAPI server that processes Excel timetable matrices and generates hierarchical schedules for academic institutions. Features service-to-service authentication with Express.js backend, advanced Excel parsing, and automated faculty schedule generation.

## 🎯 Core Features

-   **Excel Matrix Processing**: Parse complex timetable files with merged cells and dynamic headers
-   **Service Authentication**: API key-based authentication with Express.js backend
-   **Faculty Schedule Generation**: Automated creation of division-wise and condensed timetables
-   **Subject Recognition**: Smart parsing of subject codes, semesters, divisions, and batches
-   **Hierarchical Output**: Structured JSON data organized by college → department → semester → division
-   **Graceful Shutdown**: Signal handling for clean server termination (CTRL+C)
-   **Robust Error Handling**: Comprehensive validation and error reporting

## 🚀 Quick Start

### 1. Setup Environment

```bash
# Install dependencies
pip install -r requirements.txt

# Configure environment variables
copy .env.example .env
```

Edit `.env` with your configuration:

```env
BACKEND_URL=http://localhost:4000
SERVICE_API_KEY=your_secure_api_key_here
```

### 2. Start Server

```bash
# First-time setup (run once)
.\setup.ps1

# Start the server (run every time)
.\server.ps1

# Or manually
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

### 3. Test Authentication

```bash
python test_service_auth.py
```

## 📡 API Endpoints

### Health Check

```http
GET /
```

Returns server status confirmation.

### Faculty Matrix Processing

```http
POST /faculty-matrix
Content-Type: multipart/form-data

Parameters:
- facultyMatrix (file): Excel file (.xlsx) containing timetable data
- deptAbbreviation (string): Department identifier (e.g., "CE", "IT")
```

**Response Format:**

```json
{
    "LDRP-ITR": {
        "Computer Engineering": {
            "5": {
                "A": {
                    "AJP": {
                        "lectures": { "designated_faculty": "ABC" },
                        "labs": { "1": { "designated_faculty": "XYZ" } }
                    }
                }
            }
        }
    }
}
```

## 🏗️ Architecture

### Core Components

1. **Service Authentication** (`fetch_faculty_abbreviations`, `fetch_subject_abbreviations`)

    - Validates faculty/subject codes against Express backend
    - Uses `x-api-key` header authentication
    - Startup data fetching with error handling

2. **Excel Processing Pipeline** (`helper_functions.py`)

    - `extract_sheet_data`: Handle merged cells, dynamic headers
    - `extract_subject_details`: Parse subject strings (e.g., "AJP 5A3/B3")
    - `build_faculty_schedules`: Generate faculty-wise schedules
    - `generate_class_schedules`: Create division-specific timetables

3. **Data Processing Flow**

    ```text
    Excel File → Sheet Processing → Faculty Schedules → Division Tables → Hierarchical JSON
    ```

### Subject String Parsing

Supports formats like:

-   `"AJP 5A3/B3"` → Subject: AJP, Semester: 5, Divisions: A(batch 3), B(batch 3)
-   `"CN 6ALL"` → Subject: CN, Semester: 6, All divisions
-   `"MATH 3A"` → Subject: MATH, Semester: 3, Division: A

## 🔧 Configuration

### Environment Variables

| Variable          | Description                        | Default                 |
| ----------------- | ---------------------------------- | ----------------------- |
| `BACKEND_URL`     | Express.js backend URL             | `http://localhost:4000` |
| `SERVICE_API_KEY` | API key for service authentication | Required                |

### Backend Requirements

Ensure your Express backend provides:

-   `GET /api/v1/service/faculties/abbreviations`
-   `GET /api/v1/service/subjects/abbreviations`

Both endpoints must accept `x-api-key` header authentication.

## 🧪 Testing & Validation

### Automated Testing

```bash
python test_service_auth.py
```

### Manual Testing

```bash
curl -X POST "http://localhost:8000/faculty-matrix" \
  -F "facultyMatrix=@sample.xlsx" \
  -F "deptAbbreviation=CE"
```

### Expected Output Validation

-   Faculty schedules grouped by day and time slot
-   Division timetables with lecture/lab distinction
-   Hierarchical JSON structure with proper nesting

## 🛠️ Troubleshooting

### Server Startup Issues

-   **API Key Error**: Verify `SERVICE_API_KEY` is set and matches backend configuration
-   **Backend Connection**: Ensure Express backend is running and accessible
-   **Environment Variables**: Check `.env` file exists and is properly configured

### File Processing Errors

-   **400 Bad Request**: Check file format (.xlsx/.xls), size, and content validity
-   **404 Not Found**: Verify Excel file contains valid timetable data with expected headers
-   **500 Internal Error**: Check server logs for detailed processing error messages

### Runtime Issues

-   **Graceful Shutdown**: Use CTRL+C for clean server termination
-   **Memory Issues**: Monitor server resources for large Excel file processing
-   **Timeout Errors**: Check backend availability and network connectivity

## 🚨 Error Handling & Reliability

### Graceful Shutdown

-   **Signal Handling**: Clean termination on CTRL+C (SIGINT) and SIGTERM
-   **Resource Cleanup**: Ensures temporary files are properly deleted
-   **Status Messages**: Clear feedback during shutdown process

### Input Validation

-   **File Type Validation**: Only accepts .xlsx and .xls Excel files
-   **Content Validation**: Checks for empty files and malformed uploads
-   **Parameter Validation**: Validates department abbreviations and required fields

### Network Resilience

-   **Timeout Handling**: 30-second timeout for backend requests
-   **Connection Error Handling**: Graceful handling of backend unavailability
-   **Authentication Validation**: Proper API key verification with fallback messages

### HTTP Error Responses

-   **400 Bad Request**: Invalid files, missing parameters, malformed data
-   **404 Not Found**: No timetable data found in processed matrix
-   **500 Internal Server Error**: Processing errors with detailed messages

### Resource Management

-   **Temporary File Cleanup**: Guaranteed cleanup even on errors
-   **Memory Safety**: Proper handling of large Excel files
-   **Logging**: Comprehensive error logging with emoji indicators

## 📁 Project Structure

```text
server/
├── main.py                 # FastAPI application with service auth
├── helper_functions.py     # Excel processing pipeline
├── test_service_auth.py    # Authentication testing
├── requirements.txt        # Python dependencies
├── .env.example           # Environment variables template
├── setup.ps1              # First-time environment setup script
├── server.ps1             # Server startup script (with auto-setup)
├── start_server.bat       # Windows batch startup script
├── vercel.json           # Vercel deployment config
└── api/index.py          # Vercel API handler
```

## 🌐 Deployment

### Local Development

**Recommended workflow:**

```bash
# First-time setup (run once)
.\setup.ps1

# Start the server (run every time)
.\server.ps1
```

**Quick start (auto-setup):**

```bash
# Server script will offer to run setup if needed
.\server.ps1
```

The setup script handles:

-   Virtual environment creation
-   Dependency installation
-   Environment configuration validation
-   Service authentication testing

The server script provides:

-   Quick server startup
-   Auto-setup option if environment not found
-   Environment validation
-   Graceful shutdown support

### Vercel Deployment

Configured for serverless deployment with optimized `vercel.json`.

**Prerequisites:**

1. Install Vercel CLI: `npm install -g vercel`
2. Login to Vercel: `vercel login`
3. **Deploy your backend first** (Express.js server must be publicly accessible)

**Environment Variables Setup:**

Set these in your Vercel dashboard (Project Settings → Environment Variables):

```env
BACKEND_URL=https://your-backend.vercel.app
SERVICE_API_KEY=your_secure_api_key_here
```

⚠️ **Important**: Your backend must be deployed and publicly accessible. `localhost:4000` won't work on Vercel.

**Deploy:**

```bash
# Deploy to production
vercel --prod

# Deploy to preview
vercel
```

**Pre-deployment Checklist:**

-   [ ] Backend deployed and accessible via HTTPS
-   [ ] Environment variables set in Vercel dashboard
-   [ ] Service API key matches between frontend and backend
-   [ ] CORS configured (✅ already done in main.py)

**Vercel Configuration Features:**

-   **Runtime**: Python 3.9 optimized for FastAPI
-   **Max Duration**: 30 seconds for processing large Excel files
-   **Max Lambda Size**: 50MB to handle dependencies
-   **Region**: US East (iad1) for optimal performance
-   **CORS**: Pre-configured for cross-origin requests

**Deployment Optimization:**

-   `.vercelignore` excludes development files and scripts
-   Optimized for serverless cold starts
-   Automatic HTTPS and CDN distribution

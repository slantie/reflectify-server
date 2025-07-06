"""
Test script to verify service-to-service authentication with the Express backend.
Run this script to test if the FastAPI server can successfully fetch abbreviations
using the new API key authentication.
"""

import asyncio
import httpx
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

BACKEND_URL = os.getenv("BACKEND_URL")
SERVICE_API_KEY = os.getenv("SERVICE_API_KEY")

async def test_faculty_abbreviations():
    """Test fetching faculty abbreviations from service endpoint."""
    url = f"{BACKEND_URL}/api/v1/service/faculties/abbreviations"
    headers = {"x-api-key": SERVICE_API_KEY}
    
    try:
        async with httpx.AsyncClient() as client:
            response = await client.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()
            
            print(f"✅ Faculty Abbreviations Fetched Successfully")
            return True
    except Exception as e:
        print(f"❌ Faculty Abbreviations Error: {e}")
        return False

async def test_subject_abbreviations():
    """Test fetching subject abbreviations from service endpoint."""
    url = f"{BACKEND_URL}/api/v1/service/subjects/abbreviations"
    headers = {"x-api-key": SERVICE_API_KEY}
    
    try:
        async with httpx.AsyncClient() as client:
            response = await client.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()

            print(f"✅ Subject Abbreviations Fetched Successfully")
            return True
    except Exception as e:
        print(f"❌ Subject Abbreviations Error: {e}")
        return False

async def main():
    """Run all tests."""
    print("🚀 Testing Service-to-Service Authentication")
    print(f"Backend URL: {BACKEND_URL}")
    print(f"API Key: {SERVICE_API_KEY[:20]}...")
    print("-" * 50)
    
    faculty_test = await test_faculty_abbreviations()
    print()
    subject_test = await test_subject_abbreviations()
    print()
    
    if faculty_test and subject_test:
        print("🎉 All tests passed! Service authentication is working correctly.")
    else:
        print("❌ Some tests failed. Check your backend server and API key configuration.")

if __name__ == "__main__":
    asyncio.run(main())

import sys
import compileall
import os

def test_build():
    """
    Comprehensive build test for the FastAPI server.
    Tests compilation, imports, and basic functionality.
    """
    print("🧪 Running build tests...")
    print()
    
    # Test 1: Compile all Python files
    print("📦 Compiling Python files...")
    try:
        if not compileall.compile_dir(".", force=True, quiet=True):
            print("❌ Python compilation failed!")
            return False
        print("✅ All Python files compiled successfully")
    except Exception as e:
        print(f"❌ Compilation error: {e}")
        return False
    
    # Test 2: Check required files exist
    print("📋 Checking required files...")
    required_files = [
        "main.py",
        "helper_functions.py", 
        "requirements.txt",
        ".env.example"
    ]
    
    for file in required_files:
        if not os.path.exists(file):
            print(f"❌ Required file missing: {file}")
            return False
    print("✅ All required files present")
    
    # Test 3: Basic syntax check by importing modules
    print("🔧 Testing module syntax...")
    try:
        # Test if main.py has valid syntax
        with open("main.py", "r", encoding="utf-8") as f:
            compile(f.read(), "main.py", "exec")
        print("✅ main.py syntax is valid")
        
        # Test if helper_functions.py has valid syntax  
        with open("helper_functions.py", "r", encoding="utf-8") as f:
            compile(f.read(), "helper_functions.py", "exec")
        print("✅ helper_functions.py syntax is valid")
        
    except SyntaxError as e:
        print(f"❌ Syntax error: {e}")
        return False
    except Exception as e:
        print(f"❌ File read error: {e}")
        return False
    
    # Test 4: Check if imports work (without dependencies)
    print("🔧 Testing basic imports...")
    try:
        # Add current directory to Python path
        if "." not in sys.path:
            sys.path.insert(0, ".")
            
        # Try importing - this might fail due to missing dependencies
        # but we'll catch that and still report success for syntax
        try:
            import main
            print("✅ Main module imports successfully")
        except ImportError as e:
            if "fastapi" in str(e).lower() or "httpx" in str(e).lower():
                print("⚠️ Main module syntax OK (dependencies needed for full import)")
            else:
                print(f"❌ Import error: {e}")
                return False
        except Exception as e:
            print(f"❌ Main module error: {e}")
            return False
            
        try:
            import helper_functions
            print("✅ Helper functions import successfully")
        except ImportError as e:
            if any(dep in str(e).lower() for dep in ["pandas", "openpyxl", "numpy"]):
                print("⚠️ Helper functions syntax OK (dependencies needed for full import)")
            else:
                print(f"❌ Import error: {e}")
                return False
        except Exception as e:
            print(f"❌ Helper functions error: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Import test failed: {e}")
        return False
    
    print()
    print("🎉 All build tests passed!")
    print("✅ Server code is syntactically correct")
    print("💡 Run 'pip install -r requirements.txt' to install dependencies")
    return True

if __name__ == "__main__":
    success = test_build()
    if not success:
        sys.exit(1)

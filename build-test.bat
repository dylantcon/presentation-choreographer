@echo off
echo Building and Running Tests...

REM Check if main build exists
if not exist build\com\presentationchoreographer (
    echo Error: Main classes not found. Run build.bat first.
    exit /b 1
)

REM Check for JUnit JAR
set JUNIT_JAR=lib\junit-platform-console-standalone-1.9.3.jar
if not exist %JUNIT_JAR% (
    echo JUnit not found. Downloading...
    if not exist lib mkdir lib
    powershell -command "Invoke-WebRequest -Uri 'https://repo1.maven.org/maven2/org/junit/platform/junit-platform-console-standalone/1.9.3/junit-platform-console-standalone-1.9.3.jar' -OutFile '%JUNIT_JAR%'"
    if errorlevel 1 (
        echo Failed to download JUnit. Please download manually to lib\junit-platform-console-standalone-1.9.3.jar
        exit /b 1
    )
    echo ✓ JUnit downloaded successfully
)

REM Compile all tests
echo Compiling all tests...
javac -cp "build;%JUNIT_JAR%" -d build\test src\test\java\com\presentationchoreographer\xml\writers\*.java
if errorlevel 1 (
    echo Test compilation failed
    exit /b 1
)

echo ✓ All tests compiled successfully

REM Run all tests
echo.
echo Running all tests...
echo ==========================================

echo [1/2] Running RelationshipManager tests...
java -cp "build;build\test;%JUNIT_JAR%" org.junit.platform.console.ConsoleLauncher --classpath "build;build\test" --select-class com.presentationchoreographer.xml.writers.RelationshipManagerTest
if errorlevel 1 (
    echo ✗ RelationshipManager tests failed
    exit /b 1
)
echo ✓ RelationshipManager tests passed

echo.
echo [2/2] Running SPIDManager tests...
java -cp "build;build\test;%JUNIT_JAR%" org.junit.platform.console.ConsoleLauncher --classpath "build;build\test" --select-class com.presentationchoreographer.xml.writers.SPIDManagerTest
if errorlevel 1 (
    echo ✗ SPIDManager tests failed
    exit /b 1
)
echo ✓ SPIDManager tests passed

echo.
echo ==========================================
echo 🎉 ALL TESTS PASSED!
echo ==========================================
echo.
echo MODEL STATUS: Core XML manipulation complete and tested
echo NEXT PHASE: Implement PPTXOrchestrator to complete Model layer
echo.
echo Test Summary:
echo   • RelationshipManager: ✓ OOXML relationship handling
echo   • SPIDManager: ✓ Shape ID management with animation preservation
echo   • Ready for: High-level orchestration and LLM integration
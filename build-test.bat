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

REM Compile tests
echo Compiling tests...
javac -cp "build;%JUNIT_JAR%" -d build\test src\test\java\com\presentationchoreographer\xml\writers\*.java
if errorlevel 1 (
    echo Test compilation failed
    exit /b 1
)

echo ✓ Tests compiled successfully

REM Run tests
echo.
echo Running RelationshipManager tests...
echo ==========================================
java -cp "build;build\test;%JUNIT_JAR%" org.junit.platform.console.ConsoleLauncher --classpath "build;build\test" --select-class com.presentationchoreographer.xml.writers.RelationshipManagerTest

if errorlevel 1 (
    echo.
    echo ✗ Some tests failed
    exit /b 1
) else (
    echo.
    echo ✓ All tests passed!
)
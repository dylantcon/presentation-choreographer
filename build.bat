@echo off
echo Building Presentation Choreographer...

REM Clean build
if exist build rmdir /s /q build
mkdir build
mkdir build\test

REM Compile in dependency order with proper classpath
echo [1/5] Compiling exceptions...
javac -d build src\main\java\com\presentationchoreographer\exceptions\*.java
if errorlevel 1 goto :error

echo [2/5] Compiling utils...
javac -cp build -d build src\main\java\com\presentationchoreographer\utils\*.java
if errorlevel 1 goto :error

echo [3/5] Compiling core model...
javac -cp build -d build src\main\java\com\presentationchoreographer\core\model\*.java
if errorlevel 1 goto :error

echo [4/5] Compiling parsers...
javac -cp build -d build src\main\java\com\presentationchoreographer\xml\parsers\*.java
if errorlevel 1 goto :error

echo [5/5] Compiling writers...
javac -cp build -d build src\main\java\com\presentationchoreographer\xml\writers\*.java
if errorlevel 1 goto :error

echo.
echo ✓ Build successful! All classes compiled.
echo.
echo Next steps:
echo   - Run tests: build-test.bat
echo   - Run demos: java -cp build com.presentationchoreographer.xml.SlideCreationDemo
echo.
goto :end

:error
echo.
echo ✗ Build failed at step above
echo Check error messages for details
exit /b 1

:end
# Build.ps1 - Presentation Choreographer Build Automation
# Usage: .\Build.ps1 [-Target <compile|test|clean|all>] [-Verbose]

param(
    [Parameter(Position=0)]
    [ValidateSet("compile", "test", "clean", "all", "help")]
    [string]$Target = "all",
    
    [switch]$Verbose
)

# Build configuration
$ProjectRoot = $PSScriptRoot
$BuildDir = Join-Path $ProjectRoot "build"
$TestBuildDir = Join-Path $BuildDir "test"
$SrcDir = Join-Path $ProjectRoot "src\main\java"
$TestSrcDir = Join-Path $ProjectRoot "src\test\java"
$LibDir = Join-Path $ProjectRoot "lib"

# JUnit configuration
$JUnitJar = Join-Path $LibDir "junit-platform-console-standalone-1.9.3.jar"
$JUnitDownloadUrl = "https://repo1.maven.org/maven2/org/junit/platform/junit-platform-console-standalone/1.9.3/junit-platform-console-standalone-1.9.3.jar"

# Compilation order (dependencies first)
$CompilationOrder = @(
    "com\presentationchoreographer\exceptions\*.java",
    "com\presentationchoreographer\utils\*.java", 
    "com\presentationchoreographer\core\model\*.java",
    "com\presentationchoreographer\xml\parsers\*.java",
    "com\presentationchoreographer\xml\writers\*.java"
)

function Write-BuildLog {
    param([string]$Message, [string]$Level = "INFO")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $prefix = switch($Level) {
        "ERROR" { "[ERROR]" }
        "WARN"  { "[WARN] " }
        "INFO"  { "[INFO] " }
        "SUCCESS" { "[✓]    " }
        default { "[INFO] " }
    }
    
    $color = switch($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "SUCCESS" { "Green" }
        default { "White" }
    }
    
    Write-Host "$timestamp $prefix $Message" -ForegroundColor $color
}

function Test-JavaInstallation {
    try {
        $javaVersion = java -version 2>&1 | Select-String "version" | Select-Object -First 1
        Write-BuildLog "Java detected: $javaVersion" "SUCCESS"
        return $true
    }
    catch {
        Write-BuildLog "Java not found in PATH. Please install Java 11+ and add to PATH." "ERROR"
        return $false
    }
}

function Initialize-BuildEnvironment {
    Write-BuildLog "Initializing build environment..."
    
    # Test Java installation
    if (-not (Test-JavaInstallation)) {
        exit 1
    }
    
    # Create directories
    @($BuildDir, $TestBuildDir, $LibDir) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
            Write-BuildLog "Created directory: $_"
        }
    }
    
    # Download JUnit if missing
    if (-not (Test-Path $JUnitJar)) {
        Write-BuildLog "Downloading JUnit 5..."
        try {
            Invoke-WebRequest -Uri $JUnitDownloadUrl -OutFile $JUnitJar -UseBasicParsing
            Write-BuildLog "JUnit downloaded successfully" "SUCCESS"
        }
        catch {
            Write-BuildLog "Failed to download JUnit: $($_.Exception.Message)" "ERROR"
            Write-BuildLog "Please manually download JUnit to: $JUnitJar" "WARN"
        }
    }
}

function Invoke-Clean {
    Write-BuildLog "Cleaning build artifacts..."
    
    if (Test-Path $BuildDir) {
        Remove-Item -Recurse -Force $BuildDir
        Write-BuildLog "Removed build directory" "SUCCESS"
    }
    
    # Clean any class files that might be in source directories
    Get-ChildItem -Path $ProjectRoot -Recurse -Filter "*.class" | Remove-Item -Force
    Write-BuildLog "Cleaned .class files" "SUCCESS"
}

function Invoke-Compile {
    Write-BuildLog "Starting compilation..."
    
    # Ensure build directory exists
    if (-not (Test-Path $BuildDir)) {
        New-Item -ItemType Directory -Path $BuildDir -Force | Out-Null
    }
    
    $totalErrors = 0
    
    # Compile in dependency order
    foreach ($pattern in $CompilationOrder) {
        $sourcePattern = Join-Path $SrcDir $pattern
        $sourceFiles = Get-ChildItem -Path $sourcePattern -ErrorAction SilentlyContinue
        
        if ($sourceFiles) {
            Write-BuildLog "Compiling: $pattern"
            
            $javacArgs = @(
                "-d", $BuildDir,
                "-classpath", $BuildDir,
                "-Xlint:unchecked",
                "-Xlint:deprecation"
            )
            
            if ($Verbose) {
                $javacArgs += "-verbose"
            }
            
            $javacArgs += $sourceFiles.FullName
            
            try {
                $output = & javac $javacArgs 2>&1
                $exitCode = $LASTEXITCODE
                
                if ($exitCode -eq 0) {
                    Write-BuildLog "✓ $($sourceFiles.Count) files compiled" "SUCCESS"
                    if ($Verbose -and $output) {
                        $output | ForEach-Object { Write-BuildLog "  $_" }
                    }
                } else {
                    Write-BuildLog "Compilation failed for: $pattern" "ERROR"
                    $output | ForEach-Object { Write-BuildLog "  $_" "ERROR" }
                    $totalErrors += 1
                }
            }
            catch {
                Write-BuildLog "Failed to execute javac: $($_.Exception.Message)" "ERROR"
                $totalErrors += 1
            }
        } else {
            Write-BuildLog "No files found for pattern: $pattern" "WARN"
        }
    }
    
    if ($totalErrors -eq 0) {
        Write-BuildLog "Compilation completed successfully" "SUCCESS"
        return $true
    } else {
        Write-BuildLog "$totalErrors compilation errors occurred" "ERROR"
        return $false
    }
}

function Invoke-CompileTests {
    Write-BuildLog "Compiling tests..."
    
    if (-not (Test-Path $JUnitJar)) {
        Write-BuildLog "JUnit not found. Cannot compile tests." "ERROR"
        return $false
    }
    
    # Ensure test build directory exists
    if (-not (Test-Path $TestBuildDir)) {
        New-Item -ItemType Directory -Path $TestBuildDir -Force | Out-Null
    }
    
    # Find test files
    $testFiles = Get-ChildItem -Path $TestSrcDir -Recurse -Filter "*.java" -ErrorAction SilentlyContinue
    
    if (-not $testFiles) {
        Write-BuildLog "No test files found" "WARN"
        return $true
    }
    
    Write-BuildLog "Compiling $($testFiles.Count) test files..."
    
    $classpath = @($BuildDir, $TestBuildDir, $JUnitJar) -join ";"
    
    $javacArgs = @(
        "-d", $TestBuildDir,
        "-classpath", $classpath,
        "-Xlint:unchecked"
    )
    
    if ($Verbose) {
        $javacArgs += "-verbose"
    }
    
    $javacArgs += $testFiles.FullName
    
    try {
        $output = & javac $javacArgs 2>&1
        $exitCode = $LASTEXITCODE
        
        if ($exitCode -eq 0) {
            Write-BuildLog "Test compilation successful" "SUCCESS"
            return $true
        } else {
            Write-BuildLog "Test compilation failed" "ERROR"
            $output | ForEach-Object { Write-BuildLog "  $_" "ERROR" }
            return $false
        }
    }
    catch {
        Write-BuildLog "Failed to compile tests: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Invoke-Tests {
    Write-BuildLog "Running tests..."
    
    if (-not (Test-Path $JUnitJar)) {
        Write-BuildLog "JUnit not found. Cannot run tests." "ERROR"
        return $false
    }
    
    $classpath = @($BuildDir, $TestBuildDir, $JUnitJar) -join ";"
    
    $javaArgs = @(
        "-classpath", $classpath,
        "org.junit.platform.console.ConsoleLauncher",
        "--classpath", "$BuildDir;$TestBuildDir",
        "--scan-class-path"
    )
    
    if ($Verbose) {
        $javaArgs += "--details", "verbose"
    }
    
    try {
        Write-BuildLog "Executing test suite..."
        $output = & java $javaArgs 2>&1
        $exitCode = $LASTEXITCODE
        
        # Display test output
        $output | ForEach-Object { 
            if ($_ -match "SUCCESSFUL|PASSED") {
                Write-BuildLog $_ "SUCCESS"
            } elseif ($_ -match "FAILED|ERROR") {
                Write-BuildLog $_ "ERROR"  
            } else {
                Write-BuildLog $_
            }
        }
        
        if ($exitCode -eq 0) {
            Write-BuildLog "All tests passed" "SUCCESS"
            return $true
        } else {
            Write-BuildLog "Some tests failed" "ERROR"
            return $false
        }
    }
    catch {
        Write-BuildLog "Failed to run tests: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Show-Help {
    Write-Host @"
Presentation Choreographer Build Script

USAGE:
    .\Build.ps1 [TARGET] [-Verbose]

TARGETS:
    compile     Compile main source code only
    test        Compile and run tests
    clean       Clean all build artifacts
    all         Clean, compile, and test (default)
    help        Show this help message

OPTIONS:
    -Verbose    Enable detailed compilation output

EXAMPLES:
    .\Build.ps1                    # Full clean build and test
    .\Build.ps1 compile            # Compile main code only
    .\Build.ps1 test -Verbose      # Run tests with detailed output
    .\Build.ps1 clean              # Clean build artifacts

BUILD STRUCTURE:
    build/              Compiled main classes
    build/test/         Compiled test classes  
    lib/                Downloaded dependencies (JUnit)

REQUIREMENTS:
    - Java 11+ in PATH
    - Internet connection (for JUnit download)
    - PowerShell 5.1+
"@
}

function Invoke-BuildTarget {
    param([string]$Target)
    
    $success = $true
    
    switch ($Target.ToLower()) {
        "clean" {
            Invoke-Clean
        }
        "compile" {
            Initialize-BuildEnvironment
            $success = Invoke-Compile
        }
        "test" {
            Initialize-BuildEnvironment
            $success = Invoke-Compile
            if ($success) {
                $success = Invoke-CompileTests
                if ($success) {
                    $success = Invoke-Tests
                }
            }
        }
        "all" {
            Invoke-Clean
            Initialize-BuildEnvironment
            $success = Invoke-Compile
            if ($success) {
                $success = Invoke-CompileTests
                if ($success) {
                    $success = Invoke-Tests
                }
            }
        }
        "help" {
            Show-Help
            return
        }
        default {
            Write-BuildLog "Unknown target: $Target" "ERROR"
            Show-Help
            exit 1
        }
    }
    
    # Final status
    if ($success) {
        Write-BuildLog "Build target '$Target' completed successfully" "SUCCESS"
        exit 0
    } else {
        Write-BuildLog "Build target '$Target' failed" "ERROR"
        exit 1
    }
}

# Main execution
Write-BuildLog "Presentation Choreographer Build System"
Write-BuildLog "Target: $Target | Verbose: $Verbose"
Write-BuildLog "==========================================`n"

Invoke-BuildTarget $Target
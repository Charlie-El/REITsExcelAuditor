param(
    [switch]$WithOCR
)

$ErrorActionPreference = "Stop"
$ProjectRoot = $PSScriptRoot
$IconPath = Join-Path $ProjectRoot "reit_excel_auditor\assets\app_icon.ico"
$TemplateDir = Join-Path $ProjectRoot "standard_templates"
$ConfigDir = Join-Path $ProjectRoot "config"

if (-not (Test-Path -LiteralPath $IconPath)) {
    throw "Application icon was not found: $IconPath"
}
if (-not (Test-Path -LiteralPath $TemplateDir)) {
    throw "Standard template folder was not found: $TemplateDir"
}
if (-not (Test-Path -LiteralPath $ConfigDir)) {
    throw "Config folder was not found: $ConfigDir"
}

Write-Host "Checking PyInstaller..."
python -m PyInstaller --version *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host "PyInstaller not found. Installing PyInstaller..."
    python -m pip install pyinstaller
}

if ($WithOCR) {
    Write-Host "Checking optional OCR dependencies..."
    python -c "import rapidocr_onnxruntime" *> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "RapidOCR not found. Installing optional OCR dependencies..."
        python -m pip install -r requirements-ocr.txt
    }
}

Write-Host "Building REITs Excel auditor executable..."
$pyinstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--icon", $IconPath,
    "--add-data", "$IconPath;reit_excel_auditor\assets",
    "--add-data", "$TemplateDir;standard_templates",
    "--add-data", "$ConfigDir;config",
    "--hidden-import", "win32com.client",
    "--hidden-import", "pythoncom",
    "--hidden-import", "pywintypes",
    "--exclude-module", "pandas",
    "--exclude-module", "matplotlib",
    "--exclude-module", "IPython",
    "--exclude-module", "pytest",
    "--exclude-module", "pygame",
    "--exclude-module", "scipy",
    "--exclude-module", "zmq",
    "--exclude-module", "cryptography",
    "--exclude-module", "gevent",
    "--exclude-module", "lxml",
    "--exclude-module", "torch",
    "--exclude-module", "torchvision",
    "--exclude-module", "torchaudio",
    "--exclude-module", "tensorflow",
    "--exclude-module", "keras",
    "--exclude-module", "transformers",
    "--exclude-module", "modelscope",
    "--exclude-module", "datasets",
    "--exclude-module", "sklearn",
    "--exclude-module", "skimage",
    "--exclude-module", "dask",
    "--exclude-module", "distributed",
    "--exclude-module", "numba",
    "--exclude-module", "llvmlite",
    "--exclude-module", "pyarrow",
    "--exclude-module", "bokeh",
    "--exclude-module", "altair",
    "--exclude-module", "selenium",
    "--exclude-module", "notebook",
    "--exclude-module", "jupyterlab",
    "--exclude-module", "django",
    "--exclude-module", "langchain",
    "--exclude-module", "nltk",
    "--exclude-module", "geopandas",
    "--exclude-module", "folium",
    "--exclude-module", "lightning",
    "--exclude-module", "bitsandbytes",
    "--exclude-module", "xformers",
    "--exclude-module", "paddle",
    "--exclude-module", "paddleocr",
    "--name", "REITsExcelAuditor",
    "--distpath", "dist",
    "--workpath", "build\pyinstaller",
    "--specpath", "build"
)

if ($WithOCR) {
    Write-Host "OCR build enabled. The executable will be larger."
    $pyinstallerArgs += @(
        "--hidden-import", "rapidocr_onnxruntime",
        "--collect-data", "rapidocr_onnxruntime",
        "--collect-binaries", "onnxruntime",
        "--collect-binaries", "cv2"
    )
} else {
    $pyinstallerArgs += @(
        "--exclude-module", "numpy",
        "--exclude-module", "PIL",
        "--exclude-module", "cv2",
        "--exclude-module", "onnxruntime",
        "--exclude-module", "rapidocr_onnxruntime"
    )
}

$pyinstallerArgs += "reit_excel_auditor\app.py"
python -m PyInstaller @pyinstallerArgs

Write-Host ""
Write-Host "Build finished: dist\REITsExcelAuditor.exe"
if ($WithOCR) {
    Write-Host "OCR support: included RapidOCR local OCR engine."
} else {
    Write-Host "OCR support: not included. Run .\build_exe.ps1 -WithOCR to include local OCR."
}

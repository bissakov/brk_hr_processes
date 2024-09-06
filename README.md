## Installation

```bat
git clone https://github.com/bissakov/brk_hr_processes.git
cd brk_hr_processes

py -3.12 -m venv .venv
.venv/Scripts/activate
python -m pip install --upgrade pip
pip install -r .\requirements.txt
```

## Usage

### Batch script

```bat
.\start.bat
```

### CMD

```bat
set "PYTHONPATH=%cd%;%PYTHONPATH%" && python .\robots\main.py
```

### PowerShell

```bat
$env:PYTHONPATH = "$env:PYTHONPATH;$pwd"; python .\robots\main.py
```
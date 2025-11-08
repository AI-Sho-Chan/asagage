param()
$ErrorActionPreference = 'Stop'
& python "scripts/guard_no_openpyxl_xlsm.py"
exit $LASTEXITCODE


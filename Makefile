combine:
	cscript .\vbac.wsf combine

decombine: .\src\workbook.xlsm\Module1.bas
.\src\workbook.xlsm\Module1.bas: .\bin\workbook.xlsm
	cscript .\vbac.wsf decombine

XLSM_PATH = .\bin\workbook.xlsm
run:
	pwsh .\Run-Macro.ps1 $(XLSM_PATH) test

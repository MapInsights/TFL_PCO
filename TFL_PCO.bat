taskkill /IM Outlook.exe /f
start "" "%ProgramFiles(x86)%\Microsoft Office\Office14\outlook.exe"


@echo on
"C:\Program Files\R\R-3.2.3\bin\R.exe" CMD BATCH C:\Programs\gtc_tasks\TfL_PCO\Driver_Licences.R
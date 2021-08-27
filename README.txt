1. Download source work report from sigp apps 
2. Save it in this folder with name workReport.xls
3. Run WorkReportConverter from shortcut


Configuration
=============
source 
  - source work report path (downloaded from sigp apps)
destination 
  - destination work report path and file name
  - Converter appends _YYYYMM.xlsx automatically
year, month 
  - Converter creates work report for this month


SPRINT, PBI and TASK ids
========================
Converter parses SPRINT, PBIs and TASKs from description. 
Expression is removed if parsed successfully 

Example description:
I Was working on these tasks. (PBI: 000000, 111111) (TASK: XXXXXX, YYYYYY) (SPRINT: 0X.Y)

Results in:

 Task 				              | PBU/BUG	       | Task	         | Sprint
------------------------------------------------------------------------------
 I Was working on these tasks.    | 000000, 111111 | XXXXXX, YYYYYY  | 0X.Y

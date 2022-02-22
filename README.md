# ExcelComparator
This project compares two excel file and produce a compared file and a log file

--------------------------------------------------------------------------------------------------------------------------------------
	This is a excel comparator project, generally used to compare a report with two different configuration.

	Pre-requisite:
		We need to have report output of two configuration in Excel format.
		
	Steps to compare two files
	Step1: Run ExcelReader GUI.exe file
	Step2: Click on "Browse First File" and select the first file to be compared
	Step3: Click on "Browse Second File" and select the second file to be compared
	Step4: Click on Compare and wait for the comparision completion
	Step5: Click on exit

	Output: We will get two files as output.
	1. Compared File <Date and Time>.excel :-> This is the file which has all mismatch cells highlighted in red color.
	2. Log File <Date and Time>.txt :-> This is the log file having all the mismatches.

--------------------------------------------------------------------------------------------------------------------------------------
	We also have an executable file that will give excel output having "Mismatch?" column appended after last column
	We can filter based on Mismatch? column as Yes to get all the mistached column.
	
	
	
Technology Stack:
	Python is the main technology used to build this project.
	Two main python libraries are used here
	1. For UI we are using tkinter library
	2. For Excel manipulation we have used openpyxl

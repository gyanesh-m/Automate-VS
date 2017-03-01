# Automate-VS

This project was developed as part of summer internship at BHEL Valves dept. during 15th May - 15thJuly 2016.This helps to automate the
task of normalisation and unification of unstructured data in excel.

## Prerequisites

* The input file has to be in .xls format and the output file is saved in the current directory with name 'Output.xls'.

* The following files need to be in the same folder as exe with the same filename.
V_CS.xls , V_FM.xls , V_FH.xls , V_TAHH.xls
These files are required for the complete online part number extraction process.The default material number and price columns for these files 
are assumed as 2 and 6 (indexed from 0).If you have different column numbers then update them in the part number extraction
class under main function.Also mention the database server ip to connect in the part number extraction class under main function.

* For the offline part number fetch ,set the name of your file to "Offline database.xls" and change the default values before use as mentioned in the main2
function under Part number extraction class.The default column values for the part number/description,
valve code and price  are 0,8 and 14.

* For the rating extraction you need to setup your Soap client and server accordingly .Mention the server to connect for the soap
 client in main code function at line 277 inside the base class.

## Requirements
You need to install the following libraries .

* python 2.7.13
* suds==0.4
* openpyxl==2.3.0
* xlrd==1.0.0
* ttk==0.31
* Tkinter 
* cx_Oracle

##Usage
If the prerequisites and requirements are met then program can be run as follows:

` python gui.py `

##Customisations for usage

There are various customisations which can be done according to the need of the user .All changes have to be done in the Standard
Values.xls file. The changes have been discussed sheetwise and the sheet names should not be changed. 
 
**standard**

This sheet’s top row can be modified to include any new heading to be searched for and to be included in the “final” sheet in the output file.If a new heading is added in standard sheet it 
needs to be added in alias_col sheet also.It is important for the reader to know that the once a heading is put in one sheet for a variable then the variable’s heading needs to be same in 
different sheets. For example­ If the user wants to add Price column in the final sheet provided it is given in the 
input sheet, he just needs to add it to the right side of specialities column and this entry should 
also be added in the alias_col sheet with the same name if the entry has aliases. 
 
**alias**  

This sheet contains the alias for different items in a particular category.The top level entry is 
considered the standard entry and rest entries below it are considered aliases for it.This works 
as follows:Say an input sheet writes SA182F91 instead of F91. So if the user wants that all the entries in 
the sheet where SA182F91 is written should be converted to F91, then he just needs to add that 
value under F91 in the alias sheet.It is important to note however that between two different sets 
of entries there should be a blank cell left.Also the different sets of entries should be entered in 
the corresponding column category.For example any new materal and its aliases should be 
entered in ‘VALVE MATERIAL’category and not in any other category. 

**alias_col**  

This sheet contains the different names for a particular column name which can be used to 
identify a particular column.The standard value is  kept as top entry and its aliases are kept 
below it.. Also when searching for the aliases it is searched bi directional,i.e., If the word to be 
searched was ‘PRES’ and in the aliases we have a similar word ‘PRESS’ under ‘PRESSURE’ ,it 
will search ‘PRESS’ in ‘PRES’ and also ‘PRES’ in ‘PRESS’.If any one condition is satisfied it 
returns the column number corresponding to the  top level standard entry which in this case will 
be the column number corresponding to the PRESSURE.It is important to note that the top level 
entry in this sheet also needs to be mentioned below it in the aliases.For example: in first 
column PRESSURE the entries mentioned below it are PRESS and PRESSURE.Also the entry 
should be in CAPS and any special character to be searched should not be taken care 
of(should be omitted) because when searching is done the code first converts the headings into 
alphanumeric form in uppercase and then only they are searched bi directionally with the 
aliases.For example,If a heading in the input sheet represents pressure as PRESS(ATM), only 
PRESS would be required to be entered into the alias col sheet under the heading of 
PRESSURE. 

**conversion**

This sheet contains the conversion info about inches and nb.If any new entry is to be added in 
the future then it can be added anywhere in this sheet below the corresponding column.It is 
important to note that the order of the columns should not be changed,i.e. The first column 
should always be NB and the second column should always be INCHES. 
 
**table** 

This sheet contains the information to convert the data into the corresponding sap format to 
extract part numbers online.The entries in this sheet can be modified and the changes will take 
place accordingly except for the rating case for which changes need to be made in the 
bhel_rating sheet. 
 
**bhel_rating** 

This sheet displays the rating which BHEL manufactures. Thus if in future BHEL starts 
manufacturing a new rating in a particular category the user will just have to add the value in the 
particular category.For example in TAHH if BHEL starts manufacturing 1690 then the user will 
just have to add C1690 in the 1690 row below TAHH column. 
 
**db alias** 

This sheet is similar to alias sheet except for the fact that the entries in the corresponding 
column are different.The functionality remains the same.The top entry for aliases are defined 
differently and here also they can be modified according to the user needs. 

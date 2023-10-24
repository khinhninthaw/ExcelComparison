
We enable to Comparison Excel file using the EPPlus Library.
We run the macro in c# using Microsoft.Office.Interop.Excel.

## Environment
I'm developed using Visual Studio 2022 and Windows 11 Pro.

## Introduction and Usage

Project is Excel Comparison.
There are two text box and two button.

In the first text box need to write the excel file path. 
eg. D:\Miyamoto from Global Walkers (Myanmar)\programmingTest\test.xlsm

and then you can click on Compare button. 
After clicking on the Compare button, system will check invalid value and then system filled out the wrong cells with yellow color on the ToBeChecked sheet in this excel file.

In the Second text box need to write macro name.
and then you can click on Run button. 
After that system read the macro in this excel and then system will show alert box.


## About This Tool

I am using the follow Library and Dll for this project.

1. EPPlus
> EPPlus Library is using for excel value comparison.


2. Microsoft.Office.Interop.Excel
> Microsoft.Office.Interop.Excel Library is using for run the macro.


3. office.dll
> Interop.Excel Version 15 is not support for Office 2016 so that I am also added the Office.dll.
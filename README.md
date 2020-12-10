# Solidworks-Macros README

### Project Overview
The main goal of the project was to learn to use the SolidWorks API and to automate the design of tabs. The basic function of the code it to create a tab in Solidworks CAD based on dimensions inputted by the users. The user can then view the part file and continue to modify dimensions until they are ready to save the file.

### Installation
In order to run this code you will need access to Solidworks and a coding environment in windows. I used 
[Solidworks 2019](https://www.solidworks.com/sw/support/downloads.htm),
[Anaconda for Windows](https://www.anaconda.com/products/individual), and
[Git for Windows](https://gitforwindows.org/).


An python wrapper called PyWin32 also needs to be installed. It allows you to gain access to the Win32 API and allows you to create and use COM objects. To install use the following: 
```
conda install pywin32
```

### Start-Up
```
swYearLastDigit = 9
sw = win32com.client.Dispatch("SldWorks.Application.%d" % (20+(swYearLastDigit-2))) 
```

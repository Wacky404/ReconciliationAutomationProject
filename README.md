# ReconciliationAutomationProject

## Project Description
### Public Information and Non-Sensitive Data Sources

This is a data pipeline script that takes an Excel worksheet's data through reconciliation and cleansing using the ope.ed.gov 
and nces database alongside Openpyxl, OpenAI API, and the Google Places API. The development of this script was goal oriented and followed the 
specific constraints of the data project.

## How to Use this Script

The [DataFile.py](DataFile.py) file contains the base code and the [Datapipeline.py](DataPipeline.py) is the master file 
that will run all script functionality. There are two files that contain the .xlsx databases of the forementioned 
sources. 

- **If you are just using this script you only need to build the project and run [DataPipeline.py](DataPipeline.py) and follow the prompts 
in the console.**

- **If you are planning on changing any of the scripts functions or methods you can find the script's functionality/code in the 
[DataFile.py](DataFile.py) and other supporting files.**

## Script Design

This script has its functionality housed under a single class, called DataFile, here you will find a structure of: 
> class DataFile
> - static methods
> 
> - class methods
> > - reconcile_x
> > - ai_x
> > - cleanse_x

Find and edit the neccessary method in order to fix any bugs that might appear along the data pipeline or add additional functionality! 

*MIT License*

*Copyright (c) 2023 Wacky404*

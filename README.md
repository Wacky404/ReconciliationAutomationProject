# ReconciliationAutomationProject

## Project Description

This is a data pipeline script that takes an Excel worksheet's data through reconciliation and cleansing using the ope.ed.gov 
and nces database alongside Openpyxl and the OpenAI API. The development of this script was goal oriented and followed the 
specific constraints of the data project.

## How to Use this Script

There are two primary .py files that contain the code and two other files that contain the .xlsx database of the forementioned 
sources. 

- **If you are just using this script you only need to build the project and run DataPipeline.py and follow the prompts 
in the console.**

- **If you are planning on changing any of the scripts functions or methods you can find the script's functionality/code in the 
reconcile.py.**

## Script Design

This script has it's functionality housed under a single class, called DataFile, here you will find a structure of: 
> class DataFile
> - static methods
> 
> - class methods
> > - reconcile_x
> > - cleanse_x
> > - ai_x

Find and edit the neccessary method in order to fix any bugs that might appear along the data pipeline or add additional functionality! 

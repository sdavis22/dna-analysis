# Data Analysis Python Module

## Summary
This project is used to format DNA sequencing data. Various metrics are calculated and written to a summary sheet, and the raw data used to calculate these summary metrics is also written to one sheet.

## Instructions
 - Download or clone this project. Unzip it and navigate to it using your linux terminal
 - Ensure that you have both Python3 and pip installed
    - For Python3, you can just follow the instructions on [python.org](python.org/downloads)
    - Pip should come installed with the Python3 installation
 - Run `pip install -r requirements.txt`
 - Run `python3 analysis.py`
 - Wait 10-15 minutes depending on the size and number of files
 - Check outputs in the `output` directory

## Inputs
 - For the purposes of this tutorial, assume the name of the input file is `sample.xlsx`
 - Ensure that list.txt contains the names of all the files you want to analyze, separated by newlines (i.e. `sample` exists in list.txt)
 - For each name, have an input file in the `input` directory, i.e. `input/sample.xlsx`
 - Additionally, ensure you have a key file in the `keys` directory, i.e. `keys/sample-key.xlsx`
 - Output files will be generated in the `output` directory as `output/sample-output.xlsx`

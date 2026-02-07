

[![DOI](https://zenodo.org/badge/617394969.svg)](https://doi.org/10.5281/zenodo.18510283)


# REEEVR: Excel-to-R converter

## Acknowledgements
GitHub repository with code for project entitled "Reliable and Efficient Estimation of the Economic Value of medical Research (REEEVR)" funded by UK Medical Research Council grant MR/W029855/1.

The file tokenizer.py is a modified version of the tokenizer from openpyxl version 3.1, and has been edited to correctly handle dynamic ranges. This file will be removed once openpyxl has been updated to correctly handle dynamic ranges. 

We utilise PyQt6 for the user interface, which requires the GPL-3.0 license. As such all code (other than tokenizer.py which is covered under openpyxl MIT license) is also covered under the GPL-3.0 license. 

## The repo

This repository contains the code to automatically take an Excel workbook and output an R file equivalent. While strictly able to convert any workbook, we have designed this process around expecting Health Technology Assessment (HTA) workbooks, and so would liekly require modification of the outputs for other uses. 

The converter is currently able to convert 111 functions, the list of which can be found in the file "Excel function list.xlsx".

## Usage
To use this software you have two options:
1. Download the source from the Release tab and run the python directly. The master python file to run is "reeevrconverter.py"
2. Download the pre-compiled binary for windows. This will provide a .exe file that can be ran as usual. You will still need to download the source folder to access the R code.

We highly recommend reading the manual provided in this repository to understand how to use the software. 

### I am unfamilar with Github
To download the software you need to head to the Releases tab.
- Make sure you are on the Code tab of the repository. 
- Look on the right hand side and select the releases tab
- If there are multiple releases, select the release version you wish to use (we always recommend the latest version)
- Download the source code as either a .zip or .tar.gz file. You do not need both.
- If you want the pre-compiled binary you can also download REEEVR-Excel-to-R-converter.exe




## Feedback and Issues

Please raise issues using the Github issue tracker. These issues can be programmatic (bugs, error message, etc), failures of the documentation, or other feedback (GUI changes, feature requests, etc).

## Development and Contribution

Obviously this being a public repository, anyone can fork and develop on their own, so long as they abide by the licensing terms.

However, should you wish to contribute to this project, please generate a pull request and we will endeavour to work with you. Please note this is done on a voluntary basis by the maintainers.

### Testing
When contributing please run all code through the test suites in the test folder.
In addition please add tests if required.

### Compiling binaries
We currently use pyinstaller to create the binaries. The command used is:
`pyinstaller reeevrconverter.py -F --windowed --icon logo.png --hidden-import openpyxl.cell._writer -n REEEVR-Excel-to-R-converter --splash logo.png`


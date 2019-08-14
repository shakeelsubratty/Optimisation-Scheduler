
# Paycheck Optimisation Allocation

This Python script uses a Constraint Programming solver (Google's ORTools CP-Solver) to perform the monthly allocation of Paycheck administrators to Payroll tasks.

The script takes in employee and payroll information, provided by the Excel sheet **"Optimisation_Spreadsheet.xlsx"**,
as input, and outputs to the Excel sheet **"Optimisation_Allocation"** the following:
* The monthly allocation of payrolls for each employee.
* A visualisation of the capacity utilisation across the month.

# Prerequisites

Before attempting to run or build this script, the following must be installed:

* Python v3.6.4+
* pip
* virtualenv - for convenience



# Dependencies

The script relies on the following dependencies:

* pandas
* numpy
* or-tools
* xlrd
* openpyxl
* easygui
* pyinstaller - for building executable

If using this script in a new virtual environment (or not using virtualenv at all), each dependency can be installed by running:

``` pip install <dependency> ``` where \<dependency\> is as given above.

If using this script with it's supplied virtual environment, simply activate the virtual environment with the following cmd/terminal command:

### Windows

``` Scripts/activate.bat ```

### Unix

``` source bin/activate```

# Build procedures
## Running script

The script can be run with the following command:

```python paycheck_optimisation.py```

The Excel files needed for the script to function must be placed in the same directory as the paycheck_optimsation.py file.

Ensure the script can be run successfully before attempting to build an executable.

## Build script into executable

To build the script into executable, run the command:

``` pyinstaller -D paycheck_optimisation.py ```

### Windows

On Windows, it has been found that pyinstaller fails to identify a number of hidden numpy dependencies. To alleviate this, a number of imports can be found at the top of the main file:

``` 
# import numpy.random.common
# import numpy.random.bounded_integers
# import numpy.random.entropy 
```

If having trouble with building on Windows, uncommenting these imports should help.

### Executable

The -D argument builds the executable into a single directory:

``` dist/paycheck_optimisation ```

where the script can be accessed by running the paycheck_optimsation executable. Again, ensure that the required Excel files are in the same directory as the executable.


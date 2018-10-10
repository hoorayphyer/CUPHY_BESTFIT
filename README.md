# CUPhysLabToolkit

## How to install
### Fresh install
1. Go to this [link](https://raw.githubusercontent.com/hoorayphyer/CUPhysLabToolkit/master/cuphys_bestfit), right click and choose Select All, right click and choose Copy
2. Open Excel, press Alt-F11 to go to Visual Basic Editor
3. From the menu bar, click Insert -> Module
4. Right click and choose Paste. Now you should see the code appear
5. (optional) You are free to rename this module in the bottom left window. An example name would be CUPhysLab.
6. (optional) If you want to reuse this module in the future but don't want to go over the above everytime, in the Visual Basic Editor, from the menu bar, click File->Export File, choose a location and make sure the save type is .bas
7. You can close the Visual Basic Editor window
### Import from the saved .bas file
1. Open Excel, press Alt-F11 to go to Visual Basic Editor
2. From the menu bar, click File->Import File, select the saved .bas file
3. You can close the Visual Basic Editor window


## All-in-One BestFit
CUPhysLabToolkit provides functions for streamlining the bestfit process. All you need is to have four columns of data, namely xdata, xerror, ydata, yerror. A sample set of data is as follows.

| x data  |x error   | y data   | y error  | 
|---|---|---|---|
| 1 | 0.1  | 10  | 2  |
| 2 | 0.1  | 20  | 4  |
| 3|  0.1 |  30 | 6  |
| 4|  0.1 |  40 | 8  |
| 5|  0.1 | 50  | 10  |

### LinearLeastChiSquaredFit
The calling procedure is similar to LINEST but slightly modified.
1. highlight a 2-by-4 range of cells ( 2 rows 4 columns )
2. press "=LinearLeastChiSquaredFit", followed by selection of xdata, xerror, ydata and yerror, in that order
3. press "Ctrl-Shift-Enter". On Mac OS, if this didn't work out, try "Cmd-Shift-Enter"
Now the results should be formatted as follows

| slope  | error in slope   | intercept   | error in intercept  |
|---|---|---|---|
| {value}  |  {value} |  {value} |{value}   |

LinearLeastChiSquaredFit is meant to substitute the built-in LINEST in Excel because the latter doesn't factor in the uncertainty in the data. Technically speaking, LINEST is performing a linear least square regression whose error for the fitted quantities, namely slope and intercept in this case, merely represents the extent to which the (x,y) data points are not lying on the same line. Using LINEST often produces highly underestimated uncertainties in slope and intercept, which consequently will skew the comparison with norminal values.


### CUPHYS_BestFit_Plot
The calling procedure is similar to above. 
1. select a blank cell
2. press "=CUPHYS_BestFit_Plot", followed by selection of xdata, xerror, ydata and yerror, in that order
3. press "Enter"

This should generate a chart automatically with custom error bars, trendline with function displayed and R^2 value, x and y axes labels for the user to modify.

The motivation for providing this function is because of the horrendous number of steps involved to produce a decent plot of scientific quality in Excel, and it's even more troubling that generating such plots is almost one half of a physics lab's task, with the other half being error propagation. The common practice in other programmable environments is to create a routine for generating plots with common features.

### CUPHYS_BestFit(to be implemented)
This function will combine the above two together so that the user only need to call one function to get two things done. Currently it is experiencing some issue

## Trigonometric Functions In Degrees( to be implemented )

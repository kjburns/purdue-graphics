# purdue-graphics
An excel spreadsheet with VBA scripting that produces purdue split regime charts from VISSIM data collection and signal change output

Copyright 2019 Oklahoma Department of Transportation. See License.

As the code is originally posted here, it has several assumptions:
1. Your signal groups are displayed on stop bars with SC = x and SG = y
2. Your data collection points are placed upstream of each stopbar(x, y) and the number attribute is (100 * x + y) * 10 + C, where C is a unique number for each dc point between 0 and 9, inclusive

When your VISSIM model runs, the data collections need to be configured to output raw data collection (which will end up in MER files) and signal changes (which will end up in LSA files). You will need to manually move your data to the spreadsheet, as I haven't written any code yet to read it from a file. 

Currently, the references in the spreadsheet are hard-coded in the VBA script, and there are array ranges, loop ranges, and the like that specifically target the data set built into the spreadsheet. Ideally, this will eventually be refactored using tables in excel so that column names can be used instead. It even may be converted to a python script that takes text files as inputs and produces a text file or csv as output, since VBA inherently produces code held together by duct tape instead of proper fasteners.

Hard coded values (this list is hopefully exhaustive):
* Available stopline numbers (module-level array declarations)
* Column numbers for table fields
* Number of model runs
* Length of each simulation run (86400 seconds)
* Maximum simulation length is 99999 seconds.

Currently any phase that is between the begin of green and 5 seconds after end of red at the end of the simulation run is lost, and is listed as a TODO in the function cleanupEndOfRun(). However, in this case the data is not complete: you could have a partial green interval, a whole green interval without any red, or a whole green interval with less than 5 seconds of red. In any case, this is an incomplete data point that might be thrown out anyway, and so it hasn't been implemented and may never be.

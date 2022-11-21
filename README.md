# open-merchandise-management
Merchandise program management

Allow to automatically actualize the stock of merchandise according to the sells and the delivery of each when running the program.

The sells are written on the "Sorties" sheet and the deliveries on the "Entrées" sheet, you can change the name, just make sure to also change
the string (corresponding to the sheet name) in the source code between line 23 and 27.
So the file you want to modify is "stock.xlsx".

The original features that provides this program is that it allows you to search for the volume of sells, deliveries or stock between two different date
not necerssarily knowing the date. In fact the program will take as a time interval the values that are the closest to the given one.
At the end of the run a new results file will be created.
If you want to modify the number given at the end of the results file's name, modify the number in "dedicated.xlsx".

The programm offers shortcuts ("a" to actualy refresh the stock according to deliveries and sells, "w" to search between dates, and "q" to quit) 

The program has an visual interface.

![Capture d’écran du 2022-11-21 12-32-57](https://user-images.githubusercontent.com/114911243/203041110-070dbb2e-59ff-4705-9eb0-729f30f1c8e4.png)

![Capture d’écran du 2022-11-21 12-33-15](https://user-images.githubusercontent.com/114911243/203041971-4dd67455-d492-4340-83ca-ea53b1eeeeec.png)

There is an exe if you just want to click to run the program (use wine for linux)

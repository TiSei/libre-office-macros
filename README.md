# Libre Office Macros

small collection of useful functions to generate .csv-files out of Libre Office

**Installation**<br>
To use all functions of the collection, the System-Module and CSV-Module must be installed as macros in Libre Office. For correct execution, the "System.Startup" method must be assigned to the "open Document" event.

A further development as an extension is planned

## CSV-Module

supported methods and INLINE functions:
- ChangeSelector: set the used selector for the current document (permanent storage)
- SaveAsCSV: the current selection is saved as .csv-file
- CSVLINE: combines cells with document specific selector
- SCANUP: searches for the next non-empty cell value above the specified position
- DATETOTEXT/TIMETOTEXT: convertes time-representing numbers into text

## TODO

- [x] add scan up feature
- [ ] export as extension
- [ ] increase performance
- [ ] add custom menu to menu bar
- [ ] add custom dockable panels

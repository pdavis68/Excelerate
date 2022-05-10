# Excelerate
Fast C# library for creating .xlsx files.

# Features

- Fast writing of Excel .xlsx files.
- Supports adding worksheets
- Supports addeding fonts and styles (colors, fills, etc)

# Notes

The idea for this project came from the FastExcel library which is VERY fast at writing Excel files, but also extremely limited (no ability to add new worksheets, not colors or fonts, etc).

Most of the other Excel libraries out there utilize the OpenXML libraries, which limits the speed of populating and writing large Excel files. 

If absolute speed is your goal and you have no need for worksheet manipulation or decorating your spreadsheets, then FastExcel is the library for you. 

If, on the other hand, you're willing to compromise a little on speed and a lot on features, you can get Excelerate which is somewhere in-between. Excelerate is very fast (I will post benchmarks comparing various libraries at some point), but not as fast as FastExcel. In addition to having more features, Excelerate is also, I think, easier to use than FastExcel, in that you are free to deal with data as a grid, instead of being restricted to FastExcel's row-oriented flow.



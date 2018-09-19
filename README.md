# xladd-util

Simple Excel utilities based on the xladd library. This is a standalone dynamic load library (xll) that can be loaded into Excel and which adds some useful generic functions to Excel.

# Loading

The xll can be loaded into Excel by simply using File/Open to open it, or by registering it as an addin, or by using VBA to load it. The VBA loader method may be useful if you have other addins, and want to control either the load order or the environment.

# Functions

* xuTranspose - Similar to the built-in Excel function TRANSPOSE, in that it transposes an array. Rows turn into columns and vice versa. There are a number of problems with TRANSPOSE: for example, it only works if it is inside an array formula. xuTranspose works wherever it is used.

* xuGlueCols - Takes a number of arguments, all of which are optional. It takes each of the arguments from left to right, treating each as an array and gluing them together from left to right to make a single array. If the arguments are ranges with different numbers of rows, the resulting array is the size of the largest, with the holes padded with #NA.

* xuGlueRows - The same as xuGlueCols but vertically stacked rather than horizontally.

See also the help inside the Excel Function Wizard

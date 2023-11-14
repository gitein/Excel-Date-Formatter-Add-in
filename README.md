# Excel-Date-Formatter-Add-in
An Excel add-in that converts the date format from American to standard for all sheet's cells.

# 🔩 How to add and activate the add-in in MS Excel:

1. Open the following path betweem quotes via Runbox "%USERPROFILE%\AppData\Roaming\Microsoft\AddIns" via shortcut key WIN + R
2. Copy and paste the file named "Date-Formatter-Add-in.xlam"
3. Open Excel
4. Click "Options" in the side pane
5. Select "Add-ins" tab
6. Click "Go" next to "Manage", if you don't find the add-in listed and checked, click "Browse" to open a dialog where to select the add-in file
7. Select the file after making sure it's pasted to the directory/path as explained above:

Excel Add-in file to be placed in the following directory on your system %USERPROFILE%\AppData\Roaming\Microsoft\AddIns
__________________________

# 🚸 Limitations and specifications of add-in:

1. Switches between these formats for cells in the current active sheet when add-in run, i. e. MM/DD/YYYY and DD/MM/YYYY on each next click with a message box indicating change has been applied.
2. The algorithm can't distinguish between dates with where MM and YY are less than 13, thus any further clicks will toggle e.g. 05/10/YYYY and 10/05/YYYY back and forth.
3. It checks if the date to be formatted found in the current calendar of runtime, if not the change to be applied on the current cell will be skipped, e.g. if it's not a leap year, then the cell with date e.g. 02/29/YYYY will not be changed.
4. Any distinguishable date found in the current calendar during runtime e.g. 04/26/YYYY will be changed to 26/04/YYYY just once. (Any further clicks won't toggle it back)


## License

This project is licensed under the MIT with Attribution License.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:


THE SOFTWARE IS PROVIDED "AS IS," WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

2. **Commercial Use:** The Software or any derivative works thereof
   may be used for commercial purposes without the prior written
   permission of the Author.




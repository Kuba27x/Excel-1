#Excel-1

A range in Excel is a collection of two or more cells. To fill a range select cell with value or cells with pattern and click on the lower right corner of that cell and drag it down. Excel automatically fills the range based on the pattern of the first two values.
To move a range, select a range and click on the border of the range then drag the range to its new location. You can insert/delete rows/columns by right clicking them and selecting insert/delete. 
You can change the width of a column by clicking and dragging the right border of the column header. To automatically fit the widest entry in a column, double click the right border of a column header. It also works for rows; double click bottom border of a row header.

AutoFit

![screenshot](Screenshots/AutoFit.png)
Original column width.

![screenshot](Screenshots/AutoFitResult.png)
To automatically fit the widest entry in a column, double click the right border of a column header.

Auto Fill

![screenshot](Screenshots/AutoFill.png)
Here is an example of Autofill based on pattern, hour and date.

![screenshot](Screenshots/AutoFillOptions.png)
Available options of date AutoFill.

![screenshot](Screenshots/AutoFillResult.png)
Result of filling by months.

Custom list Auto Fill

It is possible to create your own list to easily fill a range in Excel. Here is how to do it.

![screenshot](Screenshots/Settings.png)
On the File tab click Options, under Advanded go to General and click Edit Custom Lists.

![screenshot](Screenshots/CustomList.png)
Add your list here.

![screenshot](Screenshots/CustomListResult.png)
Type London in cell P1 then click on the lower right corner of cell P1 and drag it down. Effect above. (Note: Custom lists are stored in computer`s registry so they can be used in other workbooks)

Flash Fill

![screenshot](Screenshots/FlashFill.png)
Here we are trying to use Flash Fill to create emails based on pattern. 

![screenshot](Screenshots/FlashFillResult.png)
Select cell C8 and press CTRL + E (Flash Fill shortcut). Result above.

![screenshot](Screenshots/FlashFill2.png)
Another example. Split contents of one cell into multiple cells.

![screenshot](Screenshots/FlashFillResult2.png)
Press CTRL + E in each column. (Note: flash fill in Excel only works when it recognizes a pattern.)

It is important to know that output received using Flash Fill will not automatically update when the source data changes. When you know source data will change use formulas instead.

TEXTSPLIT

![screenshot](Screenshots/Textsplit.png)
Instead of using Flash Fill you can use TEXTSPLIT(text; col_delimiter) formula as shown above. (Note: TEXTSPLIT is only available in Excel365/Excel2021) (Note: TEXTSPLIT fills multiple cells. This is called spilling in Excel365/2021)

SEQUENCE

![screenshot](Screenshots/Sequence.png)
Here we use SEQUENCE(rows;columns;start;step) to fill multiple cells. (Note: SEQUENCE is only available in Excel365/Excel2021) (Note: SEQUENCE fills multiple cells. This is called spilling in Excel365/2021)




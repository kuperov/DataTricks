DataTricks
==========

This is a simple Mathematica package for reading Excel data. It is 
particularly useful for economic datasets.

Mathematica's standard tools struggle with large workbook files, and 
can not have a spreadsheet refresh its data sources and recalculate.
It also allows the use of range definitions (in R1C1 format) and named
ranges, so that ranges coordinates do not have to be hard-coded.

Future versions will also include features for time series data.

This example loads two data ranges from different sheets in the same 
workbook. The second sheet ('Sheet2') contains a named range called 'NamedRange'.

    Needs["DataTricks`"]
    {data1, data2} = importData["file.xlsx", {{"Sheet1", "A:B"},{"Sheet2", "NamedRange"}}, 
                                dropHeaderRows->3]

DataTricks is released under the [LGPL licence, v3](https://www.gnu.org/licenses/lgpl.html).
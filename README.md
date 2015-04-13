DataTricks
==========

This is a simple Mathematica package for reading Excel worksheets,
intended for use with economic datasets.

Mathematica's standard tools struggle with large workbook files, and 
can not have a spreadsheet refresh its data sources.
DataTricks also allows range definitions (in R1C1 format) and named
ranges, so that range coordinates do not have to be hard-coded.

This package is experimental, so use at your own risk. If you discover
problems, please [file a bug report](https://github.com/kuperov/DataTricks/issues/new).

To install, download DataTricks.m to your computer and choose Install... from Mathematica’s File menu. Choose ‘File’ as the source, and select the downloaded file. Then, in your notebook, include the following line:

    Needs["DataTricks`"]

DataTricks provides two functions: _importExcel[...]_ and _importExcelTS[...]_ for time-series data.

Documentation is available in *demo.nb*, or [at this page](http://kuperov.github.io/DataTricks/).
Unit tests are available in *test_script.nb*, 
or [at this page](http://kuperov.github.io/DataTricks/test_script/index.html).



*The MIT License (MIT)*

Copyright (c) 2015 Alex Cooper

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

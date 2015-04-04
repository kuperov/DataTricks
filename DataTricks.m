(* ::Package:: *)

BeginPackage["DataTricks`"]

importExcel::usage="Load data from a spreadsheet. Specify a list of {sheet, range} pairs to load. By default, the spreadsheet's sources are refreshed first.";
importExcelTS::usage="Load time series data from a spreadsheet.";

importExcelTS::nosheet="Sheet `1` was not found.";
importExcel::nosheet="Sheet `1` was not found.";

importExcelTS::datecolumnletter="Can not identify date column: '`1`' is not a valid Excel column identifier. For example, use 'A' to mean the range 'A:A'.";

refreshWorkbook::usage = "Indicates that the workbook's sources should be refreshed prior to loading data.";
dropHeaderRows::usage="Number of rows from the top of the requested range to drop. Useful for header rows in time series.";

openWorkbook::nofile="File `1` was not found during import";
openWorkbook::comfailure="Failed to start Excel using COM. Ensure Excel (2010 or later) and the Primary Interop Assemblies (PIA) are installed.";

Options[importExcel]={
	refreshWorkbook-> True,
	dropHeaderRows-> 0
};

Options[importExcelTS]={
	refreshWorkbook -> True,
	dropHeaderRows -> 0,
	dateColumn -> "A"
};

Begin["`Private`"]

Needs["NETLink`"];
NETLink`LoadNETType/@{"System.GC","System.DateTime"};

 (* helpers *)
(* deal with date types and/or in-band errors(!?!), recursively preserving list structure *)
fixValues[xs_List]:=fixValues/@xs;
fixValues[d_]:=If[NETLink`NETObjectQ[d] && NETLink`InstanceOf[d,"System.DateTime"],
					DateList[d@ToString["f"]],
					If[NumberQ[d]&&d==-2146826246,
						Missing["NA"],
						N@d]
				];

(* for a range or named range within a sheet *)
grabData[wb_Symbol,dropRows_Integer,{sheetName_String,rangeName_String}]:=Module[{sheet,range},
	sheet = wb@Sheets@Item[sheetName];
	If[!NETLink`NETObjectQ[sheet],
		Message[importExcel::nosheet,Style[sheetName,Red]];
		Return[$Failed];
	];
	range=wb@Application@Intersect[sheet@Range[rangeName],sheet@UsedRange];
	Drop[fixValues@range@Value[],dropRows]
];
(* for a whole sheet *)
grabData[wb_Symbol,dropRows_Integer,{sheetName_String}]:=Module[{sheet},
	sheet = wb@Sheets@Item[sheetName];
	If[!NETLink`NETObjectQ[sheet],
		Message[importExcel::nosheet,Style[sheetName,Red]];
		Return[$Failed];
	];
	Drop[fixValues@sheet@UsedRange@Value[],dropRows]
];
(* Specialised version of grabData, designed for time series data,
 * data are arranged in columns, and one column (by default, 'A')
 * contains a list of dates corresponding to the reference period. *)
grabTSData[wb_Symbol,dropRows_Integer,dateColumn_String,{sheetName_String,rangeName_String}]:=Module[
		{sheet,range,topRow,bottomRow,dateRange,combinedRange},
	sheet = wb@Sheets@Item[sheetName];
	If[!NETLink`NETObjectQ[sheet],
		Message[importExcelTS::nosheet,Style[sheetName,Red]];
		Return[$Failed];
	];
	If[!StringMatchQ[dateColumn,RegularExpression["^[a-zA-Z]$"]],
		Message[importExcelTS::datecolumnletter,Style[dateColumn,Red]];
		Return[$Failed];
	];
	range=wb@Application@Intersect[sheet@Range[rangeName],sheet@UsedRange];
	 (*get start and end rows of data range BEFORE dropping headers *)
	topRow = range@Row;
	bottomRow = range@Row + range@Rows@Count - 1;
	dateRange = sheet@Range[dateColumn<>ToString[topRow]<>":"<>dateColumn<>ToString[bottomRow]];
	combinedRange = Transpose@Join[Transpose@dateRange@Value[],Transpose@range@Value[]];
	 (*drop headers, and keep only rows that start with a list *)
	Cases[Drop[fixValues@combinedRange,dropRows],{_List,___}]
];
(* Open workbook. This should be called within a NETBlock[ ... ] and you should
 * invoke Close[False] on this object when you're finished. You should also signal
 * a .NET garbage collection afterwards to clear COM objects that have been allocated.
 * 
 * Raises messages if Excel can't be instantiated or the file doesn't exist. *) 
openWorkbook[filename_String,refresh_] := Module[{excel,fileToOpen},
	(* emit a warning (but don't bail) if we can't see the file -- but not if it looks like a URL *)
	fileToOpen = If[StringMatchQ[filename,RegularExpression["^(http|HTTP)[sS]?://"]],
		filename, (* it's a URL, move on *)
		If[FileExistsQ[FindFile[filename]],
			FindFile[filename],
			If[!FileExistsQ[filename],Message[openWorkbook::nofile,Style[filename,Red]]];
			filename
		]
	];
	(* You need the Office PIAs installed for this to work. See
	 * https://msdn.microsoft.com/en-us/library/15s06t57.aspx and
	 * http://www.microsoft.com/en-au/download/details.aspx?id=3508 *)
	NETLink`InstallNET[];
	(* we can't recycle an open Excel instance (which would be much faster) 
     * in case the user already has the file open (chaos would ensue) *)
	excel = NETLink`CreateCOMObject["Excel.Application"];
	If[!NETLink`NETObjectQ[excel],
		Message[openWorkbook::comfailure];
		Return[$Failed]
	];
	Return[excel@Workbooks@Open[fileToOpen,If[refresh,3,0]]];
];

(* Main user-facing function, with some convenience overloads *)
importExcel[filename_String,sheet_String,opt : OptionsPattern[]] :=
	importExcel[filename, {{sheet}}, opt];
importExcel[filename_String,sheet_String, range_String,opt : OptionsPattern[]] :=
	importExcel[filename, {{sheet, range}}, opt][[1]];
importExcel[filename_String,{"Sheets", ranges_List}, opt : OptionsPattern[]] := 
	importExcel[filename, ranges, opt];
importExcel[filename_String, {sheet_String, ranges_List}, opt:OptionsPattern[]] := 
	importExcel[filename, {sheet, #}&/@ ranges, opt];
importExcel[filename_String, ranges_List, opt : OptionsPattern[]] := Module[
	{wb,excel,dataSet,rowsToDrop},
	NETLink`NETBlock[
		wb = openWorkbook[filename,OptionValue[refreshWorkbook]];
		rowsToDrop=OptionValue[dropHeaderRows];
		dataSet = grabData[wb,rowsToDrop,#]&/@ranges;
		wb@Close[False];
	];
	(* Clean up COM objects we've allocated and released as the NETBlock has ended *)
	GC`Collect[];
	dataSet
];

(* Specialised version of importExcel[], designed for time series data. *)
importExcelTS[filename_String, {sheet_String, ranges_List}, opt : OptionsPattern[]] := 
	importExcelTS[filename, {sheet, #}&/@ ranges, opt];
importExcelTS[filename_String, {"Sheets", ranges_List}, opt : OptionsPattern[]] := 
	importExcelTS[filename, ranges, opt];
importExcelTS[filename_String, sheet_String, range_String, opt : OptionsPattern[]] :=
	importExcelTS[filename, {{sheet, range}}, opt][[1]];
importExcelTS[filename_String, ranges_List, opt: OptionsPattern[]] := Module[
	{wb,dataSet,rowsToDrop,dateCol},
	NETLink`NETBlock[
		wb = openWorkbook[filename,OptionValue[refreshWorkbook]];
		rowsToDrop=OptionValue[dropHeaderRows];
		dateCol=OptionValue[dateColumn];
		dataSet = grabTSData[wb,rowsToDrop,dateCol,#]&/@ranges;
		wb@Close[False];
	];
	(* Clean up COM objects we've allocated and released as the NETBlock has ended *)
	GC`Collect[];
	dataSet
];

End[];
EndPackage[];




(* ::Package:: *)

BeginPackage["DataTricks`"]

importExcel::usage="Loads data from a spreadsheet.";
importExcel::nofile="File `1` was not found during import";
importExcel::comfailure="Failed to start Excel using COM. Ensure Excel and the Primary Interop Assemblies (PIA) are installed.";

refreshWorkbook::usage = "Indicates that the workbook's sources should be refreshed prior to loading data.";
dropHeaderRows::usage="Number of rows from the top of the requested range to drop. Useful for header rows in time series.";

Options[importExcel]={
	refreshWorkbook-> True,
	dropHeaderRows-> 0
};

Begin["`Private`"]

Needs["NETLink`"];
LoadNETType["System.GC"];
LoadNETType["System.DateTime"];

(* helpers for ImportData *)
(* deal with date types and/or in-band errors(?!), preserving list structure *)
fixValues[xs_List]:=fixValues/@xs;
fixValues[d_]:=If[NETObjectQ[d] && NETLink`InstanceOf[d,"System.DateTime"],DateList[d@ToString[]],
If[NumberQ[d]&&d==-2146826246,Missing["NA"],N@d]];
(* for a range or named range within a sheet *)
grabData[wb_Symbol,{sheetName_String,rangeName_String}]:=Module[{sheet,range},
	sheet = wb@Sheets@Item[sheetName];
	If[!NETObjectQ[sheet],Return[Missing["Sheet"]]];
	range=wb@Application@Intersect[sheet@Range[rangeName],sheet@UsedRange];
	fixValues@range@Value[]
];
(* for a whole sheet *)
grabData[wb_Symbol,{sheetName_String}]:=Module[{sheet},
	sheet = wb@Sheets@Item[sheetName];
	If[!NETObjectQ[sheet],Return[Missing["Sheet"]]];
	fixValues@sheet@UsedRange@Value[]
];

(* User-facing function, with some convenience overloads *)
importExcel[filename_String,sheet_String,opt : OptionsPattern[]]:=importExcel[filename,{{sheet}}, opt];
importExcel[filename_String,sheet_String, range_String,opt : OptionsPattern[]]:=importExcel[filename,{{sheet, range}}, opt];
importExcel[filename_String,{"Sheets", ranges_List}, opt : OptionsPattern[]] := importExcel[filename, ranges, opt];
(* The main import function. It would have been simpler to do this the other way around (have the main function be the one
 * that acts on only a single range), but this allows us to use the same instance of Excel for every range that will be loaded. *)
importExcel[filename_String,ranges_List, opt : OptionsPattern[]] := Module[{wb,excel,totalDataSet},
	(* You need the Office PIAs installed for this to work. See
	 * https://msdn.microsoft.com/en-us/library/15s06t57.aspx and
	 * http://www.microsoft.com/en-au/download/details.aspx?id=3508 *)
	NETBlock[
		InstallNET[];
		(* we can't recycle an open Excel instance (which would be much faster) in case the user 
		 * has the file open (chaos would ensue) *)
		excel= CreateCOMObject["Excel.Application"];
		If[!NETObjectQ[excel],
			Message[importExcel::comfailure, Style[Red]];
			Return[$Failed]
		];
		If[!FileExistsQ[filename],
			Message[importExcel::nofile,Style[filename,Red]];
			Return[$Failed]
		];

		wb=excel@Workbooks@Open[filename,If[OptionValue[refreshWorkbook],3,0], True];
		totalDataSet = Drop[grabData[wb,#], OptionValue[dropHeaderRows]]&/@ranges;
		wb@Close[False];
	];
	(* Clean up COM objects we've allocated and released as the NETBlock has ended *)
	GC`Collect[];

	totalDataSet
];

End[]
EndPackage[]




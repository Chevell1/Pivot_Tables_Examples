ods noresults;

ods path(prepend) work.templat(update);

filename temp url "http://support.sas.com/rnd/base/ods/odsmarkup/tableeditor/tableeditor.tpl";
%include temp;

/* Create sample data source */

ods excel file="c:\\temp\\temp.xlsx" options(sheet_name="Table_1");

proc print data=sashelp.prdsale;
run;

ods excel close;

/* Create Pivot tables */
 
ods tagsets.tableeditor file="c:\temp\pivot.js"
    options(update_target="c:\\temp\\temp.xlsx"
            output_type="script"
            sheet_name="Table_1"
			pivotrow="product"
            pivotcol="year"
            pivotdata="actual"
            Data_Bgcolor="beige"
            Data_Fgcolor="brown"
            Border_Color="red"
            Header_Fgcolor="orange"
			Header_BgCOLOR="black"
            Rowheader_Fgcolor="orange"
            Rowheader_Bgcolor="black"
            Datalabel_Fgcolor="green"
          );

data _null_;
file print;
put _all_;
run;

ods tagsets.tableeditor close;

x "c:\temp\pivot.js";

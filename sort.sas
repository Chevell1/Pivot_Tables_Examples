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
            pivot_sort="product,2 ~ quarter,2"
            pivotpage_filter="Canada"
            pivotpage="Country"
            pivotcol="quarter"
            pivotrow="product"
            pivot_format="light1" pivotdata="actual");

    data _null_;
      file print;
      put _all_;
    run;

ods tagsets.tableeditor close;

x "c:\temp\pivot.js";
 

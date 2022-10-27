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
        pivot_series="yes"
        dashboard="yes"
        pivotrow= "product  |product|product"
        pivotcol= "country  |country|country"
        pivotdata="actual   |actual |actual"
        pivotdata_stats="sum| max   |min"
        ptdest_range="a3,g3,m3"
        pivotcharts="yes"
        chart_type="doughnut"
        chart_location="Table_1_Pivot"
        chart_position="150,0,250,250|150,350,250,250|150,700,250,250");

data _null_;
  file print;
  put "test";
run;

ods tagsets.tableeditor close;

x "c:\temp\pivot.js";

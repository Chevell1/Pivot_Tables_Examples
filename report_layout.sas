ods noresults;

/* Create sample data source */

ods excel file="c:\\temp\\temp.xlsx" options(sheet_name="Table_1");

proc print data=sashelp.prdsale;
run;

ods excel close;


/* Create pivot tables */

ods tagsets.tableeditor file="c:\temp\pivot.js"                                                                                             
     options (update_target="c:\\temp\\temp.xlsx"
              output_type="script"
              sheet_name="Table_1"
              pivot_layout="tabular"
              pivot_format="light23"
              pivotdata_grandtotal="yes" 
              Pivotdata_subtotals="yes"                                                                                                           
              pivotrow= "country,region"                                                                                     
              pivotcol="year"                                                                                             
              pivotdata="actual"); 
 
data _null_;
file print;
put _all_;                                                                                                        
run;
 
ods tagsets.tableeditor close;
x "c:\temp\pivot.js";    

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
            pivotrow="country"                                                                                                      
            pivotcol="year"                                                                                                         
            addfield="returns =actual*.05,
                      net_sales =actual-returns"                                                                
            pivotdata="actual,returns,net_sales"                                                                                    
            pivotdata_caption="Sales,Return Amount,Net Sales"    
            pivotdata_fmt="$#,###.###.00");                                                                                                                            
 
 data _null_;
   file print;
   put "test";
run;

ods tagsets.tableeditor close;

x "c:\temp\pivot.js";

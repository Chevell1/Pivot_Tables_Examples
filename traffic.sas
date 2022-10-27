ods noresults;

ods path(prepend) work.templat(update);

filename temp url "http://support.sas.com/rnd/base/ods/odsmarkup/tableeditor/tableeditor.tpl";
%include temp;

/* Create sample data source */

ods excel file="c:\\temp\\temp.xlsx" options(sheet_name="Table_1");

proc print data=sashelp.prdsale;
run;

ods excel close;

ods tagsets.tableeditor file="c:\temp\pivot.js"                                                                                          
 options(update_target="c:\\temp\\temp.xlsx"  
         output_type="script"   
         sheet_name="Table_1"    
         format_condition="databar,b2:c16 ~ 
                           iconsets,d2:d16,6 ~                                                                                 
                           colorscale,e2:e16 ~                                                                                  
                           cellvalue,f2:f16,6,6000 ~                                                                          
                           expression,g2:g16,0,=mod(row(),2)=0" 
                           pivotrow="month"   
                           pivotcol="product"                                                                                                          
                           pivotdata="actual");

data _null_;
   file print;
   put _all_;
run;

ods tagsets.tableeditor close; 


x "c:\temp\pivot.js";


 

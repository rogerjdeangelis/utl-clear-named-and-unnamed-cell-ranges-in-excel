# utl-clear-named-and-unnamed-cell-ranges-in-excel
Clear named and unnamed cell ranges in excel
    Clear named and unnamed cell ranges in excel                                                                 
                                                                                                                 
      Two solutions                                                                                              
                                                                                                                 
          1. Clearing a named range (sql drop table and proc datasets delete).                                   
          2. clearing an unnamed range                                                                           
                                                                                                                 
    https://tinyurl.com/y2nx3j9y                                                                                 
    https://github.com/rogerjdeangelis/utl-clear-named-and-unnamed-cell-ranges-in-excel                          
                                                                                                                 
    SAS forum                                                                                                    
    https://tinyurl.com/y5zdrttw                                                                                 
    https://communities.sas.com/t5/SAS-Programming/cleae-content-from-excel-using-sas/m-p/582027                 
                                                                                                                 
    other excel repos                                                                                            
    https://tinyurl.com/y3p2pqcs                                                                                 
    https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=         
                                                                                                                 
    *_                   _                                                                                       
    (_)_ __  _ __  _   _| |_                                                                                     
    | | '_ \| '_ \| | | | __|                                                                                    
    | | | | | |_) | |_| | |_                                                                                     
    |_|_| |_| .__/ \__,_|\__|                                                                                    
            |_|                                                                                                  
    ;                                                                                                            
                                                                                                                 
                                                                                                                 
    %utlfkil(d:/xls/class.xlsx);                                                                                 
    libname xel "d:/xls/class.xlsx";                                                                             
    data xel.class;                                                                                              
        set sashelp.class;                                                                                       
    run;quit;                                                                                                    
    libname xel clear;                                                                                           
                                                                                                                 
    d:/xls/class.xlsx                                                                                            
                                                                                                                 
      +----------------------------------------------------------------+                                         
      |     A      |    B       |     C      |    D       |    E       |                                         
      +----------------------------------------------------------------+                                         
    1 | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |                                         
      +------------+------------+------------+------------+------------+                                         
    2 | ALFRED     |    M       |    14      |    69      |  112.5     |                                         
      +------------+------------+------------+------------+------------+                                         
       ...                                                                                                       
      +------------+------------+------------+------------+------------+                                         
    N | WILLIAM    |    M       |    15      |   66.5     |  112       |                                         
      +------------+------------+------------+------------+------------+                                         
                                                                                                                 
    [class]                                                                                                      
                                                                                                                 
    *            _               _                                                                               
      ___  _   _| |_ _ __  _   _| |_                                                                             
     / _ \| | | | __| '_ \| | | | __|                                                                            
    | (_) | |_| | |_| |_) | |_| | |_                                                                             
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                            
                    |_|                                                                                          
    ;                                                                                                            
                                                                                                                 
    d:/xls/class.xlsx                                                                                            
                                                                                                                 
      +----------------------------------------------------------------+                                         
      |     A      |    B       |     C      |    D       |    E       |                                         
      +----------------------------------------------------------------+                                         
    1 |            |            |            |            |            |                                         
      +------------+------------+------------+------------+------------+                                         
    2 |            |            |            |            |            |                                         
      +------------+------------+------------+------------+------------+                                         
       ...                                                                                                       
      +------------+------------+------------+------------+------------+                                         
    N |            |            |            |            |            |                                         
      +------------+------------+------------+------------+------------+                                         
                                                                                                                 
    [class]                                                                                                      
                                                                                                                 
    *                               _                                                                            
     _ __   __ _ _ __ ___   ___  __| |  _ __ __ _ _ __   __ _  ___                                               
    | '_ \ / _` | '_ ` _ \ / _ \/ _` | | '__/ _` | '_ \ / _` |/ _ \                                              
    | | | | (_| | | | | | |  __/ (_| | | | | (_| | | | | (_| |  __/                                              
    |_| |_|\__,_|_| |_| |_|\___|\__,_| |_|  \__,_|_| |_|\__, |\___|                                              
                                                        |___/                                                    
    ;                                                                                                            
                                                                                                                 
    * delete then create named range.;                                                                           
                                                                                                                 
    %utlfkil(d:/xls/class.xlsx);                                                                                 
    libname xel "d:/xls/class.xlsx";                                                                             
    data xel.class;                                                                                              
        set sashelp.class;                                                                                       
    run;quit;                                                                                                    
    libname xel clear;                                                                                           
                                                                                                                 
    /* delete named range using datasets and sql */                                                              
                                                                                                                 
    libname xel "d:/xls/class.xlsx";                                                                             
       proc datasets lib=xel;                                                                                    
         delete class;                                                                                           
       run;quit;                                                                                                 
    libname xel clear;                                                                                           
                                                                                                                 
    * NOTE: Deleting XEL.class (memtype=DATA). */                                                                
                                                                                                                 
    libname xel "d:/xls/class.xlsx";                                                                             
       proc sql;                                                                                                 
          drop table xel.class                                                                                   
       ;quit;                                                                                                    
    libname xel clear;                                                                                           
                                                                                                                 
    * NOTE: Table XEL.class has been dropped.;                                                                   
                                                                                                                 
    *                                           _                                                                
     _   _ _ __  _ __   __ _ _ __ ___   ___  __| |  _ __ __ _ _ __   __ _  ___                                   
    | | | | '_ \| '_ \ / _` | '_ ` _ \ / _ \/ _` | | '__/ _` | '_ \ / _` |/ _ \                                  
    | |_| | | | | | | | (_| | | | | | |  __/ (_| | | | | (_| | | | | (_| |  __/                                  
     \__,_|_| |_|_| |_|\__,_|_| |_| |_|\___|\__,_| |_|  \__,_|_| |_|\__, |\___|                                  
                                                                    |___/                                        
    ;                                                                                                            
                                                                                                                 
    * delete then create named range;                                                                            
    %utlfkil(d:/xls/class.xlsx);                                                                                 
    libname xel "d:/xls/class.xlsx";                                                                             
    data xel.class;                                                                                              
        set sashelp.class;                                                                                       
    run;quit;                                                                                                    
    libname xel clear;                                                                                           
                                                                                                                 
    libname xel "d:/xls/class.xlsx";                                                                             
       proc sql;                                                                                                 
          drop table xel.'A1:E21'n;                                                                              
       ;quit;                                                                                                    
    libname xel clear;                                                                                           
                                                                                                                 
    NOTE: Table XEL.A1:E21 has been dropped.                                                                     
                                                                                                                 
                                                                                                                 

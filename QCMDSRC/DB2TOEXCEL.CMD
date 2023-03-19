             CMD        PROMPT('Convert table to Excel') +           
                          HLPID(DB2TOEXCEL) +                        
                          HLPPNLGRP(DB2TOEXCEL)                      
             PARM       KWD(FILENAME) TYPE(Q1) MIN(1) PROMPT('File + 
                          name')                                     
             PARM       KWD(TOFILENAME) TYPE(*CHAR) LEN(64) +        
                          PROMPT('To Stream File name')              
             PARM       KWD(IFSPATH) TYPE(*CHAR) LEN(160) +          
                          DFT(*USRPRF) PROMPT('IFS Path Name')       
             PARM       KWD(REPLACE) TYPE(*CHAR) LEN(4) RSTD(*YES) + 
                          DFT(*NO) SPCVAL((*YES) (*NO)) +            
                          PROMPT('Replace file')                     
 Q1:         QUAL       TYPE(*NAME) MIN(1)                           
             QUAL       TYPE(*CHAR) LEN(10) DFT(*LIBL) +             
                          PROMPT('Library name')                     
             QUAL       TYPE(*GENERIC) LEN(10) DFT(*FIRST) +         
                          SPCVAL((*FIRST *FIRST) (*LAST *LAST) +     
                          (*ALL *ALL)) PROMPT('Member name')         

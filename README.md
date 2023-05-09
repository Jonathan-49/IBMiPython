# db2toexcel.py
Description:  
 This python script will read one or more members (partitions)
  of an IBM i Db2 table and populate an Excel.  
  One worksheet is created per member. The member name is used as the worksheet name. 
  
Pip packages needed to be installed on IBM i:  
 pip3 install ibm_db_dbi  
 pip3 install xlsxwriter
 
Runs on IBM i OS 7.3 TR 13 and 7.4 TR 7
 
It can be run from the IBM i PASE command line.  
 `python3 db2toexcel.py 'libraryname' 'tablename' 'membername'  'IFSfolder' 'nameofExcel' 'userid'`   
 ``` python3 db2toexcel.py -h  
usage db2toexcel [-h]  library table member folder filename username   
 positional arguments:                                                            
  library       library name                                                     
  table         table name                                                       
  member        Member name; can be the name of a member or a generic name       
                qualified by an asterisk (*); Special values can also be used;   
                '*ALL', '*FIRST', '*LAST' if name is blank then '*FIRST' is      
                used.                                                            
  folder        folder path                                                      
  filename      output file name                  
  username      User name                         
                                                   
 optional arguments:                               
   -h, --help    show this help message and exit  
```
 
 
 Or it can be run via a CL command created using a wrapper command, see the following link
 to download an install the command QSHPYRUN https://github.com/richardschoen/QshOni
 
 Once the command QSHPYRUN is installed on the IBM i;  the source to build command *db2toexcel* is in folders  
 QCLLESRC, QCMDSRC and QPNLSRC of this repo. 
 
 Below is the CL command *db2toexcel* prompted from IBM i ACS.  
 'CL: DB2TOEXCEL FILENAME(libname/tablename) TOFILENAME('Excelname') IFSPATH('/pathname') REPLACE(*YES)'
 
 ![image](https://github.com/Jonathan-49/IBMiPython/assets/62209270/5117f52c-b838-4993-8d8f-804d42be3032)
  
 
 ![image](https://user-images.githubusercontent.com/62209270/226182358-9e2facce-8519-4c26-a3a7-a0eac46c8232.png)
 
                                                                                                                                                    
It will check that the   
table exists and if user has permission.                          
                                                                  
The folder is also checked if it exists.                          
                                                                  
The name of Excel can entered in the Stream file name parameter.  
                                                                  
If no stream file name is given then the table name is used.      
                                                                  
The Excel will be created in the IFS Path parameter. The default  
path is *USRPRF (creates file in /home/ + userid).                
The current IBM i user id will be used as the author in the Excel   
properties.                                                         
                                                                    
FILENAME File (table) name, library name and optionally member name.
                                                                    
Member name can be a specific name or a generic name  qualified by  
an asterisk (*).   Special values '*ALL', '*FIRST' or '*LAST' can   
also be used in the member name.                                    
                                                                    
TOFILENAME Excel name. There is no need to put a suffix, the        
extension .xlsx will be added automatically. If a suffix is added,  
then it will be removed.                                            
                                                                    
IFSPATH The name of the IFS path.    

The default path is *USRPRF (creates file in /home/ + userid).  
                                                                
REPLACE *YES - will replace the file if already exists.         
                                                                    

 

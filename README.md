# db2toexcel.py
Description:  
 This python script will read 1 or more members (partitions)
  of an IBM i Db2 table and populate an Excel.  
  One worksheet is created per member.  
  
Pip packages needed:  
 pip3 install ibm_db_dbi  
 pip3 install xlsxwriter
 
It can be run from the IBM i PASE command line.  
 `python3 db2toexcel.py 'libraryname' 'tablename' 'membername'  'IFSfolder' 'nameofExcel' 'userid'` 
 

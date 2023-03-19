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
 
 Or it can be run via a CL command created using https://github.com/richardschoen/QshOni
 
 ![image](https://user-images.githubusercontent.com/62209270/226181930-12bf753e-3bb7-4428-bb6a-c23b753a0f17.png)

 

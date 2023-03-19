#!/usr/bin/python
# Name: db2toexcel.py
import ibm_db_dbi as db2
import argparse
import xlsxwriter
from pathlib import Path
from pprint import pprint
import sys
import re

def validate_args(args, conn):
    """
      Validate the command line arguments, an error message is added to the list
      errormsgs when a problem is found.
      Parameters:
             args : Namespace containing arguments to validate
             conn (object): Database connection object
      Returns:
             errormsgs (list): List of error messages, empty if no problems found.
    """
    errormsgs = []
    objectcheck = conn.cursor()
    objectcheck.execute ("select count(*) from table(QSYS2.LIBRARY_INFO(UPPER('"+args.library+"')))")
    one_row=objectcheck.fetchone()
    args.folder = args.folder.strip('/')
    args.folder = "/" + args.folder + "/"
    if not Path(args.folder).exists():
        errormsgs.append("Directory " +args.folder+ " not found")
    if not Path(args.folder).is_dir():
        errormsgs.append(args.folder+ " is not a directory")
   
    if one_row[0]==0:
        errormsgs.append("Library/schema: " +args.library+ " not found")
    objectcheck.execute ("select count(*) from QSYS2.SYSTABLES where table_schema = \
    upper('"+args.library+"') and table_name  =upper('" + args.table+ "')")
    one_row=objectcheck.fetchone()
    if one_row[0]==0:
        errormsgs.append("Table: " +args.table+ " not found in library/schema " +args.library)
    if errormsgs:
        pprint(errormsgs)
        return False
    objectcheck.close()
    
    return True

def validate_membername(membername):
    errorlist = []
    
    errorlist.clear()
    if len(membername)==0:
        errorlist.append('Member name is missing')
    if len(membername)>10:
        errorlist.append('Member name cannot more than 10 long')
    if membername[0]=='*' and membername not in ['*ALL','*FIRST', '*LAST']:
        errorlist.append('Member name cannot start with a *')
    if len(membername)>0 and membername[0].isdigit():
        errorlist.append('Name Cannot start with a digit')
    if membername.count('*')>1:
        errorlist.append('More than one asterick')
    if re.match(r"[^A-Za-z0-9\*]", membername):
        errorlist.append('Contains invalid character')
   
    
    if errorlist:
        pprint(errorlist, indent=10, width=50)
        
    return errorlist

def create_sheet(library, table, member, conn, workbook, header_format):
  cur = conn.cursor()
 
  alias = "create or replace alias " + library + ".FILENAME2 for "+library+"."+  table + "("  \
  + member + ")"
  cur.execute(alias)

  query="select * from " + library + ".FILENAME2"
  cur.execute(query)
  headers = []
  col_sizes = []
  for loop, descr in enumerate(cur.description):
    headers.append(descr[0])
    col_sizes.append(descr[2])
 
  worksheet = workbook.add_worksheet(str(member))
  for colnum,  display_size in enumerate(col_sizes):
    worksheet.set_column(colnum, colnum, display_size if display_size>len(headers[colnum]) else len(headers[colnum]) )
  worksheet.write_row('A1', headers, header_format)
  worksheet.freeze_panes(1, 0)
  print(member)
  for rownum, row in enumerate(cur, start=1):
    print(row)
    worksheet.write_row(rownum, 0, row)
  
  cur.close()


  
def main():
  parser = argparse.ArgumentParser(prog='db2toexcel', \
    description="Export Db2 table and it's members to an Excel. Each member will be populated in a seperate sheet")
  parser.add_argument("library", type=str.upper, help="library name")
  parser.add_argument("table", type=str.upper, help="table name")
  parser.add_argument("member", type=str.upper, help="Member name; can be the name of a member or \
   a generic name qualified by an asterisk (*); Special values can also be used; '*ALL', '*FIRST', '*LAST' \
    if name is blank then '*FIRST' is used. ")
  parser.add_argument("folder", help="folder path")
  parser.add_argument("filename", help="output file name")
  parser.add_argument("username", type=str.upper, help="User name")
 
  args = parser.parse_args()
  newfolder = args.folder.strip('/')
  newfolder = "/" + newfolder + "/"
  newfilename = args.filename.split('.')[0]
  if args.member.strip()=='':
    args.member='*FIRST'
  if validate_membername(args.member): 
    sys.exit(1)

  conn = db2.connect()
  if conn is None:
        print("\nERROR: Unable to connect to Db2 database")
        sys.exit(1)
  usercur = conn.cursor()
  usercur.execute("values(current user)")
  currentuser=usercur.fetchone()
 
  if args.username.strip()=='':
    args.username=currentuser[0]

  if not validate_args(args, conn):
    sys.exit(1)

  readmembers = conn.cursor()

  selectmembers =  "select table_partition as membername  \
  from qsys2.collection_services_info c, qsys2.syspartitionstat a   \
  where table_schema = upper('"+args.library+"') and table_name  =upper('" + args.table+ "')"

  findmembersSQL = selectmembers + " order by create_timestamp desc";  
  if args.member=='*ALL':
    findmembersSQL = selectmembers + " order by create_timestamp desc";
  elif args.member=='*FIRST':
    findmembersSQL = selectmembers + " order by create_timestamp desc fetch first 1 rows only";
  elif args.member=='*LAST':
    findmembersSQL = selectmembers + " order by create_timestamp asc fetch first 1 rows only";
  else: 
    if args.member.count('*')==1:
      args.member = args.member.replace('*', '%')
      print(args.member)
      findmembersSQL = selectmembers + " and table_partition like upper('"+args.member+"')  \
      order by create_timestamp desc ";
    else:
      findmembersSQL = selectmembers + " and table_partition=upper('"+args.member+"')  \
      order by create_timestamp desc ";
  print(findmembersSQL)
  try:
    readmembers.execute(findmembersSQL)
  except Exception as err:
    print(f"Error reading members; ({err}) ")
 
  allmembers=readmembers.fetchall()
  if not allmembers:
    print('No members found')
  else: 
    workbook = xlsxwriter.Workbook(newfolder +newfilename +'.xlsx');
    header_format = workbook.add_format({'bold': True,
                                         'align': 'center',
                                         'valign': 'vcenter',
                                         'fg_color': '#D7E4BC',
                                         'border': 1})
    cell_format = workbook.add_format()
    cell_format.set_bold()
    workbook.set_properties({
      'title': args.filename,
      'subject':'DB2 table ' +args.table +' in library ' + args.library,
      'author': args.username,
      'category':'Python script: ' + sys.argv[0],
      'company':' ',
      'comments':'Created with Python and XlsxWriter',
      'status':'Final'})
    
    for rownum, row in enumerate(allmembers, start=1):  
        member=str(row[0])
        create_sheet(args.library, args.table, member, conn, workbook, header_format) 
    workbook.close()

  conn.close()

if __name__== "__main__":
    main()

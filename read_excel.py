import xdrlib,sys
import xlrd
from xls_log import get_xls_logger

logger=get_xls_logger()

def open_excel(file='file.xls'):
    try:
        data=xlrd.open_workbook(file)
        return data
    except Exception as e:
        print ("can't open this file: "+ str(e))

def excel_table_byindex(file='file.xls',colnameindex=0,by_index=0):
    data=open_excel(file)
    table=data.sheets()[by_index]
    nrows=table.nrows
    ncols=table.ncols
    colnames=table.row_values(colnameindex)
    print("colnames="+str(colnames))
    list=[]
    for rownum in range(colnameindex+1,nrows-colnameindex):
        row=table.row_values(rownum)
        if row:
            app={}
            for i in range(len(colnames)):
                app[colnames[i]]=row[i]
            print (app)
            list.append(app)
    return list

def excel_table_byname(file='file.xls',colnameindex=0,by_name=""):
    data=open_excel(file)
    table=data.sheet_by_name(by_name)
    nrows=table.nrows
    ncols=table.ncols
    colnames=table.row_values(colnameindex)
    print("colnames="+str(colnames))
    list=[]
    for rownum in range(colnameindex+1,nrows-colnameindex):
        row=table.row_values(rownum)
        if row:
            app={}
            for i in range(len(colnames)):
                app[colnames[i]]=row[i]
            print (app)
            list.append(app)
    return list


def excel_read_everycell(file='test.xlsx',by_name=""):
    data=open_excel(file)
    print("===sheet name: %s"%data.sheet_names())
    table=data.sheet_by_name(by_name)
    nrows=table.nrows
    ncols=table.ncols
    print("The sheet Name."+by_name+" has "+str(nrows)+" rows and "+str(ncols)+" cols")
    cell_list=[]
    for rownum in range(nrows):
        row=table.row_values(rownum)
        if row:
            cell_list.append([])
            for colnum in range(ncols):
                cell_list[rownum].append(row[colnum])
               # print("====== Row No."+str(rownum)+": "+str(cell_list[rownum]))
    return cell_list

"""
 read one  excel file and return the content of whole file
 return excel_content={'filename':, 'sheets':{'-sheetname':[list of cell],}}
"""
def read_excel_file(filename='test.xlsx',split_title=''):
    excel_content={}
    file_splitter='\\'
    if '\\' not in filename:
        file_splitter='/'
    file_name_element=filename.split(file_splitter)
    excel_content['filename']=file_name_element[len(file_name_element)-1]
    excel_content['sheets']={}
    data=open_excel(filename)
    sheets=data.sheet_names()
    for sheet_name in sheets:
        logger.debug('-Start reading the sheet%s'%sheet_name)
        table=data.sheet_by_name(sheet_name)
        nrows=table.nrows
        ncols=table.ncols
        if nrows==0:
            logger.debug('-- has %s rows %s columns, skip this sheet'%(nrows,ncols))
            continue
        cell_list=[]
        for rownum in range(nrows):
            row=table.row_values(rownum)
            if row:
                cell_list.append([])
                for colnum in range(ncols):
                    cell_list[rownum].append(row[colnum])
        excel_content['sheets'][sheet_name]=cell_list
        logger.debug('--read %s rows , %s coloums'%(nrows,ncols))

    if split_title=='': #except Jinfeng
        return excel_content
    else: #Jinfeng put all packing list into one sheet, need split_title to divide
        new_excel_content={'filename':excel_content.get('filename'),'sheets':{}}
        for sheet_name in excel_content.get('sheets'):
            #logger.debug('sheet_name=%s'%sheet_name)
            if str(sheet_name).strip().upper()=='P014806': #Jinfen's 1st sheet no useful
                continue
            origin_sheet=excel_content.get('sheets').get(sheet_name)
            sheet_count=0
            for rownum in range(len(origin_sheet)):
                current_row=origin_sheet[rownum]
                #logger.debug('-current_row is %s'%current_row)
                if str(split_title).strip().upper() in [str(cell).strip().upper() for cell in current_row]:
                    sheet_count+=1
                    new_excel_content['sheets'][str(sheet_count)]=[]
                    #logger.debug('-sheet_count=%s, new_excel_content=%s'%(sheet_count,new_excel_content))
                new_excel_content['sheets'][str(sheet_count)].append(current_row)
        return new_excel_content
                
    
        


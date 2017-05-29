import xdrlib,sys
import xlrd
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

def excel_read_everycell(file='test.xlsx',by_index=0):
    data=open_excel(file)
    table=data.sheets()[by_index]
    nrows=table.nrows
    ncols=table.ncols
    print("The sheet No."+str(by_index)+" has "+str(nrows)+" rows and "+str(ncols)+" cols")
    cell_list=[]
    for rownum in range(nrows):
        row=table.row_values(rownum)
        if row:
            cell_list.append([])
            for colnum in range(ncols):
                cell_list[rownum].append(row[colnum])
               # print("====== Row No."+str(rownum)+": "+str(cell_list[rownum]))
    return cell_list

#if the cell given row&col is in the area of size title
def is_size_title(current_row,current_col,upper_row,left_col,right_col):
    if current_row==upper_row+1 and current_col>left_col and current_col<right_col:
        #print("&&&&now it is the size row , current row="+str(current_row))
        return 1
    else:
        return False

def parse_gz(file='test.xlsx',by_index=1):
    cell_list=excel_read_everycell(file,by_index)
    nrows=len(cell_list)
    ncols=len(cell_list[0])
    packing_list={}
    packing_list["detail"]=[]
    packing_list['summary']={}
    detail_seg=False
    summary_seg=False
    size_list=[]
    summary_list={}
    summary_line={}
    summary_colour=""
    col_from=0
    col_to=1
    col_carton_qty=2
    col_colour_detail=3
    col_per_carton_pcs=14
    col_subtotal=15
    col_per_carton_gw=16
    col_per_carton_nw=17
    col_length=18
    col_width=19
    col_height=20
    col_summary_size=3
    row_size_head=9

    for rownum in range(nrows-1):
        current_row=cell_list[rownum]
      #  print("==Current row No."+str(rownum)+" : "+str(current_row))

        if str(current_row[0]).upper()=="TOTAL":
            detail_seg=False
        
        if detail_seg:
            #handle the detail carton
            carton_info={}
            carton_info['from']=current_row[col_from]            
            carton_info['to']=current_row[col_to]
            carton_info['carton_qty']=current_row[col_carton_qty]
            carton_info['colour_detail']=current_row[col_colour_detail]
            carton_info['per_carton_pcs']=current_row[col_per_carton_pcs]
            carton_info['subtotal']=current_row[col_subtotal]
            carton_info['per_carton_gw']=current_row[col_per_carton_gw]
            carton_info['per_carton_nw']=current_row[col_per_carton_nw]
            carton_info['length']=current_row[col_length]
            carton_info['width']=current_row[col_width]
            carton_info['height']=current_row[col_height]
            size_qty={}
            for each_size in size_list:
                size_qty[each_size.get('size_name')]=current_row[each_size.get('size_col')]
            #print('=== in row No.'+str(rownum)+' the size_qty is '+str(size_qty))
            carton_info['size_qty']=size_qty
            if carton_info.get('from')!="":
                packing_list['detail'].append(carton_info)
            
        if summary_seg:
            #handle the sumary carton
            #print("this is summary block")
            if str(current_row[0]).upper().find("DESIGN")!=-1:
                size_list=[] #clear the size_list to be used by summary
                for line in range(1,4):
                    if cell_list[rownum+line][0].strip()!="":
                        summary_colour=cell_list[rownum+line][0]
                        summary_list[summary_colour]={}
                for colnum in range(ncols-1):
                    current_cell=current_row[colnum]
                    if current_cell.strip().upper()=="SIZE":
                        col_summary_size=colnum
                    if current_cell.strip().upper()=="TOTAL":
                        col_summary_total=colnum
                    if colnum>col_summary_size and str(current_cell).strip().upper()!="TOTAL":
                        size_info={"size_name":current_cell,"size_col":colnum}
                        size_list.append(size_info)
                        
            #when come to Total Quantity, end summary and add summary_list to packing_list            
            elif current_row[0].strip().upper()=='TOTAL QUANTITY': 
                summary_seg=False
                packing_list['summary']=summary_list
                for col in range(1,ncols-1):
                    if str(current_row[colnum]).strip()!="":
                        packing_list["total_quantity"]=current_row[colnum]

            #in the summary line
            else:
                size_qty={}
                for each_size in size_list:
                    size_qty[each_size.get('size_name')]=current_row[each_size.get('size_col')]
                qty_style={'size_qty':size_qty,'total':current_row[col_summary_total]}
                if str(current_row[col_summary_total]).strip()!="":
                    summary_line=summary_list.get(summary_colour)
                    summary_line[current_row[col_summary_size]]=qty_style
                    summary_list[summary_colour]=summary_line
            
        #not in detail_seg or summary_seg
        for colnum in range(ncols-1):
            current_cell=current_row[colnum]
            #print("######current cell =  " +str(current_cell))
            if is_size_title(current_row=rownum,current_col=colnum,upper_row=row_size_head,left_col=col_colour_detail,right_col=col_per_carton_pcs):
                #print("----current cell =  " +str(current_cell))
                size_info={}
                size_info["size_name"]=current_cell
                size_info["size_col"]=colnum
                size_list.append(size_info)
                #print ("size_list is " + str(size_list))
            if str(current_cell).upper()=="FROM":
                #print("****-current cell =  " +str(current_cell))
                col_from=colnum
                detail_seg=1
            if str(current_cell).upper()=="TO":
                col_to=colnum
            if str(current_cell).upper()=="CTN QTY":
                col_carton_qty=colnum
            if str(current_cell).upper()=="DESIGN COLOR":
                col_colour_detail=colnum
            if str(current_cell).upper().find("ASSORTMENT") !=-1:
                row_size_head=rownum
                #print("@@@@ the row_size_head is set to "+str(row_size_head))
            if str(current_cell).upper()=="PER CARTON" and str(cell_list[rownum+1][colnum])=="PCS":
                col_per_carton_pcs=colnum
            if str(current_cell).upper()=="PER CARTON" and str(cell_list[rownum+1][colnum])=="G.W.":
                col_per_carton_gw=colnum
            if str(current_cell).upper()=="PER CARTON" and str(cell_list[rownum+1][colnum])=="N.W.":
                col_per_carton_n=colnum
            if str(current_cell).upper()=="SUB TOTAL" and str(cell_list[rownum+1][colnum])=="PCS":
                col_subtotal=colnum
            if str(current_cell).upper()=="L" and str(cell_list[rownum-1][colnum])=="CARTON MEASUREMENT":
                col_length=colnum
            if str(current_cell).upper()=="W" and str(cell_list[rownum-1][colnum-1])=="CARTON MEASUREMENT":
                col_width=colnum
            if str(current_cell).upper()=="H" and str(cell_list[rownum-1][colnum-2])=="CARTON MEASUREMENT":
                col_length=colnum
            if str(current_cell).upper().find("SUMMARY")!=-1:
                row_summary=rownum
                summary_seg=1
            if str(current_cell).upper().find("PACKAGE")!=-1:
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["total_carton"]=current_row[col]
                        break
                        
            if str(current_cell).upper().find("GROSS")!=-1:
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["total_gw"]=current_row[col]
                        break
                
            if str(current_cell).upper().find("NET")!=-1:
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["total_nw"]=current_row[col]
                        break
            if str(current_cell).upper().find("VOLUME")!=-1:
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["total_volume"]=current_row[col]
                        break

            if str(current_cell).strip()=="Style No.":
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["style_no"]=str(current_row[col]).strip()
                        break
            if str(current_cell).strip()=="Description":
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["style_escription"]=str(current_row[col]).strip()
                        break
            if str(current_cell).strip()=="Invoice No.":
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["invoice_no"]=str(current_row[col]).strip()
                        break
            if str(current_cell).strip()=="Date":
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["date"]=xlrd.xldate.xldate_as_datetime(current_row[col],0)
                        break
            if str(current_cell).strip()=="Order No.":
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["order_no"]=str(current_row[col]).strip()
                        break

    return packing_list

def print_dict(data):
    if isinstance(data,dict):
        for item in data:
            print("%s : "%item,end="")
            print_dict(data[item])
    elif isinstance(data,list):
        print("Below is list")
        for item in data:
            print_dict(item)
    else:
        print(data)
        


def parse_gz_invoice(file='test.xlsx',by_index=0):
    cell_list=excel_read_everycell(file,by_index)
    nrows=len(cell_list)
    ncols=len(cell_list[0])
    invoice={}
    invoice["detail"]=[]
    detail_seg=False
    col_order_no=1
    col_qty=2
    col_price=4
    col_total=5
    
    for rownum in range(1,nrows-1):
        current_row=cell_list[rownum]
        #print("==Current row No."+str(rownum)+" : "+str(current_row))
        if not detail_seg:
            print ("not detail_seg")
            for colnum in range(ncols-1):
                current_cell=str(current_row[colnum]).upper()
                if current_cell=="INVOICE NO:":
                    invoice["InvNo"]=current_row[colnum+1]
                if current_cell=="DATE:":
                   # invoice["InvDate"]=xlrd.xldate_as_tuple(current_row[colnum+1],0)
                    invoice["InvDate"]=xlrd.xldate.xldate_as_datetime(current_row[colnum+1],0)
                if current_cell=="BENEFICIARY:":
                    invoice["Beneficiary"]=current_row[colnum+1]
                if current_cell=="ADDRESS:":
                    invoice["Address"]=current_row[colnum+1]
                if current_cell=="ACCOUNT NO.:":
                    invoice["AccNo"]=current_row[colnum+1]
                if current_cell=="NAME OF BANK:":
                    invoice["BankName"]=current_row[colnum+1]
                if current_cell=="SWIFT CODE:":
                    invoice["SwiftCode"]=current_row[colnum+1]
                if current_cell=="TEL:":
                    invoice["Tel"]=current_row[colnum+1]
                if current_cell=="FAX:":
                    invoice["Fax"]=current_row[colnum+1]
                if current_cell=="POST COD:":
                    invoice["PostCode"]=current_row[colnum+1]
                if str(current_row[colnum]).find("ORDER NUMBER") != -1:
                    col_order_no=colnum
                if str(current_row[colnum]).find("QTY")!= -1 and colnum!=0:
                    col_qty=colnum
                if str(current_row[colnum]).find("UNIT PRICE")!= -1:
                    col_price=colnum
                if str(current_row[colnum]).find("TOTAL AMOUNT")!= -1:
                    col_total=colnum
                
        if current_row[0]=="TOTAL":
            invoice["total_qty"]=current_row[col_qty]
            invoice["total_amount"]=current_row[col_total]
            detail_seg=False
            print("*****set seg - get total"+str(detail_seg))
            print(detail_seg)

        if str(current_row[col_order_no]).find("TIS1")!=-1:
            current_order_sub={}
            order_info=str(current_row[col_order_no]).split('/')
            current_order_sub["style"]=order_info[1]
            current_order_sub["qty"]=current_row[col_qty]
            current_order_sub["price"]=current_row[col_price]
            current_order_sub["amount"]=current_row[col_total]
            current_order={}
            current_order[order_info[0]]=current_order_sub
            invoice["detail"].append(current_order)
    return invoice
            
            
    
    

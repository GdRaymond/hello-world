import xdrlib,sys
import xlrd
import re
from read_excel import excel_read_everycell
from xls_log import get_xls_logger
from save_db import save



"""
The diction of packing_lis is as below:
{'detail':[{'form':,'to:','carton_qty':,'colour_detail':,'per_carton_pcs':,'per_carton_gw':,'per_carton_nw':,
            'subtotal':,'length':,'width':,'height':,'size_qty':{-xs:, ...}}],
 'summary':{-COBALT BLUE:{'Order Qty':{'total':, 'size_qty':{-xs:,....}},
                          'Actual Qty':{'total':, 'size_qty':{-xs:,....}},
                          'Balance':{'total':, 'size_qty':{-xs:,....}},
                          'Ratio':{'total':, 'size_qty':{-xs:,....}}
                          }
                          ...
            }
 'total_quantity':,'total_carton':,'total_gw':,'total_nw':,'total_volume':,'style_no':,'style_description':,
 'invoice_no':,'date':,'order_no':
}
"""

logger=get_xls_logger()


#if the cell given row&col is in the area of size title
def is_size_title(current_row,current_col,row_size_head,left_col,right_col):
    if current_row==upper_row+1 and current_col>left_col and current_col<right_col:
        #print("&&&&now it is the size row , current row="+str(current_row))
        return 1
    else:
        return False
    
def str_contain(string,words,all_any='all'):
    result=[str(word).strip().upper() in str(string).strip().upper() for word in words]
    if all_any=='all': #all words must all exist
        return not ( False in result)
    else: #any word exist
        return True in result

def parse_packing_list(cell_list=[],file='test.xlsx',by_name="RM500BT(TIS16-SO3466)"):
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
    col_carton_qty=1
    col_colour_detail=2
    col_per_carton_pcs=14
    col_subtotal=15
    col_per_carton_gw=16
    col_per_carton_nw=17
    col_length=18
    col_width=19
    col_height=20
    col_summary_size=2
    col_summary_colour=1
    row_size_head=15
    row_summary=50

    #print(cell_list)

    """iterate the rows to find the row No. for size head and summary for finding the size title, because in
       Tanhoo document there will be 2 row size title for trousers.
    """
    for rownum in range(nrows):
        current_row=cell_list[rownum]
        if str_contain(str(current_row[0]).strip().upper(),['CARTON','NO.']) and ('G.W.' in [str(item).strip().upper() for item in current_row]) : 
            row_size_head=rownum
            #logger.info(' -%s-%s row_size_head=%s'%(file,by_name,row_size_head))
            before_packs=True
            for colnum in range(ncols):
                current_cell=current_row[colnum]
                logger.debug('--current_cell, colnum[%s]=%s'%(colnum,current_cell))
                if str_contain(str(current_row[0]).strip().upper(),['CTN','NO.']):
                    col_from=colnum
                elif str(current_cell).strip().upper()=="CARTONS":
                    col_carton_qty=colnum
                elif str(current_cell).upper()=="COLOUR":
                    col_colour_detail=colnum
                #below find the 1st size head to get the size_list
                elif colnum>col_colour_detail and before_packs and (str(cell_list[rownum+1][colnum]).strip().upper() not in ["","PACKS"]):
                    size_info={}
                    if isinstance(cell_list[rownum+1][colnum],float) and float(int(cell_list[rownum+1][colnum]))==cell_list[rownum+1][colnum]:
                        size_info["size_name"]=str(int(cell_list[rownum+1][colnum])).strip()#for the size like 72.0,16.0,20.0,convert to 72,16,20
                    else:
                        size_info["size_name"]=str(cell_list[rownum+1][colnum]).strip()#keep 5R, 5.5 
                    size_info["size_col"]=colnum
                    size_list.append(size_info)

                elif str(current_cell).upper()=="PACKS":
                    col_per_carton_pcs=colnum
                    before_packs=False
                elif str_contain(str(current_cell),['TOTAL','PACKS']):
                    col_subtotal=colnum
                elif str(current_cell).strip().upper()=="G.W.":
                    col_per_carton_gw=colnum
                elif str(current_cell).upper()=="N.W.":
                    col_per_carton_nw=colnum
                elif str(current_cell).upper()=="MEASUREMENT":
                    col_length=colnum
                    col_width=colnum+1
                    col_height=colnum+2
            logger.debug('---col_carton=%s,col_subtotal=%s,col_per_carton_gw=%s, sizt_list=%s'%(col_carton_qty,col_subtotal,col_per_carton_gw,size_list))
            break
    

    detail_current_colour='' #in the case the colour is omit to follow previous line
    summary_actual_total=0 #some summary has 2 actual shipment, like 1st sea , 2nd sea. We only get the value of last shipment. for the total, we need
                           #to check if it is the last shipment, when we encount the 'BALANCE', it means what in actual qty is the last one, so we can
                           #get this value and accumulate to the summary_actual_total and also write back to summary
    for rownum in range(nrows):
        current_row=cell_list[rownum]
        #print("==Current row No."+str(rownum)+" : "+str(current_row))

        #find row_size_head to enter the detial_seg
        if rownum==row_size_head:
            detail_seg=True
            summary_seg=False
            continue
        
        if set(['COLOR','SIZE','TTL']).issubset(set([str(cell).strip().upper() for cell in current_row])):
            detail_seg=False
            summary_seg=True
        
        #in detail_seg
        if detail_seg:
            #check if 2nd row size head
            if str_contain(str(current_row[0]).strip().upper(),['CARTON','NO.']) and ('G.W.' in [str(item).strip().upper() for item in current_row]): #in the case of CNT NO. omited:
                size_list=[]
                for colnum in range(col_colour_detail+1,col_subtotal):
                    current_cell=current_row[colnum]
                    if str(current_cell).strip()=="":
                        continue
                    size_info={}
                    if isinstance(current_cell,float) and float(int(current_cell))==current_cell:
                        size_info["size_name"]=str(int(current_cell)).strip()#for the size like 72.0,16.0,20.0,convert to 72,16,20
                    else:
                        size_info["size_name"]=str(current_cell).strip() #keep 5R, 5.5 
                    size_info["size_col"]=colnum
                    size_list.append(size_info)

            elif str(current_row[0]).strip().upper()=="TOTAL":
                detail_seg=False
                packing_list["total_carton"]=current_row[col_carton_qty]
                packing_list["total_quantity"]=current_row[col_subtotal]
                packing_list["total_volume"]=current_row[col_height]

            elif str(current_row[0]).strip().upper()=="":
                continue
            #handle the detail carton
            else: 
                    if str(current_row[col_from]).strip()=='':
                        continue
                    carton_info={}
                    from_to=str(current_row[col_from]).split('-')
                    
                    try:
                        carton_info['from']=int(float(from_to[0]))
                    except Exception as err:
                        logger.info("--%s-%s from_to=%s, maybe at the line of detail total, error=%s, current_row=%s"%(file,by_name,from_to,err,current_row))
                        continue
                    if len(from_to)==2: #normal more than 1 carton
                        carton_info['to']=int(float(from_to[1]))
                    elif len(from_to)==1: #only 1 carton
                        carton_info['to']=int(float(from_to[0]))
                    else: #sometimes, it is 1--2
                        for index in range(1,len(from_to)):
                            if from_to[index]!='':
                                carton_info['to']=int(float(from_to[index]))
                                break
                    try:
                        carton_info['carton_qty']=int(float(current_row[col_carton_qty]))
                    except Exception as err:
                        logger.error("--%s-%s current_row[%s]=%s,from_to=%s,  error=%s"%(file,by_name,col_carton_qty,\
                                                                                         current_row[col_carton_qty],from_to,err))
                    try:
                        if current_row[col_colour_detail].strip() not in ['',"‘’"]:
                            detail_current_colour=current_row[col_colour_detail].strip().upper()
                            carton_info['colour_detail']=current_row[col_colour_detail].strip().upper()
                            #logger.debug('--%s-%s-current_row[%s]=%s'%(file,by_name,col_colour_detail,current_row[col_colour_detail]))
                        else:
                            carton_info['colour_detail']=detail_current_colour
                            #logger.debug('--%s-%s-detail_current_colour=%s'%(file,by_name,detail_current_colour))
                        
                    except Exception as err:
                        logger.error('--%s-%s col_colour_detail=%s current_cell=%s error=%s'\
                                        %(file,by_name,current_row[col_colour_detail],err))
                    carton_info['per_carton_pcs']=current_row[col_per_carton_pcs]
                    carton_info['subtotal']=current_row[col_subtotal]
                    carton_info['per_carton_gw']=current_row[col_per_carton_gw]
                    carton_info['per_carton_nw']=current_row[col_per_carton_nw]
                    carton_info['length']=current_row[col_length]                   
                    carton_info['length']=current_row[col_width]
                    carton_info['length']=current_row[col_height]

                    size_qty={}
                    for each_size in size_list:
                        if str(current_row[each_size.get('size_col')]).strip()=="":
                            qty=0
                        else:
                            qty=int(float(str(current_row[each_size.get('size_col')]).strip()))
                        size_qty[each_size.get('size_name')]=qty
                    #print('=== in row No.'+str(rownum)+' the size_qty is '+str(size_qty))
                    carton_info['size_qty']=size_qty
                    if carton_info.get('from')!="":
                        packing_list['detail'].append(carton_info)
            
            
        if summary_seg:
            #handle the sumary carton
            #below identify the coloumn for summary of each colour
            #Colour Design Size quantity
            if set(['COLOR','SIZE','TTL']).issubset(set([str(cell).strip().upper() for cell in current_row])):
                logger.debug('-%s-%s-Start to parse summary'%(file,by_name))
                size_list=[] #clear the size_list to be used by summary
                for colnum in range(ncols):
                    current_cell=current_row[colnum]
                    if str(current_cell).strip().upper()=="COLOR":
                        col_summary_colour=colnum
                    if str(current_cell).strip().upper()=="SIZE":
                        col_summary_size=colnum
                    if str(current_cell).strip().upper()=="TTL":
                        col_summary_total=colnum
                    if colnum>col_summary_size and str(current_cell).strip().upper()not in ["TTL",""]:
                        size_info={}
                        if isinstance(current_cell,float) and float(int(current_cell))==current_cell:
                            size_info["size_name"]=str(int(current_cell)).strip() #for the size like 72.0,16.0,20.0,convert to 72,16,20
                        else:
                            size_info["size_name"]=str(current_cell).strip() #keep 5R , 5.5

                        size_info["size_col"]=colnum
                        size_list.append(size_info)
                logger.debug('--%s-%s-the size list is%s '%(file,by_name,size_list))
                for line in range(1,4):
                    if cell_list[rownum+line][col_summary_colour].strip()!="":
                        summary_colour=cell_list[rownum+line][col_summary_colour].strip().upper()
                        if summary_list.get(summary_colour) is None:
                            summary_list[summary_colour]={}
                            summary_actual_total=0
                        break
                        
            #when come to %, end summary and add summary_list to packing_list            
            elif '%' in [str(cell).strip().upper() for cell in current_row] :
                summary_seg=False
                packing_list['summary']=summary_list

            #in the summary line for each colour
            else:
                size_qty={}
                for each_size in size_list:
                    if str(current_row[each_size.get('size_col')]).strip()=="":
                        size_qty[each_size.get('size_name')]=0
                    else:
                        size_qty[each_size.get('size_name')]=current_row[each_size.get('size_col')]
                    
                qty_style={'size_qty':size_qty,'total':current_row[col_summary_total]}
                logger.debug('--%s-%s-the qty_size is%s '%(file,by_name,qty_style))

                if str(current_row[col_summary_total]).strip()!="":
                    summary_line=summary_list.get(summary_colour)
                    """
                     if the colour is none , like LM16ZF214-SO2617 Hanger, the summary_line is none
                    """
                    if summary_line is None:
                        logger.error('---%s - %s - No colour in summary line  '%(file,by_name))
                        continue 
                    """
                      fix the type of qty line to Order Qty, Actual Qty,Balance,Ratio as the name will change,
                      for example, the 1st Sea Qty should be Actual Qty
                    """
                    qty_type=current_row[col_summary_size]
                    if str(qty_type).strip().upper()=='ORDER QTY' or str(qty_type).strip()=='' \
                       and cell_list[rownum-1][col_summary_size].strip().upper()=='SIZE':
                        qty_type='Order Qty'
                    elif str(qty_type).strip().upper()=='BALANCE' or str(qty_type).strip()=='' \
                        and cell_list[rownum-3][col_summary_size].strip().upper()=='SIZE':
                        qty_type='Balance'
                        logger.debug('--%s-%s start balance the  summary_line is %s'%(file,by_name,summary_line))
                        summary_actual_total+=summary_line.get('Actual Qty').get('total') #when comes to balance, the last ship qty is actual qty,then sum
                        summary_line['Actual Qty']['total']=summary_actual_total
                        """
                         As Jinfeng's document has no balance qty for each size and ratio, we have to calculate by Actual qty and order qty
                        """
                        size_qty={}
                        size_ratio={}
                        for each_size in size_list:
                            size_name=each_size.get('size_name')
                            logger.debug('-%s-%s - calculate balance, actual size[%s}=%spcs, order size[%s]=%spcs'\
                                            %(file,by_name,size_name,summary_line.get('Actual Qty').get('size_qty').get(size_name),\
                                              size_name,summary_line.get('Order Qty').get('size_qty').get(size_name)))
                            if (summary_line.get('Actual Qty').get('size_qty').get(size_name)==0) or\
                               (summary_line.get('Order Qty').get('size_qty').get(size_name)==0):
                                continue
                            size_qty[size_name]=summary_line.get('Actual Qty').get('size_qty').get(size_name)\
                                                 -summary_line.get('Order Qty').get('size_qty').get(size_name)
                            try:
                                size_ratio[size_name]=size_qty.get(size_name)/summary_line.get('Order Qty').get('size_qty').get(size_name)
                            except Exception as err:
                                 logger.error('---%s - %s - Ratio caculation wrong at size=%s error =%s '%(file,by_name,size_name,err))

                        qty_style={'size_qty':size_qty,'total':current_row[col_summary_total]}
                        summary_line['Balance']=qty_style
                        qty_style={'size_qty':size_ratio,'total':0}
                        summary_line['Ratio']=qty_style
                        logger.debug('--%s-%s the  summary_lin is =%s'%(file,by_name,summary_line))
                    elif str(qty_type).strip().upper()=='RATIO' or str(qty_type).strip()=='' \
                        and cell_list[rownum-4][col_summary_size].strip().upper()=='SIZE':
                        qty_type='Ratio'
                    elif str_contain(qty_type,['SEA','ACTUAL','SHIPMENT','SHIP'],all_any='any') or str(qty_type).strip()=='' \
                        and cell_list[rownum-2][col_summary_size].strip().upper()=='SIZE':
                        qty_type='Actual Qty'
                    """in some packinglist, like QPSTO with many size, the summary is split to 2 line
                       for the 1st line, assign the qty_style directly, for the 2nd line need to combine the size and sum the total
                    """
                    logger.debug("--%s-%s summary_actual_total=%s, summary_line['%s']=%s"%(file,by_name,summary_actual_total,qty_type,summary_line.get(qty_type)))
                    if summary_actual_total==0 or summary_line.get(qty_type) is None: #1st line for this colour
                        if qty_type in ['Order Qty','Actual Qty']:
                            summary_line[qty_type]=qty_style   
                    elif str(qty_type).strip().upper()=='BALANCE': #for Jinfeng, the Balance seg has done all work, no need anything
                        continue
                    else:#2nd line, never happen in Jinfeng
                        combined_size_qty=summary_line.get(qty_type).get('size_qty')
                        logger.debug('--%s-%s before combining the size_qty, the 1st=%s , 2nd=%s'%(file,by_name,combined_size_qty,qty_style.get('size_qty')))
                        combined_size_qty.update(qty_style.get('size_qty'))
                        logger.debug('--%s-%s after combining the size_qty  =%s'%(file,by_name,combined_size_qty))
                        """
                        sum_total=summary_line.get(qty_type).get('total')
                        sum_total+=qty_style.get('total')
                        qty_style['total']=summary_actual_total
                        logger.debug('--%s-%s the  total is sum =%s'%(file,by_name,summary_actual_total))
                        """
                        qty_style['size_qty']=combined_size_qty
                        logger.debug('--%s-%s the new qty_style =%s'%(file,by_name,qty_style))
                        if qty_type in ['Order Qty','Actual Qty']:
                            summary_line[qty_type]=qty_style   
                    summary_list[summary_colour]=summary_line
                    logger.debug('--%s-%s-the summary_list is%s '%(file,by_name,summary_list))
            
        #not in detail_seg or summary_seg
        for colnum in range(ncols):
            current_cell=current_row[colnum]
            if str(current_cell).strip().upper().startswith("G.W.:"):
                packing_list["total_gw"]=str(current_cell).strip().upper()[5:len(str(current_cell).strip())-3]
                
            elif str(current_cell).strip().upper().startswith("N.W.:"):
                packing_list["total_gw"]=str(current_cell).strip().upper()[5:len(str(current_cell).strip())-3]
                
            elif (re.match(r'(.*)(TIS1\d-SO\d{4})(.*)',str(current_cell).strip())) is not None:
                match=re.match(r'(.*)(TIS1\d-SO\d{4})(.*)',str(current_cell).strip())
                packing_list["style_no"]=str(match.group(1)).replace('/','').replace('\\','').strip()
                packing_list["style_description"]=str(match.group(3)).strip()
                packing_list["order_no"]=str(match.group(2)).strip()
                
            elif (re.match(r'(INVOICE NO.)(.*)',str(current_cell).strip())) is not None:
                match=re.match(r'(INVOICE NO.)(.*)',str(current_cell).strip())
                packing_list["invoice_no"]=match.group(2)
                
            elif (re.match(r'(DATE:)(.*)',str(current_cell).strip())) is not None:
                match=re.match(r'(DATE:)(.*)',str(current_cell).strip())
                if match.group(2)=='':
                    if current_row[colnum+1]!='':
                        packing_list["date"]=current_row[colnum+1].strip()
                else:
                    packing_list["date"]=match.group(2).strip()

    print(packing_list)
    return packing_list

"""from detail to calculate the summary
   return: summary_by_cal={'-Cobalt blue':{'total':, 'size_qty'{'-2xs':,...}},
                           ...
                          }
"""
def calculate_detail(packing_list={},file='test.xlsx',by_name="RM500BT(TIS16-SO3466)"):
    logger.debug('======================================')
    logger.debug('Start calculate_detail and validate detail')
    detail_correct = True
    summary_by_cal={}  #init the result dict
    list_carton_no_from=[]
    list_carton_no_to=[]
    carton_total_by_cal=0

    logger.debug('-Start  validate carton No. in each line')
    all_correct=True
    for line in packing_list.get('detail'):
        #validate the carton No. and quantity
        carton_no_from=line.get('from')
        carton_no_to=line.get('to')
        carton_line_total=line.get('carton_qty')
        if  carton_line_total!=(carton_no_to-carton_no_from+1):
            logger.error('--%s - %s - Carton No. wrong at line from:%s to:%s total:%s '%(file,by_name,carton_no_from,carton_no_to,carton_line_total))
            all_correct=False       
        
        #put the carton no to list and accumulate the carton qty
        list_carton_no_from.append(carton_no_from)
        list_carton_no_to.append(carton_no_to)
        carton_total_by_cal+=carton_line_total
        
        colour=line.get('colour_detail')
        temp_dict=summary_by_cal.get(colour)
        if temp_dict is None:
            temp_dict={}
        temp_size_qty=temp_dict.get('size_qty')
        if temp_size_qty is None:
            temp_size_qty={}
        for size in line.get('size_qty'):
            #print("====in colour=%s, size=%s , carton qty=%s, per carton=%s "%(colour,size,line.get('to')-line.get('from')+1,line.get('size_qty').get(size)))
            base_qty=temp_size_qty.get(size)
            if base_qty is None:
                base_qty=0
            temp_size_qty[size]=base_qty+ carton_line_total*line.get('size_qty').get(size)
        temp_dict['size_qty']=temp_size_qty
        temp_dict['total']=0
        summary_by_cal[colour]=temp_dict
    if all_correct:
        logger.debug('--Correct of carton No. in each line ')
    else:
        detail_correct=False


    # validate the consistency of no. and quantity by iterate\ing the carton no list
    all_correct=True
    logger.debug('-Start validate the consistency of carton no.')
    for no in range(0,len(list_carton_no_from)-1):
        if list_carton_no_from[no+1]!=list_carton_no_to[no]+1:
            all_correct=False
            logger.error('--%s - %s - Carton No. not consistant at line to:%s next from:%s  '\
                         %(file,by_name,list_carton_no_to[no],list_carton_no_from[no+1]))
    if all_correct:
        logger.debug('--Correct of the consistency of carton No. ')
    else:
        detail_correct=False

    if detail_correct:
        logger.debug('-Correct detail carton No.')
        
    # validate the carton quantity 
    logger.debug('========================================')
    logger.debug('Start validate the total carton quantity')
    if carton_total_by_cal==packing_list.get('total_carton'):
        logger.debug('-Correct of he total carton quantity.')
    else:
        logger.error('-%s-%s-Wrong of the total carton, in sheet - %s, but by calculation - %s'\
                     %(file,by_name,packing_list.get('total_carton'),carton_total_by_cal))

    for item in summary_by_cal:
        temp_dict=summary_by_cal.get(item)
        temp_total=temp_dict.get('total')
        for size in temp_dict.get('size_qty'):
            temp_total+=temp_dict.get('size_qty').get(size)
        temp_dict['total']=temp_total
        summary_by_cal[item]=temp_dict
    return summary_by_cal

#validate the summary with detail
def validate_summary(packing_list={},file='test.xlsx',by_name="RM500BT(TIS16-SO3466)"):
    #calculate detail
    detail=calculate_detail(packing_list=packing_list,file=file,by_name=by_name)

    all_summary_correct=True
    logger.debug('======================================')
    logger.debug('Start validate summary')
    #validate sumary on ratio and balance
    value_summary=packing_list.get('summary')
    for colour in value_summary:     
        value_colour=value_summary.get(colour)
        
        #validate the ratio, find the ration greater than 5% and balance greater than 10pcs
        logger.debug('-Start validate the ratio for colour:%s , checking if ratio> %%5 and balance>10pcs'%(colour))
        all_correct=True
        try:            
            ratio=value_colour.get('Ratio').get('size_qty')
            balance=value_colour.get('Balance').get('size_qty')
        except Exception as err:
            logger.error('--%s - %s - Ratio or balance can not find, error--%s'\
                             %(file,by_name,err))
            continue
        balance_by_cal=0
        for size in ratio:
            ratio_value=ratio.get(size)
            balance_value=balance.get(size)
            balance_by_cal+=balance_value
            #logger.debug('size:%s--ratio:%s'%(size,ratio_value))
            if str(ratio_value).strip()!="" and ratio_value>0.05 \
               and str(balance_value).strip()!="" and balance_value>9:
                logger.error('--%s - %s - Ratio warning at colour:%s size:%s - ratio: %s, balance:%s'\
                             %(file,by_name,colour,size,ratio_value,balance_value))
                all_correct=False
        if balance_by_cal!=value_colour.get('Balance').get('total'):
            logger.error('--%s-%s - Balance total is not correct at colour: %s, in doc=%s, cal=%s'\
                         %(file,by_name,colour,value_colour.get('Balance').get('total'),balance_by_cal))
            all_correct=False
        if all_correct:
            logger.debug('--The total balance is correct and  ratio for colour: %s is ok without exceeding 5%% and balance greater than 10pcs'%colour)
        else:
            all_summary_correct=False

    #validate summary the quantity  comparing with detail quantity : 1- colour from summary to detail
    logger.debug('-Start validate the quantity for colour:, summary-> detail')
    summary_correct=True
    summary_total_qty=0
    for colour in value_summary:
        logger.debug('--Start validate the quantity for colour:%s, summary-> detail'%colour)
        all_correct=True
        value_colour=value_summary.get(colour)

        """
          if the Actual Qty name is varied, like LM17ZF202-ob-sox, the name is BY SEA, then error here
        if by_name=='RM301EWS(TIS15-SO3127)':
            logger.error('RM301EWS(TIS15-SO3127)--value_colour %s'%value_colour)
        """
        

        try:   
            actual_qty=value_colour.get('Actual Qty')
            summary_total_qty+=actual_qty.get('total')
            logger.debug('---%s - %s - summary_total=: %s'%(file,by_name,summary_total_qty))
            actual_qty_size_qty=actual_qty.get('size_qty')
        except Exception as e:
            logger.error('---%s - %s - Can not find Actual Qty name: %s'%(file,by_name,e))
            continue
        """
          if the colour name in detail not consistant with summary, error will occur
          for example, in LM16ZF214-SO3469, coulur in detail is INK, in summary is INK NAVY
        """
        if detail.get(colour) is None:
            logger.error('---%s - %s - Can not find colour:%s in detail '\
                             %(file,by_name,colour))
            continue
        detail_size_qty=detail.get(colour).get('size_qty')
        #validate the actural qty , 1.1 - size from summary to detail
        total_by_cal=0
        for size in actual_qty_size_qty:
            actual_qty_size_qty_value=actual_qty_size_qty.get(size)
            if str(actual_qty_size_qty_value).strip()!="":
               # logger.debug('---%s - %s - Qty data at colour:%s size:%s - size qty: %s, temp total:%s'\
                           #  %(file,by_name,colour,size,actual_qty_size_qty_value,total_by_cal))
                total_by_cal+=actual_qty_size_qty_value

                if actual_qty_size_qty_value!=detail_size_qty.get(size):
                    if ((actual_qty_size_qty_value is None) and detail_size_qty.get(size)==0) or \
                       (actual_qty_size_qty_value==0 and (detail_size_qty.get(size) is None)):
                        continue
                    logger.error('---%s - %s - Qty warning at colour:%s size:%s - summary: %s, detail:%s'\
                             %(file,by_name,colour,size,actual_qty_size_qty_value,detail_size_qty.get(size)))
                    all_correct=False
                    summary_correct=False
        
        #validate the actual qty total, comparing with calculation
        if actual_qty.get('total')!=total_by_cal:
            logger.error('---%s - %s - Actual qty total warning at colour:%s  - in doc: %s, by calculation:%s'\
                        %(file,by_name,colour,actual_qty.get('total'),total_by_cal))

        
        #validate the actural qty , 1.2 - size from detail to summary
        for size in detail_size_qty:
            actual_qty_size_qty_value=actual_qty_size_qty.get(size)
            if str(actual_qty_size_qty_value).strip()!="":
                if actual_qty_size_qty_value!=detail_size_qty.get(size):
                    if ((actual_qty_size_qty_value is None) and detail_size_qty.get(size)==0) or \
                       (actual_qty_size_qty_value==0 and (detail_size_qty.get(size) is None)):
                        continue
                    logger.error('---%s - %s - Qty warning at colour:%s size:%s - detail: %s, summary:%s'\
                             %(file,by_name,colour,size,detail_size_qty.get(size),actual_qty_size_qty_value))
                    all_correct=False
                    summary_correct=False

        if all_correct:
             logger.debug('---Correct: the quantity for colour:%s, summary-> detail '%colour)
    if summary_correct:
        logger.debug('--All Correct: the quantity colour summary-> detail ')
    else:
        all_summary_correct=False
        
        

    #validate summary the quantity  comparing with detail quantity : 2- colour from detail to summary
    logger.debug('-Start validate the quantity for colour, detail-> summary')
    summary_correct=True
    for colour in detail:
        logger.debug('--Start validate the quantity for colour:%s, detail-> summary'%colour)
        all_correct=True
        value_colour=value_summary.get(colour)
        
        """
          if the colour name in detail not consistant with summary, error will occur
          for example, in LM16ZF214-SO3469, coulur in detail is INK, in summary is INK NAVY
        """
        if value_colour is None:
            logger.error('---%s - %s - Can not find colour:%s in summary '\
                             %(file,by_name,colour))
            continue

        try:
            actual_qty=value_colour.get('Actual Qty')
            actual_qty_size_qty=actual_qty.get('size_qty')
        except Exception as e:
            logger.error('---%s-%s-Can not find Actural Qty name - %s'%(file,by_name,e))
            continue
        detail_size_qty=detail.get(colour).get('size_qty')
        #validate the actural qty , 2.1 - size from summary to detail
        for size in actual_qty_size_qty:
            actual_qty_size_qty_value=actual_qty_size_qty.get(size)
            if str(actual_qty_size_qty_value).strip()!="":
                if actual_qty_size_qty_value!=detail_size_qty.get(size):
                    if ((actual_qty_size_qty_value is None) and detail_size_qty.get(size)==0) or \
                       (actual_qty_size_qty_value==0 and (detail_size_qty.get(size) is None)):
                        continue
                    logger.error('---%s - %s - Qty warning at colour:%s size:%s - summary: %s, detail:%s'\
                             %(file,by_name,colour,size,actual_qty_size_qty_value,detail_size_qty.get(size)))
                    all_correct=False
                    summary_correct=False
        
        #validate the actural qty , 2.2 - size from detail to summary
        for size in detail_size_qty:
            actual_qty_size_qty_value=actual_qty_size_qty.get(size)
            if str(actual_qty_size_qty_value).strip()!="":
                if actual_qty_size_qty_value!=detail_size_qty.get(size):
                    if ((actual_qty_size_qty_value is None) and detail_size_qty.get(size)==0) or \
                       (actual_qty_size_qty_value==0 and (detail_size_qty.get(size) is None)):
                        continue
                    logger.error('---%s - %s - Qty warning at colour:%s size:%s - detail: %s, summary:%s'\
                             %(file,by_name,colour,size,detail_size_qty.get(size),actual_qty_size_qty_value))
                    all_correct=False
                    summary_correct=False

        if all_correct:
             logger.debug('---Correct: the quantity for colour:%s, detail-> summary '%colour)
    if summary_correct:
        logger.debug('--All Correct: the quantity colour detail-> summary ')
    else:
        all_summary_correct=False
        
    #validate total quantity in doc comparing with sum of each colour in summary
    logger.debug('-Start validate total quantity ')
    if packing_list.get('total_quantity')!=summary_total_qty:
        logger.error('--%s - %s -Validate total quantity Wrong: in doc:%s - sum  in summary %s '\
                     %(file,by_name,packing_list.get('total_quantity'),summary_total_qty))
        all_summary_correct=False

    else:    
        logger.debug('--Validate total quantity Correct')

    if all_summary_correct:
        logger.debug('-All correct of summary')

    return all_summary_correct

"""

"""
def validate_packinglist_by_sheet(cell_list=[],filename='',sheetname='',save_db=False):
    packing_list=parse_packing_list(cell_list=cell_list,file=filename,by_name=sheetname)
    result=validate_summary(packing_list=packing_list,file=filename,by_name=sheetname)
    if save_db==True:
        count=save(packing_list=packing_list,fty='JF',file=filename,by_name=sheetname)
    return result
        


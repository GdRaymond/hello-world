import xdrlib,sys
import xlrd
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
def is_size_title(current_row,current_col,upper_row,left_col,right_col):
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


    detail_current_colour='' #in the case the colour is omit to follow previous line
    summary_actual_total=0 #some summary has 2 actual shipment, like 1st sea , 2nd sea. We only get the value of last shipment. for the total, we need
                           #to check if it is the last shipment, when we encount the 'BALANCE', it means what in actual qty is the last one, so we can
                           #get this value and accumulate to the summary_actual_total and also write back to summary

    #print(cell_list)
    for rownum in range(nrows):
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
            try:
                if current_row[col_colour_detail].strip()!='':
                    detail_current_colour=current_row[col_colour_detail].strip().upper()
                    carton_info['colour_detail']=current_row[col_colour_detail].strip().upper()
                    #logger.debug('--%s-%s-current_row[%s]=%s'%(file,by_name,col_colour_detail,current_row[col_colour_detail]))
                else:
                    carton_info['colour_detail']=detail_current_colour
                    #logger.debug('--%s-%s-detail_current_colour=%s'%(file,by_name,detail_current_colour))
                        
            except Exception as err:
                logger.error('--%s-%s col_colour_detail=%s current_cell=%s error=%s'\
                                %(file,by_name,current_row[col_colour_detail],current_cell,err))
            carton_info['per_carton_pcs']=current_row[col_per_carton_pcs]
            carton_info['subtotal']=current_row[col_subtotal]
            carton_info['per_carton_gw']=current_row[col_per_carton_gw]
            carton_info['per_carton_nw']=current_row[col_per_carton_nw]
            carton_info['length']=current_row[col_length]
            carton_info['width']=current_row[col_width]
            carton_info['height']=current_row[col_height]
            #logger.error('@@@@@@@the current_row=%s, col_height=%s'%(current_row,col_height))
            
            size_qty={}
            for each_size in size_list:
                if str(current_row[each_size.get('size_col')]).strip()=="":
                    qty=0
                else:
                    qty=current_row[each_size.get('size_col')]
                size_qty[each_size.get('size_name')]=qty
            #print('=== in row No.'+str(rownum)+' the size_qty is '+str(size_qty))
            carton_info['size_qty']=size_qty
            if carton_info.get('from')!="":
                packing_list['detail'].append(carton_info)
            
        if summary_seg:
            #handle the sumary carton
            #below identify the coloumn for summary of each colour
            #Colour Design Size quantity
            if str(current_row[0]).upper().find("DESIGN")!=-1:
                size_list=[] #clear the size_list to be used by summary
                for line in range(1,5):
                    if cell_list[rownum+line][0].strip()!="":
                        summary_colour=cell_list[rownum+line][0].strip().upper()
                        summary_list[summary_colour]={}
                        summary_actual_total=0
                for colnum in range(ncols):
                    current_cell=current_row[colnum]
                    if str(current_cell).strip().upper()=="SIZE":
                        col_summary_size=colnum
                    if str(current_cell).strip().upper()=="TOTAL":
                        col_summary_total=colnum
                    if colnum>col_summary_size and str(current_cell).strip().upper()not in ["TOTAL",""]:
                        size_info={}
                        if isinstance(current_cell,float):
                            size_info["size_name"]=str(int(current_cell)).strip() #for the size like 72.0,16.0,20.0,convert to 72,16,20
                        else:
                            size_info["size_name"]=str(current_cell).strip()

                        size_info["size_col"]=colnum
                        size_list.append(size_info)
                        
            #when come to Total Quantity, end summary and add summary_list to packing_list            
            elif str(current_row[0]).strip().upper()=='TOTAL QUANTITY': 
                summary_seg=False
                packing_list['summary']=summary_list
                for col in range(1,ncols-1):
                    #logger.debug('*****Total quantity col:%s - %s'%(col,current_row[col]))
                    if str(current_row[col]).strip()!="":
                        packing_list["total_quantity"]=current_row[col]
                        break

            #in the summary line for each colour
            else:
                size_qty={}
                for each_size in size_list:
                    if str(current_row[each_size.get('size_col')]).strip()=="":
                        size_qty[each_size.get('size_name')]=0
                    else:
                        size_qty[each_size.get('size_name')]=current_row[each_size.get('size_col')]
                    
                qty_style={'size_qty':size_qty,'total':current_row[col_summary_total]}
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
                    elif str(qty_type).strip().upper()=='RATIO' or str(qty_type).strip()=='' \
                        and cell_list[rownum-4][col_summary_size].strip().upper()=='SIZE':
                        qty_type='Ratio'
                    elif str_contain(qty_type,['SEA','ACTUAL','ship','shipment','shipped','actural','actrual'],all_any='any') or str(qty_type).strip()=='' \
                        and cell_list[rownum-2][col_summary_size].strip().upper()=='SIZE':
                        qty_type='Actual Qty'
                    summary_line[qty_type]=qty_style
                    summary_list[summary_colour]=summary_line
            
        #not in detail_seg or summary_seg
        for colnum in range(ncols):
            current_cell=current_row[colnum]
            if is_size_title(current_row=rownum,current_col=colnum,upper_row=row_size_head,left_col=col_colour_detail,right_col=col_per_carton_pcs):
                if str(current_cell).strip()=="":
                    continue
                size_info={}
                if isinstance(current_cell,float):
                    size_info["size_name"]=str(int(current_cell)).strip()#for the size like 72.0,16.0,20.0,convert to 72,16,20
                else:
                    size_info["size_name"]=str(current_cell).strip()
                size_info["size_col"]=colnum
                size_list.append(size_info)
            if str(current_cell).upper()=="FROM":
                col_from=colnum
                detail_seg=1
            if str(current_cell).upper()=="TO":
                col_to=colnum
            if str_contain(str(current_cell),["CTN","QTY"]):
                col_carton_qty=colnum
            if str_contain(str(current_cell),["DESIGN", "COLOR"]):
                col_colour_detail=colnum
            if str(current_cell).upper().find("ASSORTMENT") !=-1:
                row_size_head=rownum
            if str_contain(current_cell,['per','carton'],'any') and str(cell_list[rownum+1][colnum]).upper()=="PCS":
                col_per_carton_pcs=colnum
            if str_contain(current_cell,['per','carton'],'any') and str(cell_list[rownum+1][colnum]).upper()=="G.W.":
                col_per_carton_gw=colnum
            if str_contain(current_cell,['per','carton'],'any') and str(cell_list[rownum+1][colnum]).upper()=="N.W.":
                col_per_carton_nw=colnum
            if str_contain(current_cell,['sub','total']) and str(cell_list[rownum+1][colnum]).upper()=="PCS":
                col_subtotal=colnum
            if str(current_cell).upper()=="L" and str_contain(cell_list[rownum-1][colnum],['carton','measurement']):
                col_length=colnum
            if str(current_cell).upper()=="W" and str_contain(cell_list[rownum-1][colnum-1],['carton','measurement']):
                col_width=colnum
            if str(current_cell).strip().upper()=="H" and str_contain(cell_list[rownum-1][colnum-2],['carton','measurement']):
                col_height=colnum
            if str_contain(current_cell,['packing','summary']):
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
                        packing_list["style_description"]=str(current_row[col]).strip()
                        break
            if str(current_cell).strip()=="Invoice No.":
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["invoice_no"]=str(current_row[col]).strip()
                        break
            if str(current_cell).strip().upper().startswith("DATE"):
                for col in range(colnum+1,ncols-1):
                    logger.debug('-%s-%s-read invoice date, col-%s,value-%s'%(file,by_name,col,current_row[col]))
                    if str(current_row[col]).strip()!="":
                        try:
                            packing_list["date"]=xlrd.xldate.xldate_as_datetime(current_row[col],0)
                            break
                        except Exception as err:
                            logger.error('--%s-%s- read date error: %s'%(file,by_name,err))
                            break
            if str(current_cell).strip()=="Order No.":
                for col in range(colnum+1,ncols-1):
                    if str(current_row[col]).strip()!="":
                        packing_list["order_no"]=str(current_row[col]).strip()
                        break

    print ('invoice_date=%s,order=%s,style=%s,qty=%s'%(packing_list.get('date'),packing_list.get('order_no')\
                                                       ,packing_list.get('style_no'),packing_list.get('total_quantity')))
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
        count=save(packing_list=packing_list,fty='GZ',file=filename,by_name=sheetname)

    return result
        

    

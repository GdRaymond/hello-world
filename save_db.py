import os
import glob
import size_chart
from tismodels import db,Order_info,Packings,Actual_qty,Detail_carton
from xls_log import get_xls_logger
import datetime
import re

print('save_db before get logger')
logger=get_xls_logger()
print('save_db after get logger')

base_dir=os.path.abspath(os.path.dirname(__file__))

def parse_date(s):
    if s is None:
        logger.info(' not found date info in packing sheet')
        return None
    if type(s) is datetime.date or type(s) is datetime.datetime:
        return s
    if type(s) is str:
        match=re.match(r'([a-zA-Z]+)\s*(\d+)\s*[,.]?(\d+)',s)
        if match is not None:
            month_dict={'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12}
            month=month_dict.get(match.group(1).strip().upper()[0:3])
            day=int(match.group(2))
            year=int(match.group(3))
            if year<100:
                year=2000+year
            result=datetime.date(year,month,day)
            logger.debug('-get the invoice date is %s'%result)
            return result

        match=re.match(r'(\d+)\s*([a-zA-Z]+)\s*[,.]?(\d+)',s)
        if match is not None:
            month_dict={'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12}
            month=month_dict.get(match.group(2).strip().upper()[0:3])
            day=int(match.group(1))
            year=int(match.group(3))
            if year<100:
                year=2000+year
            result=datetime.date(year,month,day)
            logger.debug('-get the invoice date is %s'%result)
            return result

        match=re.match(r'(\d{4})\s*[,.-]\s*(\d+)\s*[,.-]\s*(\d+)',s)  #yyyy.mm.dd
        if match is not None:
            month=int(match.group(2))
            day=int(match.group(3))
            year=int(match.group(1))
            if year<100:
                year=2000+year
            logger.debug('-get the invoice date yyyy-%s,mm-%s,dd-%s'%(year,month,day))
            result=datetime.date(year,month,day)
            logger.debug('-get the invoice date is %s'%result)
            return result

        match=re.match(r'(\d+)\s*[,.-]\s*(\d+)\s*[,.-]\s*(\d{4})',s)  #dd.mm.yyyy
        if match is not None:
            month=int(match.group(2))
            day=int(match.group(1))
            year=int(match.group(3))
            if year<100:
                year=2000+year
            logger.debug('-get the invoice date yyyy-%s,mm-%s,dd-%s'%(year,month,day))
            result=datetime.date(year,month,day)
            logger.debug('-get the invoice date is %s'%result)
            return result

        match=re.match(r'(\d+)\s*[,.-]\s*(\d+)\s*[,.-]\s*(\d{2})',s)  #dd.mm.yy
        if match is not None:
            month=int(match.group(2))
            day=int(match.group(1))
            year=int(match.group(3))
            if year<100:
                year=2000+year
            logger.debug('-get the invoice date yyyy-%s,mm-%s,dd-%s'%(year,month,day))
            result=datetime.date(year,month,day)
            logger.debug('-get the invoice date is %s'%result)
            return result

        logger.error('- can not match format May 20,2012 or 20 May,2012 or 2012.5.20 or 2012-5-20')
        return None
    logger.error(' - the type of parameter-s is %s,can not parse'%type(s))
    return None

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


#read packing list ,save to table order_info
def save(packing_list={},fty='',file='filename',by_name='sheet_name'):
    summary=packing_list.get('summary')
    size_list=size_chart.get_size_list(packing_list.get('style_no'))
    if size_list is None:
        logger.error('-%s-%s-when save style-%s to DB, error : no such style in size_chart,please add'\
                     %(file,by_name,packing_list.get('style_no')))
        return False
    
    count_summary=0
    """
    get general common info in this sheet, save to table packings
    """
    packings=Packings()
    packings.tis_no=packing_list.get('order_no')
    packings.fty=fty
    packings.style_no=packing_list.get('style_no')
    packings.commodity=packing_list.get('style_description')
    packings.invoice_no=packing_list.get('invoice_no')
    packings.total_qty=packing_list.get('total_quantity')
    packings.total_carton=packing_list.get('total_carton')
    packings.total_gw=packing_list.get('total_gw')
    packings.total_volume=packing_list.get('total_volume')
    packings.invoice_date=parse_date(packing_list.get('date'))
    packings.source='%s-%s'%(file,by_name)
    db.session.add(packings)
    db.session.commit()
    packing_id=packings.id

    """
    get the summary in packing list to save order_info, actual_qty
    """
    for colour in summary:

        """
        below save to order_info
        """
        new_order_info=Order_info()
        new_order_info.tis_no=packing_list.get('order_no')
        new_order_info.fty=fty
        new_order_info.style_no=packing_list.get('style_no')
        new_order_info.commodity=packing_list.get('style_description')
        new_order_info.colour=colour
        new_order_info.status='finished'
        new_order_source='%s-%s'%(file,by_name)

        """
          below code  set the size quantity dinamicly according to the garment type, use the techonoly of Class __dict__ to
          get the memeber variable name
        """
        size_qty=summary.get(colour).get('Order Qty').get('size_qty')
        for size_no in range(len(size_list)):
            size_name=size_list[size_no]
            quantity=size_qty.get(size_name)
            if quantity is None:
                logger.debug('- %s-%s-packing list order_info not contain size-qty, size[%s]-%s'%(file,by_name,size_no+1,size_name))
                continue
            logger.debug('- %s-%s-writing size-qty order_info, size[%s]-%s, qty=%s'%(file,by_name,size_no+1,size_name,quantity))
            new_order_info.__dict__['size_%s'%(size_no+1)]=quantity

        logger.debug('- %s-%s-assemble order info %s'%(file,by_name,new_order_info))
        db.session.add(new_order_info)
        db.session.commit()

        """
        below save to actual_qty
        """
        new_actual_qty=Actual_qty()
        new_actual_qty.packing_id=packing_id
        new_actual_qty.colour=colour
        new_actual_qty.total_qty=summary.get(colour).get('Actual Qty').get('total')
        """
          below code  set the size quantity dinamicly according to the garment type, use the techonoly of Class __dict__ to
          get the memeber variable name
        """
        size_qty=summary.get(colour).get('Actual Qty').get('size_qty')
        for size_no in range(len(size_list)):
            size_name=size_list[size_no]
            quantity=size_qty.get(size_name)
            if quantity is None:
                logger.debug('- %s-%s-packing list actual_qty not contain size-qty, size[%s]-%s'%(file,by_name,size_no+1,size_name))
                continue
            logger.debug('- %s-%s-writing actual_qty size-qty, size[%s]-%s, qty=%s'%(file,by_name,size_no+1,size_name,quantity))
            new_actual_qty.__dict__['size_%s'%(size_no+1)]=quantity

        logger.debug('- %s-%s-assemble actual qty %s'%(file,by_name,new_actual_qty))
        db.session.add(new_actual_qty)
        db.session.commit()
        
        count_summary+=1
    logger.debug('-%s-%s-finish save order_info, actual_qty to DB records each %s'%(file,by_name,count_summary))

    """
    save detail carton record
    """
    detail=packing_list.get('detail')
    count_detail_carton=0
    for line in detail:
        detail_carton=Detail_carton()
        detail_carton.packing_id=packing_id
        detail_carton.from_carton=line.get('from')
        detail_carton.to_carton=line.get('to')
        if line.get('carton_qty') is not None and str(line.get('carton_qty')).strip()!='':
            detail_carton.carton_qty=line.get('carton_qty')
        if line.get('colour_detail') is not None and str(line.get('colour_detail')).strip()!='':
            detail_carton.colour=line.get('colour_detail')
        if line.get('per_carton_pcs') is not None and str(line.get('per_carton_pcs')).strip()!='':
            detail_carton.per_carton_pcs=line.get('per_carton_pcs')
        if line.get('per_carton_gw') is not None and str(line.get('per_carton_gw')).strip()!='':
            detail_carton.per_carton_gw=line.get('per_carton_gw')
        if line.get('per_carton_nw') is not None and str(line.get('per_carton_nw')).strip()!='':
            detail_carton.per_carton_nw=line.get('per_carton_nw')
        if line.get('subtotal') is not None and str(line.get('subtotal')).strip()!='':
            detail_carton.subtotal=line.get('subtotal')  
        if line.get('length') is not None and str(line.get('length')).strip()!='':
            detail_carton.length=line.get('length')
        if line.get('width') is not None and str(line.get('width')).strip()!='':
            detail_carton.width=line.get('width')
        if line.get('height') is not None and str(line.get('height')).strip()!='':
            detail_carton.height=line.get('height')
        size_qty=line.get('size_qty')
        for size_no in range(len(size_list)):
            size_name=size_list[size_no]
            quantity=size_qty.get(size_name)
            if quantity is None:
                logger.debug('- %s-%s-packing list detail not contain size-qty, size[%s]-%s'%(file,by_name,size_no+1,size_name))
                continue
            logger.debug('- %s-%s-writing detail size-qty, size[%s]-%s, qty=%s'%(file,by_name,size_no+1,size_name,quantity))
            detail_carton.__dict__['size_%s'%(size_no+1)]=quantity

        logger.debug('- %s-%s-assemble detail carton %s'%(file,by_name,detail_carton))
        db.session.add(detail_carton)
        db.session.commit()
        count_detail_carton+=1
    logger.debug('-%s-%s-finish save detail carton to DB records each %s'%(file,by_name,count_detail_carton))
        
    
    return count_summary+count_detail_carton

            
        

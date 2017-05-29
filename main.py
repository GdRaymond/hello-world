import os
import glob
from read_excel import read_excel_file
import xls_log
logger=xls_log.init_xls_logger()
import th_packing
import gz_packing
import st_packing
import jf_packing
import lt_packing



logger=xls_log.get_xls_logger()


base_dir=os.path.abspath(os.path.dirname(__file__))

#def validate_packing_list_gz(path='gz'):
def abc(path='th',save_db=False):
    data_path=os.path.join(base_dir,path)
    logger.info('Data_path=%s, len=%s'%(data_path,len(data_path)))
    file_position=len(data_path)+1 #including the char \
    data_path=os.path.join(data_path,"*.xls*")
    logger.debug('data files path=%s'%data_path)
    files=glob.glob(data_path)
    logger.info('Total %s files'%len(files))
    
    for file in files:
        logger.info('Start reading file%s'%(file))
        if file.startswith('~',file_position):
            continue
        if path=='jf':
            excel_content=read_excel_file(file,'宁波金丰进出口有限公司')
        else:
            excel_content=read_excel_file(file)

        if excel_content is None:
            continue
        logger.info('- file%s has total %s sheets'%(file,len(excel_content.get('sheets'))))
        for sheetname in sorted([each for each in excel_content.get('sheets')]):
            logger.info('-Start to check the sheet %s'%sheetname)
            if path=='gz':
                result=gz_packing.validate_packinglist_by_sheet(cell_list=excel_content.get('sheets').get(sheetname), \
                                                    filename=excel_content.get('filename'), \
                                                    sheetname=sheetname,save_db=save_db)
            elif path=='th':
                result=th_packing.validate_packinglist_by_sheet(cell_list=excel_content.get('sheets').get(sheetname), \
                                                    filename=excel_content.get('filename'), \
                                                    sheetname=sheetname,save_db=save_db)
            elif path=='st':
                result=st_packing.validate_packinglist_by_sheet(cell_list=excel_content.get('sheets').get(sheetname), \
                                                    filename=excel_content.get('filename'), \
                                                    sheetname=sheetname)
            elif path=='jf':
                result=jf_packing.validate_packinglist_by_sheet(cell_list=excel_content.get('sheets').get(sheetname), \
                                                    filename=excel_content.get('filename'), \
                                                    sheetname=sheetname,save_db=save_db)
            elif path=='lt':
                result=lt_packing.validate_packinglist_by_sheet(cell_list=excel_content.get('sheets').get(sheetname), \
                                                    filename=excel_content.get('filename'), \
                                                    sheetname=sheetname,save_db=save_db)
            if result:
                logger.info('--Correct')
        logger.info('-Finish file')
    

import os
import re
import shutil
def cp(src_dir='.',dest_dir='lt',total=0):
    for lists in os.listdir(src_dir):
        path=os.path.join(src_dir,lists)
        match=re.match(r'.*(packing.*list|京纺箱单).*\.xls[x]?',path,re.I)
        if match :
            total+=1
            shutil.copy(path,dest_dir)
            print('%s:%s'%(str(total),path))
        if os.path.isdir(path):
            total=cp(path,dest_dir,total)
    return total    

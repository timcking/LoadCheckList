import sys
from PyAccess import PyAccess

if __name__ == '__main__':
    accessData = PyAccess()    
    
    # Insert some rows
    args = ('Tim Top', 4, 'HR', 'A ref')
    accessData.insert_chklist(args)    
    
    # Run a query
    rows = accessData.get_chklists('HR')
    iTotRecs = len(rows)
    print ("%s found" % (iTotRecs))
    
    for chklst_id in rows:
        print chklst_id[0]
    
    # Cleanup
    accessData.close_cursor()
import sys
from PyAccess import PyAccess

if __name__ == '__main__':
    accessData = PyAccess()    
    
    # Insert some rows
    # args = ('Tim Top', 4, 'HR', 'A ref')
    # newID = accessData.insert_chklist(args)    
    # print ("New Checklist = %s" % (newID))
    
    # Run a query
    rows = accessData.get_chklists('HR')
    iTotRecs = len(rows)
    print ("%s found" % (iTotRecs))
    
    for chklst_id in rows:
        print chklst_id[0]
    
    # Update
    count = accessData.update_chklist('101', 'CM')
    print ("%s rows updated" % (count))
        
    # Delete
    # count = accessData.delete_chklist(100)
    # print ("%s rows deleted" % (count))
    
    # Query Excel using a named range
    rows = accessData.query_excel()
    for exData in rows:
        print exData
    
    # Cleanup
    accessData.close_cursor()
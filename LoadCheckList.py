import sys
from PyAccess import PyAccess

if __name__ == '__main__':
    accessData = PyAccess()

    # print "*************** Insert New Checklist ***************"
    # args = ('Tim Top', 4, 'HR', 'A ref')
    # newID = accessData.insert_chklist(args)
    # print ("New Checklist = %s" % (newID))

    #print "**************** Get Headers ***********************"
    #rows = accessData.get_excel_headers()
    #for row in rows:
        #print ("Seq: %d, Text: %s" % (row[0], row[1]))

    # Could use this to insert into another table
    # print rows:  [(1.0, 'Professional Delivery'), (2.0, 'Defects'), (3.0, 'Unit Tests')]
    # executemany("insert into tblChkLstHeader(sequence, chklst_header_text) values (?, ?)", rows)

    print "*********** Insert a Checklist from Excel ************"
    args = accessData.get_excel_chklist()
    newID = accessData.insert_chklist(args)
    print ("Last Checklist = %s" % (newID))

    #print "**************** Run a Query ***********************"
    #rows = accessData.get_chklists('HR')
    #iTotRecs = len(rows)
    #print ("%s found" % (iTotRecs))
    #for chklst_id in rows:
        #print chklst_id[0]

    #print "*************** Update Area Code *******************"
    #count = accessData.update_areacode('101', 'RM')
    #print ("%s rows updated" % (count))

    #print "*************** Delete a Row ***********************"
    #count = accessData.delete_chklist(100)
    #print ("%s rows deleted" % (count))

    #print "*************** Query Excel ************************"
    #rows = accessData.get_excel_items()
    #for exData in rows:
        #print exData

    # Cleanup
    accessData.close_cursor()
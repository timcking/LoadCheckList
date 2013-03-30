import pyodbc

class PyAccess():
    connect_str = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\signal_dev.accdb;'
    conn = pyodbc.connect(connect_str)
    cursor = conn.cursor()

    def __init__(self):
        pass
    
    def insert_chklist(self, args):
        sql = "INSERT INTO tblChkLst(chklst_name, template_id, area_code, reference) " +\
              "VALUES (?, ?, ?, ?)"
        
        self.cursor.execute(sql, args)
        self.conn.commit()
        
    def get_chklists(self, area):
        sql = "SELECT chklst_id " +\
               " FROM tblChkLst " +\
               "WHERE area_code = ? " +\
               "ORDER BY chklst_id;"
    
        self.cursor.execute(sql, (area, ))
        rows = self.cursor.fetchall()
        return rows
        
    def close_cursor(self):
        self.cursor.close()        
        
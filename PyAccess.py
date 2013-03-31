import pyodbc

class PyAccess():
    connect_str = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\signal_dev.accdb;'
    conn = pyodbc.connect(connect_str)
    cursor = conn.cursor()

    def __init__(self):
        pass
    
    def query_excel(self):
        DSN = 'Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=.\unit_tests.xlsx;'
        exConn = pyodbc.connect(DSN, autocommit = True)
        exCursor = exConn.cursor()
        
        sql = "SELECT * from tblItems"
        
        exCursor.execute(sql, )
        rows = exCursor.fetchall()
        
        exCursor.close()
        exConn.close()
        return rows
    
    def delete_chklist(self, chklst_id):
        sql = "DELETE FROM tblChkLst WHERE chklst_id = ?;"
        count = self.cursor.execute(sql, (chklst_id, )).rowcount
        self.conn.commit()
        return count
        
    def update_chklist(self, chklst_id, area_code):
        sql = "UPDATE tblChkLst " +\
              "SET area_code = ? " +\
              "WHERE chklst_id = ?;"
        count = self.cursor.execute(sql, (area_code, chklst_id, )).rowcount
        self.conn.commit()
        return count
        
    def insert_chklist(self, args):
        sql = "INSERT INTO tblChkLst(chklst_name, template_id, area_code, reference) " +\
              "VALUES (?, ?, ?, ?)"
        
        self.cursor.execute(sql, args)
        self.conn.commit()
        
        # Get ID of the one we just inserted
        sql = "SELECT MAX(chklst_id) FROM tblChkLst;"
        self.cursor.execute(sql, )
        row = self.cursor.fetchone()
        return row[0]
        
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
        
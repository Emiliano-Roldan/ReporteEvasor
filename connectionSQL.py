import pyodbc
from logger import logger
import tkinter.messagebox as messagebox

class SQLServerConnection:
    def __init__(self, server, database, username, password, port):
        self.server = server
        self.database = database
        self.username = username
        self.password = password
        self.port = port
        self.connection = None
        self.log = logger()

    def connect(self):
        try:
            self.connection = pyodbc.connect(
                'DRIVER={ODBC Driver 11 for SQL Server};'
                f'SERVER={self.server},{self.port};' 
                f'DATABASE={self.database};'
                f'UID={self.username};'
                f'PWD={self.password}'
            )
        except pyodbc.Error as e:
            self.log.write_to_log(f"(SQLServerConnection - connect) - Error connecting to SQL Server: {str(e)}")
            messagebox.showerror("Error", f"Error connecting to SQL Server: {str(e)}")

    def disconnect(self):
        if self.connection:
            self.connection.close()

class SQLServerQueryExecutor:
    def __init__(self, connection):
        self.connection = connection
        self.log = logger()
    
    def execute_query(self, query):
        try:
            cursor = self.connection.cursor()
            cursor.execute(query)
            rows = cursor.fetchall()
            return rows
        except pyodbc.Error as e:
            self.log.write_to_log(f"(SQLServerQueryExecutor - execute_query) - Error executing query: {str(e)}")
            messagebox.showerror("Error", f"Error executing query: {str(e)}")
            return None

class SQLServerDataManipulator:
    def __init__(self, connection):
        self.connection = connection
        self.log = logger()
    
    def execute_non_query(self, query):
        try:
            cursor = self.connection.cursor()
            cursor.execute(query)
            self.connection.commit()
        except pyodbc.Error as e:
            self.log.write_to_log(f"(SQLServerDataManipulator - execute_non_query) - Error executing non-query: {str(e)}")
            messagebox.showerror("Error", f"Error executing non-query: {str(e)}")

    def insert(self, query):
        self.execute_non_query(query)

    def update(self, query):
        self.execute_non_query(query)

    def delete(self, query):
        self.execute_non_query(query)
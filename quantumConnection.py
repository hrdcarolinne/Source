import pyodbc
import traceback
from displayColorama import *
from colorama import *

def testQuantumConnection():
    try: 
        print('\tQuantum Database', end='')
        quantum_connection = pyodbc.connect('DSN=MAXQ;UID=qctl;PWD=quantum')
        
        printColor(' \u2713', Fore.GREEN + Style.BRIGHT, endWith='\t')

        quantum_connection.close()

        return True
    
    except Exception as e:
        error_string = "ERROR: Quantum Connection \n" + str(e) + ":\t" + str(traceback.format_exc()) + "\n\n" 
        printError('\t' + error_string)
        return False
    
def quantumConnectionCursor(quantumSQL):
    quantum_connection = pyodbc.connect('DSN=MAXQ;UID=qctl;PWD=quantum')
    quantum_cursor = quantum_connection.cursor()
    sql_access = quantumSQL
    try:
        quantum_cursor.execute(sql_access)
        
    except Exception as e:
        error_string = "ERROR: " + str(e) + "\n" + str(traceback.format_exc()) + "\n\n" + "SQL: " + str(quantumSQL)
        report_string = report_string + "\n" + error_string + "\n"
        print(error_string)

    finally:
        return quantum_cursor
def quantumConnectionFetchAll(quantumSQL):
    quantum_query = []
    quantum_connection = pyodbc.connect('DSN=MAXQ;UID=qctl;PWD=quantum')
    quantum_cursor = quantum_connection.cursor()
    sql_access = quantumSQL
    try:
        quantum_cursor.execute(sql_access)
        
    except Exception as e:
        error_string = "ERROR: " + str(e) + "\n" + str(traceback.format_exc()) + "\n\n" + "SQL: " + str(quantumSQL)
        report_string = report_string + "\n" + error_string + "\n"
        print(error_string)

    finally:
        quantum_query = quantum_cursor.fetchall()
        quantum_cursor.close

        return quantum_query

    
def quantumConnectionFetchOne(quantumSQL):
    quantum_query = ()
    quantum_connection = pyodbc.connect('DSN=MAXQ;UID=qctl;PWD=quantum')
    quantum_cursor = quantum_connection.cursor()
    sql_access = quantumSQL
    try:
        quantum_cursor.execute(sql_access)
        
    except Exception as e:
        error_string = "ERROR: " + str(e) + "\n" + str(traceback.format_exc()) + "\n\n" + "SQL: " + str(quantumSQL)
        report_string = report_string + "\n" + error_string + "\n"
        print(error_string)

    finally:

        quantum_query = quantum_cursor.fetchone()
        quantum_cursor.close

        return quantum_query

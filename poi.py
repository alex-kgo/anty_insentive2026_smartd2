import pyodbc

conn_str = (
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\smartd2\poi.dat;"
)

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    print("연결 성공")
    print([row for row in cursor.tables()])
except Exception as e:
    print("연결 실패:", e)

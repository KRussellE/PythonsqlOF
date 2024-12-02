import mysql.connector
from mysql.connector import Error
import mysqlconnection


def connect_to_sql(output_box, status_label):
    try:
        connection = mysql.connector.connect(
            host="access-sync.cnomqm8qwozn.eu-north-1.rds.amazonaws.com",
            user="Ogden",
            password="wLzp7ueqgGigbzL",
            database="Access-Info"
        )
        
        if connection.is_connected():
            cursor = connection.cursor()
            cursor.execute("SHOW TABLES")
            tables = cursor.fetchall()

            output_box.insert(tk.END, "\n\nSQL kapcsolat sikeresen létrejött!\n")
            output_box.insert(tk.END, "Táblák az adatbázisban:\n")
            for table in tables:
                output_box.insert(tk.END, f"{table[0]}\n")

            cursor.close()
            connection.close()
            status_label.config(text="Successfully connected to SQL", fg="green")

    except Error as e:
        output_box.insert(tk.END, f"Hiba az SQL kapcsolat során: {e}\n")
        status_label.config(text="SQL connection failed", fg="red")


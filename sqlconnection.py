import mysql.connector
from mysql.connector import Error

def connect_to_sql(output_box, status_label):
    try:
        # SQL kapcsolat létrehozása
        connection = mysql.connector.connect(
            host="your_host",       # Példa: "127.0.0.1" vagy távoli szerver címe
            user="your_username",   # Felhasználónév
            password="your_password",  # Jelszó
            database="your_database"   # Adatbázis neve
        )
        
        if connection.is_connected():
            # Táblák lekérdezése
            cursor = connection.cursor()
            cursor.execute("SHOW TABLES")
            tables = cursor.fetchall()
            
            output_box.insert(tk.END, "\n\nSQL kapcsolat sikeresen létrejött!\n")
            output_box.insert(tk.END, "Táblák az adatbázisban:\n")
            for table in tables:
                output_box.insert(tk.END, f"{table[0]}\n")
            
            cursor.close()
            connection.close()

            # Állapot visszajelzése
            status_label.config(text="Successfully connected to SQL", fg="green")

    except Error as e:
        output_box.insert(tk.END, f"Hiba az SQL kapcsolat során: {e}\n")
        status_label.config(text="SQL connection failed", fg="red")

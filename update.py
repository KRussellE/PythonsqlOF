import sqlconnection  # Reuse the SQL connection module
import tkinter as tk

def update_data_from_sql(data_list, output_box):
    """Fetch and update records based on Docket Number."""
    try:
        for record in data_list:
            docket_number = record.get('Docket Number')
            if docket_number:
                conn = sqlconnection.get_connection()
                cursor = conn.cursor()

                # Query for the docket number
                query = "SELECT * FROM orders WHERE docket_number = ?"
                cursor.execute(query, (docket_number,))
                result = cursor.fetchone()

                if result:
                    # Assuming result columns: docket_number, quantity, prices, etc.
                    record.update({
                        'Docket Number': result[0],
                        'quantity': result[1],
                        'Prices': result[2],
                    })
                else:
                    output_box.insert(tk.END, f"No data found for Docket Number: {docket_number}\n")

                conn.close()

        output_box.insert(tk.END, f"Updated Data: {data_list}\n")

    except Exception as e:
        output_box.insert(tk.END, f"Error updating data: {e}\n")

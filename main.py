import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
import pandas as pd
from io import BytesIO
from email.mime.base import MIMEBase
from email import encoders
import threading


class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Management App")

        # Connect to SQLite database
        self.conn = sqlite3.connect('inventory.db')
        self.cursor = self.conn.cursor()

        # Create inventory table if not exists
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS inventory (
                            part_name TEXT PRIMARY KEY,
                            quantity INTEGER
                            )''')
        self.conn.commit()

        # Set font styles
        font_style = ("Helvetica", 14)

        # Create style for rounded buttons
        self.style = ttk.Style()
        self.style.configure('TButton', font=font_style, padding=5, relief=tk.RAISED)

        # Create labels and entry widgets for adding parts
        ttk.Label(root, text="Part Name:", font=font_style).grid(row=0, column=0, padx=10, pady=5)
        self.add_part_name_entry = ttk.Entry(root, font=font_style)
        self.add_part_name_entry.grid(row=0, column=1, padx=10, pady=5)

        ttk.Label(root, text="Quantity:", font=font_style).grid(row=1, column=0, padx=10, pady=5)
        self.add_quantity_entry = ttk.Entry(root, font=font_style)
        self.add_quantity_entry.grid(row=1, column=1, padx=10, pady=5)

        self.add_button = ttk.Button(root, text="Add Part", command=self.add_part)
        self.add_button.grid(row=2, column=0, columnspan=2, padx=10, pady=5)

        # Create label and entry widget for using parts
        ttk.Label(root, text="Part Name:", font=font_style).grid(row=3, column=0, padx=10, pady=5)
        self.use_part_name_entry = ttk.Entry(root, font=font_style)
        self.use_part_name_entry.grid(row=3, column=1, padx=10, pady=5)

        ttk.Label(root, text="Quantity Used:", font=font_style).grid(row=4, column=0, padx=10, pady=5)
        self.use_quantity_entry = ttk.Entry(root, font=font_style)
        self.use_quantity_entry.grid(row=4, column=1, padx=10, pady=5)

        self.use_button = ttk.Button(root, text="Use Part", command=self.use_part)
        self.use_button.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

        # Create button for adding quantity to an existing part
        ttk.Label(root, text="Part Name to Add Quantity:", font=font_style).grid(row=6, column=0, padx=10, pady=5)
        self.add_quantity_to_part_entry = ttk.Entry(root, font=font_style)
        self.add_quantity_to_part_entry.grid(row=6, column=1, padx=10, pady=5)

        ttk.Label(root, text="Quantity to Add:", font=font_style).grid(row=7, column=0, padx=10, pady=5)
        self.quantity_to_add_entry = ttk.Entry(root, font=font_style)
        self.quantity_to_add_entry.grid(row=7, column=1, padx=10, pady=5)

        self.add_quantity_button = ttk.Button(root, text="Add Quantity to Part", command=self.add_quantity_to_part)
        self.add_quantity_button.grid(row=8, column=0, columnspan=2, padx=10, pady=5)

        # Create button for viewing inventory
        self.view_inventory_button = ttk.Button(root, text="View Inventory", command=self.view_inventory)
        self.view_inventory_button.grid(row=9, column=0, columnspan=2, padx=10, pady=5)

        # Create button for viewing low quantity parts
        self.view_low_quantity_button = ttk.Button(root, text="View Low Quantity Parts",
                                                   command=self.show_low_quantity_parts)
        self.view_low_quantity_button.grid(row=10, column=0, columnspan=2, padx=10, pady=5)

        # Create labels and entry widgets for deleting parts
        ttk.Label(root, text="Part Name to Delete:", font=font_style).grid(row=11, column=0, padx=10, pady=5)
        self.delete_part_name_entry = ttk.Entry(root, font=font_style)
        self.delete_part_name_entry.grid(row=11, column=1, padx=10, pady=5)

        self.delete_button = ttk.Button(root, text="Delete Part", command=self.delete_part)
        self.delete_button.grid(row=12, column=0, columnspan=2, padx=10, pady=5)

        # Create label and entry widget for searching parts
        ttk.Label(root, text="Search Parts:", font=font_style).grid(row=13, column=0, padx=10, pady=5)
        self.search_entry = ttk.Entry(root, font=font_style)
        self.search_entry.grid(row=13, column=1, padx=10, pady=5)

        self.search_button = ttk.Button(root, text="Search", command=self.search_parts)
        self.search_button.grid(row=14, column=0, columnspan=2, padx=10, pady=5)

        # Create label to display last update time
        self.last_update_label = ttk.Label(root, text="", font=font_style)
        self.last_update_label.grid(row=15, column=0, columnspan=2, padx=10, pady=5)

        # Display last update time
        self.display_last_update_time()

    def add_part(self):
        part_name = self.add_part_name_entry.get()
        quantity = self.add_quantity_entry.get()

        if not part_name or not quantity:
            messagebox.showerror("Error", "Please enter both part name and quantity")
            return

        try:
            quantity = int(quantity)
        except ValueError:
            messagebox.showerror("Error", "Quantity must be a number")
            return

        try:
            self.cursor.execute("INSERT INTO inventory VALUES (?, ?)", (part_name, quantity))
            self.conn.commit()
            self.display_last_update_time()
            self.send_email_inventory_updated()
            messagebox.showinfo("Success", "Part added successfully")
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"Error adding part: {e}")

    def use_part(self):
        part_name = self.use_part_name_entry.get()
        quantity_used = self.use_quantity_entry.get()

        if not part_name or not quantity_used:
            messagebox.showerror("Error", "Please enter both part name and quantity used")
            return

        try:
            quantity_used = int(quantity_used)
            current_quantity = self.get_current_quantity(part_name)
            if current_quantity is None:
                messagebox.showerror("Error", "Part does not exist")
                return
            new_quantity = max(0, current_quantity - quantity_used)
            self.cursor.execute("UPDATE inventory SET quantity = ? WHERE part_name = ?", (new_quantity, part_name))
            self.conn.commit()
            self.display_last_update_time()
            self.send_email_inventory_updated()
            messagebox.showinfo("Success", f"{quantity_used} parts used successfully")
        except ValueError:
            messagebox.showerror("Error", "Quantity used must be a number")

    def add_quantity_to_part(self):
        part_name = self.add_quantity_to_part_entry.get()
        quantity_to_add = self.quantity_to_add_entry.get()

        if not part_name or not quantity_to_add:
            messagebox.showerror("Error", "Please enter both part name and quantity to add")
            return

        try:
            quantity_to_add = int(quantity_to_add)
        except ValueError:
            messagebox.showerror("Error", "Quantity to add must be a number")
            return

        current_quantity = self.get_current_quantity(part_name)
        if current_quantity is None:
            messagebox.showerror("Error", "Part does not exist")
            return

        new_quantity = current_quantity + quantity_to_add
        self.cursor.execute("UPDATE inventory SET quantity = ? WHERE part_name = ?", (new_quantity, part_name))
        self.conn.commit()
        self.display_last_update_time()
        self.send_email_inventory_updated()
        messagebox.showinfo("Success", f"{quantity_to_add} parts added to {part_name} successfully")

    def view_inventory(self):
        self.cursor.execute("SELECT * FROM inventory")
        inventory = self.cursor.fetchall()
        if not inventory:
            messagebox.showinfo("Inventory", "Inventory is empty")
        else:
            # Create a new window to display the inventory data
            inventory_window = tk.Toplevel(self.root)
            inventory_window.title("Inventory")

            # Create a Treeview widget to display inventory data in a grid
            tree = ttk.Treeview(inventory_window, style="Treeview")
            tree["columns"] = ("Part Name", "Quantity")
            tree.heading("#0", text="ID")
            tree.heading("Part Name", text="Part Name")
            tree.heading("Quantity", text="Quantity")

            for i, row in enumerate(inventory, start=1):
                tree.insert("", tk.END, text=str(i), values=row)

            tree.pack(fill="both", expand=True)

    def show_low_quantity_parts(self):
        self.cursor.execute("SELECT * FROM inventory WHERE quantity < 5")
        low_quantity_parts = self.cursor.fetchall()

        if not low_quantity_parts:
            messagebox.showinfo("Low Quantity Parts", "No parts with quantity less than 5.")
            return

        # Create a new window to display the low quantity parts data
        low_quantity_window = tk.Toplevel(self.root)
        low_quantity_window.title("Low Quantity Parts")

        # Create a Treeview widget to display low quantity parts data in a grid
        tree = ttk.Treeview(low_quantity_window, style="Treeview")
        tree["columns"] = ("Part Name", "Quantity")
        tree.heading("#0", text="ID")
        tree.heading("Part Name", text="Part Name")
        tree.heading("Quantity", text="Quantity")

        for i, row in enumerate(low_quantity_parts, start=1):
            tree.insert("", tk.END, text=str(i), values=row)

        tree.pack(fill="both", expand=True)

    def get_current_quantity(self, part_name):
        self.cursor.execute("SELECT quantity FROM inventory WHERE part_name = ?", (part_name,))
        row = self.cursor.fetchone()
        if row:
            return row[0]
        return None

    def delete_part(self):
        part_name = self.delete_part_name_entry.get()

        if not part_name:
            messagebox.showerror("Error", "Please enter part name to delete")
            return

        try:
            self.cursor.execute("DELETE FROM inventory WHERE part_name = ?", (part_name,))
            self.conn.commit()
            self.display_last_update_time()
            self.send_email_inventory_updated()
            messagebox.showinfo("Success", f"{part_name} deleted successfully")
        except sqlite3.Error as e:
            messagebox.showerror("Error", f"Error deleting part: {e}")

    def search_parts(self):
        search_query = self.search_entry.get().strip()

        if not search_query:
            messagebox.showerror("Error", "Please enter search query")
            return

        self.cursor.execute("SELECT * FROM inventory WHERE part_name LIKE ?", (f'%{search_query}%',))
        search_result = self.cursor.fetchall()

        if not search_result:
            messagebox.showinfo("Search Result", "No matching parts found")
            return

        search_result_str = "\n".join([f"{row[0]}: {row[1]}" for row in search_result])
        messagebox.showinfo("Search Result", search_result_str)

    def display_last_update_time(self):
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.last_update_label.config(text=f"Last Updated: {current_time}")

    def send_email_inventory_updated(self):
        # Define a function to send the email
        def send_email():
            # Connect to a new SQLite database
            conn = sqlite3.connect('inventory.db')
            cursor = conn.cursor()

            # Fetch the updated inventory and create a DataFrame for demonstration
            cursor.execute("SELECT * FROM inventory")
            inventory = cursor.fetchall()
            df = pd.DataFrame(inventory, columns=['Part Name', 'Quantity'])

            # Close the database connection
            cursor.close()
            conn.close()

            # Email configuration
            email_sender = "aaravkumarjeet042@gmail.com"
            email_receiver = "aaravkumarjeet78@gmail.com"
            email_password = "xkpg pouq ktxj pwjf"

            # Email content
            subject = "Inventory Updated"

            # Connect to the SMTP server
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login(email_sender, email_password)

                # Construct the email message
                msg = MIMEMultipart()
                msg['From'] = email_sender
                msg['To'] = email_receiver
                msg['Subject'] = subject

                # Convert DataFrame to an Excel file in memory
                with BytesIO() as output:
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    output.seek(0)  # Go to the beginning of the BytesIO stream

                    # Attach the Excel file
                    part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    part.set_payload(output.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename='inventory.xlsx')
                    msg.attach(part)

                # Send the email
                server.send_message(msg)

        # Create a thread for sending email
        email_thread = threading.Thread(target=send_email)
        email_thread.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()

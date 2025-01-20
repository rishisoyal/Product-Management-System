import pandas as pd
import os


# Base class
class PMS:

    def save_to_excel(self):
        try:
            self.data_df.to_excel("products.xlsx", index=False)
        except PermissionError:
            print("Permission to access file is denied, close the file and try again")

    def __init__(self):
        # Create products.xlsx if not present
        if "products.xlsx" not in os.listdir():
            print("Creating products.xlsx...")
            # Schema
            self.data = {
                "id": [],
                "customer_name": [],
                "customer_address": [],
                "product_name": [],
                "description": [],
                "brand": [],
                "stock_quantity": [],
                "price": [],
                "discount": [],
                "order_date": [],
                "delivery_date": [],
            }
            # Get dataframe
            self.data_df = pd.DataFrame(self.data)

            # Convert specific fields to the appropriate data types
            self.data_df["order_date"] = pd.to_datetime(
                self.data_df["order_date"], errors="coerce"
            )
            self.data_df["delivery_date"] = pd.to_datetime(
                self.data_df["delivery_date"], errors="coerce"
            )
            self.data_df["price"] = pd.to_numeric(
                self.data_df["price"], errors="coerce"
            )
            self.data_df["id"] = pd.to_numeric(self.data_df["id"], errors="coerce")

            self.save_to_excel()
        else:
            # Get dataframe if products.xlsx allready present
            print("Reading existing products.xlsx...")
            try:
                self.data_df = pd.read_excel("./products.xlsx")
            except PermissionError:
                print("Permission to read file is denied, close the file and try again")

    def show_entries(self):
        print(self.data_df)

    def add_entry(self, entry):
        # Generate a new ID
        new_id = len(self.data_df) + 1  # Incremental ID
        entry_with_id = [new_id] + entry  # Add the ID as the first element

        # using loc to add a single row at the end of the DF
        if len(entry_with_id) != len(self.data_df.columns):
            print("Error: Entry does not match the required number of columns.")
            return
        self.data_df.loc[len(self.data_df)] = entry_with_id
        self.save_to_excel()

    def delete_entry(self, id):
        # Delete a row at the given index
        # Ex: df = df.drop(index=1)   # Deletes the row with index 1
        # Delete row based on condition
        # Ex: df = df[df["customer_name"] != "Bob"]  # Deletes rows where customer_name is "Bob"
        if id not in self.data_df["id"].values:
            print(f"Id {id} not found...")
            return
        self.data_df = self.data_df[self.data_df["id"] != id]
        self.data_df.reset_index(drop=True, inplace=True)
        self.save_to_excel()

    def modify_entry(self, id: int, **kwargs):
        if id not in self.data_df["id"].values:
            print(f"Id {id} not found...")
            return
        try:
            row_index = self.data_df[self.data_df["id"] == id].index[0]
            for column, value in kwargs.items():
                if column not in self.data_df.columns:
                    print(f"Column {column} not found...")
                else:
                    self.data_df.loc[row_index, column] = value
            self.data_df.reset_index(drop=True, inplace=True)
            self.save_to_excel()
        except Exception as e:
            print(f"Error deleting entry: {e}")


if __name__ == "__main__":

    instructions = """
       ****************************
        1.) Enter 1 to show entries
        2.) Enter 2 to add entry
        3.) Enter 3 to modify entry
        4.) Enter 4 to delete entry
       ****************************
    """

    pms = PMS()
    os.system("cls")
    while 1:
        print(instructions)
        try:
            command = int(input(">>> "))
        except:
            continue
        if command == 0:
            print("Exiting...")
            break
        elif command == 1:
            pms.show_entries()
        elif command == 2:
            customer_name = input("enter customer name: ")
            customer_address = input("enter customer address: ")
            product_name = (input("enter product name: "),)
            description = input("enter product description: ")
            brand = (input("enter brand name: "),)
            stock_quantity = input("enter stock quantity: ")
            price = (input("enter product price: "),)
            discount = input("enter product discount: ")
            order_date = (input("enter order date: "),)
            delivery_date = input("enter delivery date: ")

            pms.add_entry(
                [
                    customer_name,
                    customer_address,
                    product_name,
                    description,
                    brand,
                    stock_quantity,
                    price,
                    discount,
                    order_date,
                    delivery_date,
                ]
            )
        elif command == 3:
            id = int(input("Enter entry id: "))
            pms.modify_entry(id=id)
        elif command == 4:
            id = int(input("Enter entry id: "))
            pms.delete_entry(id=id)

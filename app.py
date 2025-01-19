import pandas as pd
import os


# product management system class
class PMS:

    def save_to_excel(self):
        try:
            self.data_df.to_excel("products.xlsx", index=False)
        except PermissionError:
            print(
                "Permission to access excel file is denied, close the file and try again"
            )

    def __init__(self):
        # Create products.xlsx if not present
        if "products.xlsx" not in os.listdir():
            print("Creating products.xlsx...")
            # Schema
            self.data = {
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
            # Get dataframe and exclude index
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

            self.save_to_excel()
        else:
            # Get dataframe if products.xlsx allready present
            print("Reading existing products.xlsx...")
            try:
                self.data_df = pd.read_excel("./products.xlsx")
            except PermissionError:
                print(
                    "Permission to read excel file is denied, close the file and try again"
                )

    def add_entry(self, entry):
        # using loc to add a single row at the end of the DF
        if len(entry) != len(self.data_df.columns):
            print("Error: Entry does not match the required number of columns.")
            return
        self.data_df.loc[len(self.data_df)] = entry
        self.save_to_excel()

    def delete_entry(self, **kwargs):
        # Delete a row at the given index
        # Ex: df = df.drop(index=1)   # Deletes the row with index 1
        # Delete row based on condition
        # Ex: df = df[df["customer_name"] != "Bob"]  # Deletes rows where customer_name is "Bob"
        try:
            for key, value in kwargs.items():
                if key in self.data_df.columns:
                    self.data_df = self.data_df[self.data_df[key] != value]
                else:
                    print(f"Key '{key}' does not exist in the DataFrame.")
            self.data_df.reset_index(drop=True, inplace=True)
            self.save_to_excel()
        except Exception as e:
            print(f"Error deleting entry: {e}")


if __name__ == "__main__":
    pms = PMS()
    print(pms.data_df)
    entry = [
        "john",
        "prefer to stay anonymous",
        "alienware laptop",
        "best gaming laptop",
        "dell",
        "12",
        "100000",
        "12%",
        "19-01-2025",
        "25-01-2025",
    ]
    # pms.add_entry(entry)
    pms.delete_entry(customer_name="john")
    print(pms.data_df)

from robocorp.tasks import task
from RPA.Assistant import Assistant
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from pathlib import Path


assistant = Assistant()
customer_data_list = []


class Customer:
    def __init__(self):
        self.first_name = None
        self.last_name = None
        self.age = None
        self.email = None
        self.phone = None
        self.address = None
        self.city = None
        self.postal_code = None

    def __str__(self):
        return f"{self.first_name},{self.last_name},{self.age},{self.email},{self.phone},{self.address},{self.city},{self.postal_code}"

    def update_from_result(self, result):
        self.first_name = result.first_name
        self.last_name = result.last_name
        self.age = result.age
        self.email = result.email
        self.phone = result.phone
        self.address = result.address
        self.city = result.city
        self.postal_code = result.postal_code

    def to_dict(self):
        return {
            "First Name": self.first_name,
            "Last Name": self.last_name,
            "Age": self.age,
            "Email": self.email,
            "Phone": self.phone,
            "Address": self.address,
            "City": self.city,
            "Postal Code": self.postal_code,
        }    


@task
def add_customers_to_excel_file():
    start_screen()
    customer_table = create_or_update_table()
    if customer_table:
        write_into_excel_file(customer_table)
    post_screen()


def write_into_excel_file(customer_table):
    excel = Files()
    file_path = Path("customers.xlsx")
    if not file_path.exists():
        excel.create_workbook(file_path)
        excel.append_rows_to_worksheet(customer_table, header=True)
        excel.save_workbook()
        excel.close_workbook()
    else:
        excel.open_workbook(str(file_path))
        excel.append_rows_to_worksheet(customer_table, header=True)
        excel.save_workbook()
        excel.close_workbook()
    print(f"Customer data written to {file_path} successfully!")


def start_screen():
    assistant = Assistant()
    assistant.add_heading("Customer input helper")
    assistant.add_text("This is customer information input helper. You add one customer at a time, and when you close the window the customers file will be updated")
    assistant.add_button("Add a new customer", add_new_customer)
    assistant.run_dialog(title="Customer input helper")


def post_screen():
    assistant = Assistant()
    assistant.add_heading("Customer Input Summary")
    num_customers = len(customer_data_list)
    assistant.add_text(f"{num_customers} customer(s) have been successfully added to the Excel file.")
    assistant.add_icon("success")
    assistant.run_dialog(title="Customer Input Summary")


def add_new_customer():
    assistant = Assistant()
    assistant.add_heading("Fill in customer details")
    assistant.add_text_input("first_name", label="First name", required=True)
    assistant.add_text_input("last_name", label="Last name", required=True)
    assistant.add_text_input("age", label="Age", required=True)
    assistant.add_text_input("email", label="Email", required=True)
    assistant.add_text_input("phone", label="Phone", required=True)
    assistant.add_text_input("address", label="Address", required=True)
    assistant.add_text_input("city", label="City", required=True)
    assistant.add_text_input("postal_code", label="Postal code", required=True)
    assistant.add_submit_buttons("Submit", default="Submit")
    result = assistant.run_dialog() 
    update_customer(result)


def update_customer(result):
    customer = Customer()
    customer.update_from_result(result)
    customer_data_list.append(customer.to_dict())
    print(f"Added customer: {customer}")
    

def create_or_update_table():
    table = Tables()
    if customer_data_list:
        print("Customer table updated successfully!")
        return table.create_table(customer_data_list)  # Create table from the list of dictionaries
    else:
        print("No customer data available.")
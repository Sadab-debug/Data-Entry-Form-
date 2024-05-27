# Function to use placeholder text in entry box
def onEntryClick(entry, placeholder_text):
    if entry.get() == placeholder_text:
        entry.delete(0, 'end')  # Clear the entry when clicked
        entry.config(fg='black')  # Change text color to black

def onFocusOut(entry, placeholder_text):
    if entry.get() == "":
        entry.insert(0, placeholder_text)  # Add placeholder text
        entry.config(fg='grey')  # Change text color to grey


def clearAll(name_var, age_var, phone_var, birth_var, blood_group_var, address_var, email_var):
    name_var.set('')
    age_var.set('')
    phone_var.set('')
    birth_var.set('')
    blood_group_var.set('A+')
    address_var.set('')
    email_var.set('')


def convertToExcel(Workbook, load_workbook, dataframe_to_rows, dataframe, filename="sample.xlsx"):
    try:
        # Load the existing workbook if it exists
        wb = load_workbook(filename)
    except FileNotFoundError:
        # If the file doesn't exist, create a new workbook
        wb = Workbook()

    # Select the active worksheet
    ws = wb.active

    # Determine the starting row for appending new data
    # starting_row = ws.max_row + 1  # Start appending from the next row after the last existing row

    # Convert the DataFrame to rows and append them to the worksheet
    for row in dataframe_to_rows(dataframe, index=False, header=False):
        ws.append(row)

    # Save the workbook
    wb.save(filename)


def createDataFrame(pd,name_var, age_var, phone_var, birth_var, blood_group_var, address_var, email_var):

    # Check if any of the tkinter variables is empty
    if name_var.get() == "" or age_var.get() == "" or phone_var.get() == "" or birth_var.get() == "" or blood_group_var.get() == "" or address_var.get() == "" or email_var.get() == "":
        print("Error: All fields must be filled.")
        return None  # Return None if any field is empty

    # Create INFO dictionary capturing current values
    INFO = {
        "Name": name_var.get(),
        "Age": age_var.get(),
        "Contact Number": int(phone_var.get()), #convert the phone number in integer
        "Birth Date": birth_var.get(),
        "Blood Group": blood_group_var.get(),
        "Address": address_var.get(),
        "Email ID": email_var.get(),
    }

    # Append a new row to the DataFrame
    df = pd.DataFrame([INFO])

    clearAll(name_var, age_var, phone_var, birth_var, blood_group_var, address_var, email_var)

    # Print the DataFrame to verify the changes
    return df
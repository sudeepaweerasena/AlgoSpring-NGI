import win32com.client
import pyautogui
import time
import xlwings as xw
import win32com.client as win32
import pandas as pd
import threading
import os

# Function to click a specified OLEObject button by name on a given sheet
def click_ole_button(sheet, button_name):
    print(f"Searching for button '{button_name}' to click...")
    for obj in sheet.OLEObjects():
        if obj.Name == button_name:
            print(f"Found button: {obj.Name}")
            try:
                # Check if the object supports the DoClick() or Click() method
                if hasattr(obj.Object, 'DoClick'):
                    obj.Object.DoClick()
                    print("Button clicked successfully using 'DoClick' method!")
                    return True
                elif hasattr(obj.Object, 'Click'):
                    obj.Object.Click()
                    print("Button clicked successfully using 'Click' method!")
                    return True
                elif hasattr(obj.Object, 'Value'):
                    obj.Object.Value = True
                    print("Button clicked successfully using 'Value' property!")
                    return True
                else:
                    print("Button does not support 'DoClick', 'Click', or 'Value' properties.")
                    return False
            except Exception as e:
                print(f"Error clicking the button: {str(e)}")
                return False
    print("Button not found.")
    return False

# Determine Categories
def determine_category(category):
    return '1' if category == 'A' else '2' if category == 'B' else '3' if category == 'C' else None

# Count number of categories
def count_unique_categories(data):
    categories = [row[3].strip() for row in data] 
    unique_categories = set(categories) 
    return len(unique_categories)

# Fill 'Plan' in column I based on the condition from the Network column for each category
def fill_plan_based_on_condition(sheet, df):
    for _, row in df.iterrows():
        category = row['Category']  
        pl_value = row['Network'] 
        
        if category == 'A':
            sheet.Range('G2').Value = pl_value
           
        elif category == 'B':
            sheet.Range('G3').Value = pl_value
            
        elif category == 'C':
            sheet.Range('G4').Value = pl_value

# Fill the number of unique categories
def fill_unique_categories_to_cell(sheet, df3, cell="D5"):
    categories = df3['Category']

    # Count the unique categories
    unique_categories = set(categories)
    num_categories = len(unique_categories)

    # Ensure the sheet and cell are valid
    try:
        sheet.Range(cell).Value = num_categories
        print(f"Filled {num_categories} unique categories into cell {cell}.")
    except Exception as e:
        print(f"Error accessing range {cell} in the sheet: {e}")

# Read data from the qatar.xlsx Sheet2 file into a DataFrame(df)
def read_qatar_data(pd_file_path):
    df = pd.read_excel(pd_file_path, sheet_name='Sheet2')   
    return df

# Read data from the qatar.xlsx Sheet1 file into a DataFrame(df2)
def read_qatar_data_sheet1(pd_file_path):
    df2 = pd.read_excel(pd_file_path, sheet_name='Sheet1')   
    return df2

# Read data from the source file's Sheet1 into a DataFrame(df3)
def read_source_data(source_path):
    df3 = pd.read_excel(source_path, sheet_name='Sheet1')   
    return df3

# Fill Company name 
def insert_company(sheet, df2):
    company_name = df2.loc[df2['KEY'] == 'Company Name', 'VALUE'].values
    print(company_name)
    company_name = company_name[0]
    print(company_name)
    sheet.Range('D7').Value = company_name 

# Fill Plan Details page data
def insert_category_data_to_plan_details(sheet, df):
    for _, row in df.iterrows():
        category = row['Category']  

        # Dental and Optical 
        if category == 'A':
            if row['Dental'] == 'Not Covered' and row['Optical'] == 'Not Covered':
                sheet.Range('D15').Value = 'Not_Covered'
                sheet.Range('D16').Value = 'Not_Covered'
                print("Dental and Optical are 'Not Covered' in Category A, so D15 and D16 set to 'Not_Covered'.")
            else:
                sheet.Range('D15').Value = 'Basic+Scaling'
                sheet.Range('D16').Value = row['Optical']
            
            sheet.Range('D12').Value = row['Annual Limit']
            sheet.Range('D17').Value = row['Telehealth Consultation']
            sheet.Range('D19').Value = row['Wellness Package']
            sheet.Range('D20').Value = row['Data20']
            sheet.Range('D21').Value = row['Data21']
            sheet.Range('D24').Value = row['Data24']
            sheet.Range('D25').Value = row['Data25']
            sheet.Range('D26').Value = row['Data26']
            sheet.Range('D27').Value = row['Data27']
            sheet.Range('D28').Value = row['Data28']
            print("Data for Category A inserted.")
        
        elif category == 'B':
            if row['Dental'] == 'Not Covered' and row['Optical'] == 'Not Covered':
                sheet.Range('E15').Value = 'Not_Covered'
                sheet.Range('E16').Value = 'Not_Covered'
                print("Dental and Optical are 'Not Covered' in Category B, so E15 and E16 set to 'Not_Covered'.")
            else:
                sheet.Range('E15').Value = 'Basic+Scaling'
                sheet.Range('E16').Value = row['Optical']
         
            sheet.Range('E12').Value = row['Annual Limit']
            sheet.Range('E17').Value = row['Telehealth Consultation']
            sheet.Range('E19').Value = row['Wellness Package']
            sheet.Range('E20').Value = row['Data20']
            sheet.Range('E21').Value = row['Data21']
            sheet.Range('E24').Value = row['Data24']
            sheet.Range('E25').Value = row['Data25']
            sheet.Range('E26').Value = row['Data26']
            sheet.Range('E27').Value = row['Data27']
            sheet.Range('E28').Value = row['Data28']
            print("Data for Category B inserted.")

        elif category == 'C':
            if row['Dental'] == 'Not Covered' and row['Optical'] == 'Not Covered':
                sheet.Range('F15').Value = 'Not_Covered'
                sheet.Range('F16').Value = 'Not_Covered'
                print("Dental and Optical are 'Not Covered' in Category C, so F15 and F16 set to 'Not_Covered'.")
            else:
                sheet.Range('F15').Value = 'Basic+Scaling'
                sheet.Range('F16').Value = row['Optical']
           
            sheet.Range('F12').Value = row['Annual Limit']
            sheet.Range('F17').Value = row['Telehealth Consultation']
            sheet.Range('F19').Value = row['Wellness Package']
            sheet.Range('F20').Value = row['Data20']
            sheet.Range('F21').Value = row['Data21']
            sheet.Range('F24').Value = row['Data24']
            sheet.Range('F25').Value = row['Data25']
            sheet.Range('F26').Value = row['Data26']
            sheet.Range('F27').Value = row['Data27']
            sheet.Range('F28').Value = row['Data28']
            print("Data for Category C inserted.")

# Fill Dental and Optical Benefit
def fill_d13_based_on_o2_q2(sheet, df):
    # Logic: For the dental and Optical Benefit is Covered when Dental is Covered and Optical is covered
    for _, row in df.iterrows():
        category = row['Category']
        
        if category == 'A':
            if row['Dental'] == 'Covered' and row['Optical'] == 'Covered':
                sheet.Range('D13').Value = 'Covered'
            else:
                sheet.Range('D13').Value = 'Not_Covered'
        
        elif category == 'B':
            if row['Dental'] == 'Covered' and row['Optical'] == 'Covered':
                sheet.Range('E13').Value = 'Covered'
            else:
                sheet.Range('E13').Value = 'Not_Covered'

        elif category == 'C':
            if row['Dental'] == 'Covered' and row['Optical'] == 'Covered':
                sheet.Range('F13').Value = 'Covered'
            else:
                sheet.Range('F13').Value = 'Not_Covered'

# Click the 'Proceed to Output' button
def click_proceed_to_output_button(sheet):
    button_name = 'CommandButton1'
    click_ole_button(sheet, button_name)

# Click the 'Generate pdf Quotation' button
def click_generate_pdf_output_button(sheet):
    """Click the 'Generate pdf Quotation' button."""
    button_name = 'CommandButton1'
    click_ole_button(sheet, button_name)

# Ipunt Box
def simulate_input():
    time.sleep(2) 
    
    # Define the full path where the PDF will be saved
    file_path = r'D:\\AlgoSpring\\python\\Qatar\\downloadFile.pdf'
    
    # Ensure the directory exists; create it if it doesn't
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"Created directory: {directory}")
    
    # Type the full file path
    pyautogui.write(file_path, interval=0.05)
    print("File path typed successfully!")
    
    # Press Enter to confirm
    pyautogui.press("enter")

# Click 'End' Button on source file
def simulate2_input():
    time.sleep(2) 
    pyautogui.press("enter")

#------------------------------------------------ main ------------------------------------------------------

def main():
    # Initialize Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True 
    
    # Open the workbook
    workbook = excel.Workbooks.Open(r'C:\\Users\\sudeepa.w\\Desktop\\Healthnet.xlsm')
    
    # Maximize the workbook window
    try:
        # Access the first window of the workbook
        workbook_window = workbook.Windows(1)
        
        # Set the window state to maximized using constants
        workbook_window.WindowState = win32.constants.xlMaximized
        print("Excel window maximized successfully!")
    except AttributeError:
        # If win32.constants.xlMaximized is not accessible, define it manually
        xlMaximized = -4137
        workbook_window.WindowState = xlMaximized
        print("Excel window maximized successfully using manual constant!")
    except Exception as e:
        print(f"Error maximizing Excel window: {e}")
    
    # Optionally, bring Excel to the foreground
    try:
        workbook_window.Activate()
        print("Excel window activated and brought to foreground.")
    except Exception as e:
        print(f"Error activating Excel window: {e}")
    
    sheet = workbook.Sheets('Intro Sheet')
    time.sleep(4)

    # Click Intro sheet button
    click_ole_button(sheet, 'CommandButton1')

    # Paths to CensusData.xlsx
    source_path = "D:\\AlgoSpring\\python\\Qatar\\CensusData.xlsx" 

    # Open CensusData.xlsx
    wb_source = xw.Book(source_path)

    # Get Sheet1 data
    df3 = read_source_data(source_path)

    # Get the row count
    row_count = df3.shape[0]

    # Assign Census workbook sheet to instructions_sheet
    instructions_sheet = workbook.Sheets('Census')

    # Fill row count
    thread2 = threading.Thread(target=simulate2_input)  # Corrected 'target2' to 'target'
    thread2.start()
    instructions_sheet.Range('B6').Value = row_count
    thread2.join()

    # Get the number of unique categories in the 'Category' column
    unique_categories = df3['Category'].nunique()
    
    # Print the number of unique categories
    print(f"Number of unique categories: {unique_categories}")

    # Populate Excel with data
    for index, row in df3.iterrows():
        category = row['Category'] 
        # Logic to set Plan value based on category
        if category == 'A':
            plan_value = 'Plan_1'
        elif category == 'B':
            plan_value = 'Plan_2'
        elif category == 'C':
            plan_value = 'Plan_3'
        # Writing data to Excel from DataFrame
        instructions_sheet.Range(f'B{index + 9}').Value = row['Beneficiary First Name'] 
        instructions_sheet.Range(f'C{index + 9}').Value = row['Gender']  
        instructions_sheet.Range(f'D{index + 9}').Value = row['DOB']  
        instructions_sheet.Range(f'E{index + 9}').Value = determine_category(row['Category'])
        instructions_sheet.Range(f'F{index + 9}').Value = row['Marital status']  
        instructions_sheet.Range(f'G{index + 9}').Value = row['Relation'] 
        instructions_sheet.Range(f'H{index + 9}').Value = row['Visa Issued Emirates']  
        instructions_sheet.Range(f'I{index + 9}').Value = plan_value  

    # Path to the qatar data file
    pd_file_path = r"D:\\AlgoSpring\\python\\Qatar\\qatar.xlsx"
    
    # Read the data from qatar.xlsx Sheet2 into a DataFrame
    df = read_qatar_data(pd_file_path)

    # Read the data from qatar.xlsx Sheet1 into a DataFrame
    df2 = read_qatar_data_sheet1(pd_file_path)
  
    # Fill 'Plan' in column I based on the condition from the pl column for each category
    fill_plan_based_on_condition(workbook.Sheets('Census'), df)
    time.sleep(4)

    # Close the source workbook
    wb_source.close()
    time.sleep(3)
 
    # Click 'Proceed to Plan Details' Button
    click_ole_button(workbook.Sheets('Census'), 'CommandButton2')  
    time.sleep(2)
    # Correct reference to the 'Plan Details' sheet
    fill_unique_categories_to_cell(workbook.Sheets('Plan Details'), df3)
    
    # Fill Company name 
    insert_company(workbook.Sheets('Plan Details'), df2)

    # Fill Dental and Optical Benefit
    fill_d13_based_on_o2_q2(workbook.Sheets('Plan Details'), df)
    
    # Insert the data into the PlanDetails sheet
    insert_category_data_to_plan_details(workbook.Sheets('Plan Details'), df)

    # Click Proceed to Output button
    click_proceed_to_output_button(workbook.Sheets('Plan Details'))

    # Click the Generate pdf Quotation button
    thread = threading.Thread(target=simulate_input)
    thread.start()
    click_generate_pdf_output_button(workbook.Sheets('Output Sheet'))
    thread.join()


if __name__ == "__main__":
    main()

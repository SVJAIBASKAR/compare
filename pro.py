import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

def rerun():
    raise st.script_runner.RerunException(st.script_request_queue.RerunData(None))
# Function to add customer details

def set_page_size(doc, width_mm=148, height_mm=210):
    # Convert mm to twips (1 mm = 567 twips)
    width_twips = int(width_mm * 567)
    height_twips = int(height_mm * 567)

    # Access the document's section properties
    for section in doc.sections:
        # Get the section's XML element
        sect_pr = section._sectPr

        # Create or modify the pgSize element
        pg_size = OxmlElement('w:pgSize')
        pg_size.set(qn('w:w'), str(width_twips))
        pg_size.set(qn('w:h'), str(height_twips))

        # Remove any existing pgSize element
        for existing_pg_size in sect_pr.findall(qn('w:pgSize')):
            sect_pr.remove(existing_pg_size)

        # Add the new pgSize element
        sect_pr.append(pg_size)

def add_customer_details(doc, customer_name, address, phone, product_name,bill_number):
    right_paragraph = doc.add_paragraph( bill_number)
    right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right_run = right_paragraph.runs[0]
    right_run.font.size = Pt(16)
    table = doc.add_table(rows=7, cols=2)

    # Add customer name
    cell1 = table.cell(0, 0)
    cell1.text = "Customer Name:"
    cell1_run = cell1.paragraphs[0].runs[0]
    cell1_run.font.size = Pt(16)

    cell2 = table.cell(0, 1)
    cell2.text = customer_name
    cell2_run = cell2.paragraphs[0].runs[0]
    cell2_run.font.size = Pt(16)

    # Add address
    cell3 = table.cell(1, 0)
    cell3.text = "Address:"
    cell3_run = cell3.paragraphs[0].runs[0]
    cell3_run.font.size = Pt(16)

    cell4 = table.cell(1, 1)
    cell4.text = address
    cell4_run = cell4.paragraphs[0].runs[0]
    cell4_run.font.size = Pt(16)

    # Add phone
    cell5 = table.cell(2, 0)
    cell5.text = "Customer Phone:"
    cell5_run = cell5.paragraphs[0].runs[0]
    cell5_run.font.size = Pt(16)

    cell6 = table.cell(2, 1)
    cell6.text = phone
    cell6_run = cell6.paragraphs[0].runs[0]
    cell6_run.font.size = Pt(16)

    # Add product name
    cell7 = table.cell(3, 0)
    cell7.text = "Product Name:"
    cell7_run = cell7.paragraphs[0].runs[0]
    cell7_run.font.size = Pt(16)

    cell8 = table.cell(3, 1)
    cell8.text = product_name
    cell8_run = cell8.paragraphs[0].runs[0]
    cell8_run.font.size = Pt(16)



        # Add product name
    cell10 = table.cell(5, 0)
    cell10.text = "FROM"
    cell10_run = cell10.paragraphs[0].runs[0]
    cell10_run.font.size = Pt(16)

            # Add product name
    cell11 = table.cell(6, 0)
    cell11.text = "Krish Accessories"
    cell11_run = cell11.paragraphs[0].runs[0]
    cell11_run.font.size = Pt(16)

    byte_stream = BytesIO()
    doc.save(byte_stream)  # Save the document into the BytesIO stream
    byte_stream.seek(0)
    return byte_stream
# Insert page break before the next customer if needed
    if i < len(Prepaid_order) - 1:
        doc.add_page_break()

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data
# Title of the app
st.title("Excel File Uploader")

# Header for single file upload
st.header('Single File Upload')

# Initialize session state for download
if 'downloaded' not in st.session_state:
    st.session_state.downloaded = False

# File uploader for Excel files
data = st.file_uploader("Upload your Excel file here:", type=['xlsx'])

if data is not None and not st.session_state.downloaded:
    # Read the uploaded Excel file
    try:
        order = pd.read_excel(data, sheet_name='Orders')
        inquiry = pd.read_excel(data, sheet_name='Inquiries')
        st.success("File uploaded and read successfully!")

        true_order = order[
    (order['Confirmed Order'].astype(str).str.lower() == 'true') &
    (order['Order Status'].astype(str).str.upper() == 'COMPLETED')
]
        true_order["Payment Mode"] = true_order["Payment Mode"].astype(str).str.upper()
        true_order.loc[true_order["Payment Mode"].str.upper() == "RAZORPAY", "Payment Method"] = "Prepaid"
        true_order.loc[true_order["Payment Mode"].str.upper() == "PARTIAL COD", "Payment Method"] = "PCOD"
        true_order.loc[true_order["Payment Mode"].astype(str).str.upper() == "CASH ON DELIVERY", "Payment Method"] = "COD"
        true_inquiry = inquiry[
    (inquiry['Confirmed Order'].astype(str).str.lower() == 'true') &
    (inquiry['Order Status'].astype(str).str.upper() == 'COMPLETED')
]
        true_inquiry ['Qty Ordered'] = inquiry['Product Name'] + " "+ "["+ inquiry['Item Count'].astype(str) + "]"
        true_order_number = true_order['Order Number'].astype(str).str.upper()
        filtered_inquiry_df = true_inquiry[true_inquiry['Order Number'].astype(str).str.upper().isin(true_order_number)]

        merged_inquiry_df = filtered_inquiry_df.groupby('Order Number', as_index=False).agg({
    'Product Name': lambda x: ', '.join(x.astype(str)),
    'Qty Ordered': lambda x: ','.join(x.astype(str)),
    'Product Weight': lambda x: x.sum() * 1000  # convert weight to grams
})

        column_names = [
            '*Sale Order Number', '*Pickup Location Name', '*Transport Mode',
            '*Payment Mode', 'COD Amount', '*Customer Name', '*Customer Phone',
            '*Shipping Address Line1', 'Shipping Address Line2', '*Shipping City',
            '*Shipping State', '*Shipping Pincode', '*Item Sku Code',
            '*Item Sku Name', '*Quantity Ordered', 'Packaging Type',
            '*Unit Item Price', 'Length (cm)', 'Breadth (cm)', 'Height (cm)',
            'Weight (gm)', 'Fragile Shipment', 'Discount Type', 'Discount Value',
            'Tax Class Code', 'Customer Email',
            'Billing Address same as Shipping Address', 'Billing Address Line1',
            'Billing Address Line2', 'Billing City', 'Billing State',
            'Billing Pincode', 'e-Way Bill Number', 'Seller Name',
            'Seller GST Number', 'Seller Address Line1', 'Seller Address Line2',
            'Seller City', 'Seller State', 'Seller Pincode'
        ]

        upload = pd.DataFrame(columns=column_names)
        new_rows = []

        for index, row in merged_inquiry_df.iterrows():
            group_order = order[order['Order Number'] == row['Order Number']]
            payment_mode = true_order.loc[true_order["Order Number"] == row['Order Number'], "Payment Method"].values[0]

            if not group_order.empty:
                customer_phone = str(group_order['Customer Mobile Number'].values[0]).replace('91', '', 1).strip()
                cod_amount = None
                if (payment_mode =="PCOD"):
                    cod_amount= int(group_order['Total Amount'].values[0]) - 400
                elif ((payment_mode =="COD")):
                    cod_amount= int(group_order['Total Amount'].values[0])









                # Create a new row dictionary
                new_row = {
                    '*Sale Order Number': row['Order Number'],
                    '*Pickup Location Name': 'KRISH ACCESSORIES D2C',
                    '*Transport Mode': "Surface",
                    '*Payment Mode':  payment_mode,
                    'COD Amount': cod_amount,
                    '*Customer Name': group_order['Customer Name'].values[0],
                    '*Customer Phone': customer_phone,
                    '*Shipping Address Line1': group_order['Shipping Address'].values[0],
                    'Shipping Address Line2': "",
                    '*Shipping City': group_order['City'].values[0],
                    '*Shipping State': group_order['State'].values[0],
                    '*Shipping Pincode': group_order['Pincode'].values[0],
                    '*Item Sku Code': row['Product Name'],
                    '*Item Sku Name': row['Qty Ordered'],
                    '*Quantity Ordered': "1",
                    'Packaging Type': "",
                    '*Unit Item Price': group_order['Total Amount'].values[0],
                    'Length (cm)': "10",
                    'Breadth (cm)': "10",
                    'Height (cm)': "10",
                    'Weight (gm)': row['Product Weight'],
                    'Fragile Shipment': "",
                    'Discount Type': "",
                    'Discount Value': "",
                    'Tax Class Code': "",
                    'Customer Email': "",
                    'Billing Address same as Shipping Address': "",
                    'Billing Address Line1': "",
                    'Billing Address Line2': "",
                    'Billing City': "",
                    'Billing State': "",
                    'Billing Pincode': "",
                    'e-Way Bill Number': "",
                    'Seller Name': "",
                    'Seller GST Number': "",
                    'Seller Address Line1': "",
                    'Seller Address Line2': "",
                    'Seller City': "",
                    'Seller State': "",
                    'Seller Pincode': ""
                }
                # Append the new row to new_rows list
                new_rows.append(new_row)

        upload = pd.concat([upload, pd.DataFrame(new_rows)], ignore_index=True)
        doc = Document()
        set_page_size(doc)
        payment_mode = upload.loc[upload["*Payment Mode"].str.upper() == "PREPAID"]
        Prepaid_order=payment_mode[["*Customer Name","*Shipping Address Line1","*Customer Phone","*Item Sku Name","*Sale Order Number"]]
        Prepaid_order = Prepaid_order.rename(columns={
    '*Customer Name': 'Customer_Name',
    '*Sale Order Number':'Order_Number',
    '*Shipping Address Line1': 'Address',
    '*Customer Phone': 'Customer_Phone',
    '*Item Sku Name': 'Product_Name'
})

        for i in range(len(Prepaid_order)):
            row = Prepaid_order.iloc[i]
   # print(row)

    # Correctly reference the columns without extra spaces
            customer_name = row['Customer_Name']
            address = row['Address']  # Make sure column names match exactly
            customer_phone = row['Customer_Phone']
            product_name = row['Product_Name']
            order_number=row['Order_Number']

    # Call the function to add customer details to the document
            word_data=add_customer_details(doc, customer_name, address, customer_phone, product_name,order_number)
            #if (i + 1) % 2 == 0:
            doc.add_page_break()




        # Display the first few rows of the DataFrame
        #st.write(upload.head())

        # Provide a download button
        excel_data = convert_df_to_excel(upload)
        #excel_data = upload.to_excel('upload.xlsx', index=False)

        if st.download_button(
                    label="Download data as Excel",
                    data=excel_data,
                    file_name='Output.xlsx',
                    #mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                ):
                    st.session_state.downloaded = True
                    rerun()





        if st.download_button(
            label="Download data as word",
            data=word_data,
            file_name='Output.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'

        ):
            st.session_state.downloaded = True
            rerun()

    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")

# Reset the state for a new upload
if st.session_state.downloaded:
    st.session_state.downloaded = False
    rerun()

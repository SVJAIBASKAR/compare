import streamlit as st
import pandas as pd
from io import BytesIO

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
        true_inquiry = inquiry[
    (inquiry['Confirmed Order'].astype(str).str.lower() == 'true') &
    (inquiry['Order Status'].astype(str).str.upper() == 'COMPLETED')
]
        true_inquiry ['Qty Ordered'] = inquiry['Product Name'] + " "+ "["+ inquiry['Item Count'].astype(str) + "]"
        true_order_number = true_order['Order Number'].str.lower()
        filtered_inquiry_df = true_inquiry[true_inquiry['Order Number'].str.lower().isin(true_order_number)]

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
            group_order = order[order['Order Number'].astype(str).str.lower() == row['Order Number'].lower()]

            if not group_order.empty:
                customer_phone = str(group_order['Customer Mobile Number'].values[0]).replace('91', '', 1).strip()

                # Create a new row dictionary
                new_row = {
                    '*Sale Order Number': row['Order Number'],
                    '*Pickup Location Name': 'KRISH ACCESSORIES D2C',
                    '*Transport Mode': "Surface",
                    '*Payment Mode': "Prepaid",
                    'COD Amount': "",
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
            st.experimental_rerun()

    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")

# Reset the state for a new upload
if st.session_state.downloaded:
    st.session_state.downloaded = False
    st.experimental_rerun()

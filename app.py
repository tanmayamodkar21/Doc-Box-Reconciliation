import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Set page configuration
st.set_page_config(page_title="🌟 Booking Data Report", layout="wide")

# Custom CSS for styling
st.markdown("""
<style>
    .reportview-container .main .block-container {
        padding: 2rem;
    }
    .css-18e3th9 {
        border: 2px solid #007bff;
        border-radius: 10px;
        padding: 1rem;
        background-color: #f8f9fa;
    }
    .header {
        color: #007bff;
        font-weight: bold;
        font-size: 24px;
        margin-bottom: 20px;
    }
    .section-title {
        color: #007bff;
        font-weight: bold;
        margin-top: 20px;
        margin-bottom: 10px;
        text-align: center;
        font-size: 20px;
    }
    .button {
        background-color: #007bff;
        color: white;
        border-radius: 5px;
        padding: 10px 20px;
    }
    .info {
        margin-bottom: 20px;
        padding: 10px;
        background-color: #e7f3fe;
        border-left: 6px solid #2196F3;
    }
    .summary-table {
        background-color: #34495e;
        color: #ffffff;
        border-radius: 10px;
        margin-top: 20px;
        overflow: hidden; /* To ensure border-radius is effective */
    }
    .summary-table th {
        background-color: #007bff;
        color: #ffffff;
        padding: 12px;
        text-align: left;
    }
    .summary-table td {
        color: #ffffff;
        padding: 10px;
        background-color: #34495e;
    }
    .summary-table tr:nth-child(even) td {
        background-color: #2c3e50; /* Darker for even rows */
    }
    .summary-table tr:hover td {
        background-color: #5d6d7e; /* Highlight on hover */
    }
</style>
""", unsafe_allow_html=True)

st.title("🌟 Booking Data Report")

# Upload and read the Excel files
file1 = st.file_uploader("📂 Upload the first Excel file (Subject & Office)", type=["xlsx"])
file2 = st.file_uploader("📂 Upload the second Excel file (Latest Data)", type=["xlsx"])

if file1 and file2:
    df1 = pd.read_excel(file1, sheet_name="Sheet1")
    df2 = pd.read_excel(file2, sheet_name="Sheet1")

    # Extract booking numbers from the first file using VLOOKUP logic
    booking_count = {'found': 0, 'not_found': 0}

    def extract_booking(subject_line):
        booking_pattern = r"([A-Z]{3}/[A-Z]{3}/\d{7})"
        house_bl_pattern = r"([A-Z]{3}/[A-Z]{3}/\d{5})"

        # Try to extract booking number directly
        booking_match = re.search(booking_pattern, subject_line)
        if booking_match:
            booking_count['found'] += 1
            return booking_match.group(0)

        # If booking number not found, try to extract House BL and look up in df2
        house_bl_match = re.search(house_bl_pattern, subject_line)
        if house_bl_match:
            house_bl = house_bl_match.group(0)
            if 'House BL' in df2.columns and 'Booking' in df2.columns:
                matched_rows = df2[df2['House BL'].str.strip() == house_bl.strip()]
                if not matched_rows.empty:
                    booking_count['found'] += 1
                    return matched_rows.iloc[0]['Booking']

        booking_count['not_found'] += 1
        return None

    # Convert all values in 'Subject' column to strings before applying the function
    df1['Booking Number'] = df1['Subject'].astype(str).apply(extract_booking)

    # Display counts of found and not found booking numbers
    st.markdown("<div class='info'>Total Booking Numbers Found: <strong>{}</strong><br>Total Booking Numbers Not Found: <strong>{}</strong></div>".format(
        booking_count['found'], booking_count['not_found']), unsafe_allow_html=True)

    # Ensure 'Booking' in df2 for merging
    if 'Booking' not in df2.columns:
        st.error("⚠️ 'Booking' column not found in the second Excel file")
    else:
        # Merge the data based on Booking Number
        merged_df = pd.merge(df1, df2, left_on='Booking Number', right_on='Booking', how='left')

        # Strip whitespace from column names
        merged_df.columns = merged_df.columns.str.strip()

        # Fill NaN values with 0
        merged_df['Doc Receive'] = merged_df['Doc Receive'].fillna(0)
        merged_df['Posted'] = merged_df['Posted'].fillna(0)

        # Create office-wise summary
        office_summary = merged_df.groupby('Office_x').agg(
            total_bookings=('Booking', 'count'),
            doc_received=('Doc Receive', lambda x: (x == 1).sum()),
            posted=('Posted', lambda x: (x == 1).sum()),
            pre_release=('Status', lambda x: (x == 'Pre-Release').sum()),
            loaded=('Status', lambda x: (x == 'Loaded').sum())
        ).reset_index()

        # Calculate percentages and format the output
        office_summary['doc_received_formatted'] = office_summary.apply(
            lambda row: f"{row['doc_received']} ({(row['doc_received'] / row['total_bookings']) * 100:.2f}%)", axis=1
        )
        office_summary['posted_formatted'] = office_summary.apply(
            lambda row: f"{row['posted']} ({(row['posted'] / row['total_bookings']) * 100:.2f}%)", axis=1
        )
        office_summary['pre_release_formatted'] = office_summary.apply(
            lambda row: f"{row['pre_release']} ({(row['pre_release'] / row['total_bookings']) * 100:.2f}%)", axis=1
        )
        office_summary['loaded_formatted'] = office_summary.apply(
            lambda row: f"{row['loaded']} ({(row['loaded'] / row['total_bookings']) * 100:.2f}%)", axis=1
        )

        # Calculate totals for all offices
        total_summary = office_summary[['total_bookings', 'doc_received', 'posted', 'pre_release', 'loaded']].sum()
        total_summary_df = pd.DataFrame({
            'Office_x': ['Total'],
            'total_bookings': [total_summary['total_bookings']],
            'doc_received': [total_summary['doc_received']],
            'posted': [total_summary['posted']],
            'pre_release': [total_summary['pre_release']],
            'loaded': [total_summary['loaded']],
        })

        total_summary_df['doc_received_formatted'] = f"{total_summary_df['doc_received'].values[0]} ({(total_summary_df['doc_received'].values[0] / total_summary_df['total_bookings'].values[0]) * 100:.2f}%)"
        total_summary_df['posted_formatted'] = f"{total_summary_df['posted'].values[0]} ({(total_summary_df['posted'].values[0] / total_summary_df['total_bookings'].values[0]) * 100:.2f}%)"
        total_summary_df['pre_release_formatted'] = f"{total_summary_df['pre_release'].values[0]} ({(total_summary_df['pre_release'].values[0] / total_summary_df['total_bookings'].values[0]) * 100:.2f}%)"
        total_summary_df['loaded_formatted'] = f"{total_summary_df['loaded'].values[0]} ({(total_summary_df['loaded'].values[0] / total_summary_df['total_bookings'].values[0]) * 100:.2f}%)"

        # Select and rename columns for display
        office_summary_display = office_summary[[
            'Office_x', 'total_bookings', 'doc_received_formatted', 'posted_formatted', 'pre_release_formatted', 'loaded_formatted'
        ]]
        office_summary_display.columns = [
            'Office', 'Total Bookings', 'Doc Received', 'Posted', 'Pre-Release', 'Loaded'
        ]

        # Append total summary row directly to the existing office-wise summary
        office_summary_display.loc['Total'] = [
            'Total',
            total_summary_df['total_bookings'].values[0],
            total_summary_df['doc_received_formatted'].values[0],
            total_summary_df['posted_formatted'].values[0],
            total_summary_df['pre_release_formatted'].values[0],
            total_summary_df['loaded_formatted'].values[0]
        ]

        # Transpose office-wise summary for better readability
        office_summary_display = office_summary_display.set_index('Office').T.reset_index()

        # Display office-wise summary with enhanced styling
        st.markdown("<div class='section-title'>📊 Office-Wise Summary:</div>", unsafe_allow_html=True)
        st.dataframe(office_summary_display.style.set_table_attributes("class='summary-table'").set_table_styles(
            [{'selector': 'th', 'props': [('font-size', '16px'), ('text-align', 'center')]},
             {'selector': 'td', 'props': [('font-size', '14px'), ('text-align', 'center')]}]
        ))

        # Filtering section
        st.markdown("<div class='section-title'>🔍 Filter Bookings:</div>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([3, 1, 1])  # Create columns for layout

        with col1:
            # Filter by Office with checkboxes
            st.markdown("**Select Office(s):**")
            offices = merged_df['Office_x'].unique().tolist()
            selected_offices = st.multiselect("Options", offices, default=offices)

        with col2:
            # Filter by Doc Received with radio buttons
            st.markdown("**Select Doc Received:**")
            selected_doc_received = st.radio("Options", options=["Both", "Yes", "No"], index=0, key="doc_received_filter")

        with col3:
            # Filter by Posted with radio buttons
            st.markdown("**Select Posted:**")
            selected_posted = st.radio("Options", options=["Both", "Yes", "No"], index=0, key="posted_filter")

        # Apply filters
        filtered_df = merged_df[
            (merged_df['Office_x'].isin(selected_offices)) &
            ((merged_df['Doc Receive'] == 1) if selected_doc_received == "Yes" else True) &
            ((merged_df['Doc Receive'] == 0) if selected_doc_received == "No" else True) &
            ((merged_df['Posted'] == 1) if selected_posted == "Yes" else True) &
            ((merged_df['Posted'] == 0) if selected_posted == "No" else True)
        ]

        # Remove rows where 'Booking Number' is None
        filtered_df = filtered_df[filtered_df['Booking Number'].notna()]

        # Format the ETD column to only show the date
        filtered_df['ETD'] = pd.to_datetime(filtered_df['ETD']).dt.date

        # Display filtered booking numbers
        st.markdown("<div class='section-title'>📅 Filtered Booking Numbers:</div>", unsafe_allow_html=True)
        st.dataframe(filtered_df[['Booking Number', 'Office_x', 'ETD', 'Doc Receive', 'Posted']].style.set_table_attributes("class='summary-table'"))

        # Display counts of bookings by disposition and status
        disposition_counts = filtered_df.groupby('Office_x')['Disposition'].value_counts().unstack(fill_value=0)
        status_counts = filtered_df.groupby('Office_x')['Status'].value_counts().unstack(fill_value=0)

        # Interactive table for Disposition and Status counts side by side
        st.subheader("📊 Bookings by Disposition and Status (Office-wise):")
        for office in disposition_counts.index:
            col1, col2 = st.columns(2)  # Create two columns for disposition and status
            with col1:
                st.markdown(f"### {office} - Disposition")
                for disposition, count in disposition_counts.loc[office].items():
                    if count > 0:  # Only show counts greater than 0
                        if st.button(f"{disposition}: {count}", key=f"disp_{office}_{disposition}"):
                            booking_numbers = filtered_df[(filtered_df['Office_x'] == office) & (filtered_df['Disposition'] == disposition)]['Booking Number']
                            st.write(f"Booking Numbers for {disposition} in {office}:")
                            st.write(booking_numbers.to_list())

            with col2:
                st.markdown(f"### {office} - Status")
                for status, count in status_counts.loc[office].items():
                    if count > 0:  # Only show counts greater than 0
                        if st.button(f"{status}: {count}", key=f"stat_{office}_{status}"):
                            booking_numbers = filtered_df[(filtered_df['Office_x'] == office) & (filtered_df['Status'] == status)]['Booking Number']
                            st.write(f"Booking Numbers for {status} in {office}:")
                            st.write(booking_numbers.to_list())

        # Provide an option to download the filtered bookings data as an Excel file
        @st.cache_data
        def convert_filtered_to_excel(filtered_df):
            output = BytesIO()
            selected_columns = ['Office_x', 'Booking Number', 'ETD', 'Doc Receive', 'Posted', 'Disposition', 'Status']
            filtered_df[selected_columns].to_excel(output, index=False, engine='openpyxl')
            return output.getvalue()

        st.download_button(
            label="⬇️ Download Filtered Bookings as Excel",
            data=convert_filtered_to_excel(filtered_df),
            file_name='filtered_bookings.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key='download_button'
        )

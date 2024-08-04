import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("Booking Data Report")

# Upload and read the Excel files
file1 = st.file_uploader("Upload the first Excel file (Subject & Office)", type=["xlsx"])
file2 = st.file_uploader("Upload the second Excel file (Latest Data)", type=["xlsx"])

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
    st.write(f"Total Booking Numbers Found: {booking_count['found']}")
    st.write(f"Total Booking Numbers Not Found: {booking_count['not_found']}")

    # Ensure 'Booking' in df2 for merging
    if 'Booking' not in df2.columns:
        st.error("'Booking' column not found in the second Excel file")
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

        # Select and rename columns for display
        office_summary_display = office_summary[[
            'Office_x', 'total_bookings', 'doc_received_formatted', 'posted_formatted', 'pre_release_formatted', 'loaded_formatted'
        ]]
        office_summary_display.columns = [
            'Office', 'Total Bookings', 'Doc Received', 'Posted', 'Pre-Release', 'Loaded'
        ]

        # Transpose office-wise summary for better readability
        office_summary_display = office_summary_display.set_index('Office').T.reset_index()

        # Display office-wise summary
        st.write("Office-Wise Summary:")
        st.dataframe(office_summary_display)

        # Sidebar for filtering
        st.sidebar.header("Filter Options")

        # Filter by Office
        offices = merged_df['Office_x'].unique()
        selected_offices = st.sidebar.multiselect("Select Office", offices, default=offices)

        # Filter by Doc Received
        selected_doc_received = st.sidebar.selectbox("Select Doc Received", options=["Both", "Yes", "No"], index=0)

        # Filter by Posted
        selected_posted = st.sidebar.selectbox("Select Posted", options=["Both", "Yes", "No"], index=0)

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
        st.write("Filtered Booking Numbers:")
        st.dataframe(filtered_df[['Booking Number', 'Office_x', 'ETD', 'Doc Receive', 'Posted']])

        # Display counts of bookings by disposition and status
        disposition_counts = filtered_df.groupby('Office_x')['Disposition'].value_counts().unstack(fill_value=0)
        status_counts = filtered_df.groupby('Office_x')['Status'].value_counts().unstack(fill_value=0)

        # Interactive table for Disposition counts
        st.write("Bookings by Disposition (Office-wise):")
        for office in disposition_counts.index:
            st.markdown(f"### {office}")
            for disposition, count in disposition_counts.loc[office].items():
                if count > 0:  # Only show counts greater than 0
                    if st.button(f"{disposition}: {count}", key=f"disp_{office}_{disposition}"):
                        booking_numbers = filtered_df[(filtered_df['Office_x'] == office) & (filtered_df['Disposition'] == disposition)]['Booking Number']
                        st.write(f"Booking Numbers for {disposition} in {office}:")
                        st.write(booking_numbers.to_list())

        # Interactive table for Status counts
        st.write("Bookings by Status (Office-wise):")
        for office in status_counts.index:
            st.markdown(f"### {office}")
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
            label="Download Filtered Bookings as Excel",
            data=convert_filtered_to_excel(filtered_df),
            file_name='filtered_bookings.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

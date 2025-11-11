import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Define the Excel file path
EXCEL_FILE = 'Painter Review v2.xlsx'

def load_data():
    """Loads the Excel file into a DataFrame."""
    try:
        df = pd.read_excel(EXCEL_FILE)
        return df
    except FileNotFoundError:
        st.error(f"Error: The file '{EXCEL_FILE}' was not found. Please ensure it's in the correct directory.")
        return pd.DataFrame()

def save_data(df):
    """Saves the DataFrame back to the Excel file."""
    try:
        df.to_excel(EXCEL_FILE, index=False)
        st.success("Data saved successfully!")
    except Exception as e:
        st.error(f"Error saving data: {e}")

# --- Streamlit App Layout ---
st.set_page_config(page_title="Painter Review Analysis", layout="wide")
st.title("Painter Review Cost Analysis")

# Load initial data
df = load_data()

if not df.empty:
    st.header("1. Average Paint Cost by Room Type")

    # Calculate average cost
    average_cost_by_room_type = df.groupby('RoomType')['Cost'].mean().sort_values(ascending=False)

    st.write("#### Calculated Average Costs:")
    st.dataframe(average_cost_by_room_type)

    # Visualize average cost
    st.write("#### Visualization:")
    fig, ax = plt.subplots(figsize=(12, 6))
    average_cost_by_room_type.plot(kind='bar', color='skyblue', ax=ax)
    ax.set_title('Average Paint Cost by Room Type')
    ax.set_xlabel('Room Type')
    ax.set_ylabel('Average Cost')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    st.pyplot(fig)

    st.markdown("---")

    st.header("2. Add New Quote Data")
    st.write("Use the form below to add new painting quote information.")

    with st.form("new_quote_form", clear_on_submit=True):
        st.subheader("New Quote Details")
        col1, col2 = st.columns(2)
        with col1:
            room_type = st.text_input("Room Type", "New Room")
            cost = st.number_input("Cost", min_value=0.0, value=1000.0)
            doors = st.number_input("Doors", min_value=0, value=1)
        with col2:
            ceiling_inc = st.selectbox("Ceiling Included?", [1.0, 0.0])
            walls_inc = st.selectbox("Walls Included?", [1.0, 0.0])
            windows = st.number_input("Windows", min_value=0, value=1)

        # Add more fields as needed, or simplify
        # For simplicity, other columns might be left as NaN or default values for this example

        submitted = st.form_submit_button("Add Quote")
        if submitted:
            new_data = {
                'RoomType': room_type,
                'Cost': cost,
                'Doors': float(doors),
                'Windows': float(windows),
                'Ceiling Inc': ceiling_inc,
                'Walls Inc': walls_inc,
                'Painter': 'New', # Defaulting for example
                'Quote Number': 'NEW_Q' + str(len(df) + 1),
                # Add other columns with default or dummy values if they are required
                # For a real application, you'd want to ensure all relevant columns are populated
                # and handle missing values gracefully.
            }

            # Convert new_data to a DataFrame and ensure column order matches original df
            new_row_df = pd.DataFrame([new_data], columns=df.columns)
            df = pd.concat([df, new_row_df], ignore_index=True)

            save_data(df)
            st.experimental_rerun() # Rerun to update the charts with new data

else:
    st.warning("Data could not be loaded. Please check the file path and try again.")

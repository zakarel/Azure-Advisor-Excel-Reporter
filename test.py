''' This project is a Python script that generates an Excel workbook report from your Azure Advisor recommendations.'''

import os
from azure.identity import ClientSecretCredential
from azure.mgmt.advisor import AdvisorManagementClient
import pandas as pd
from dotenv import load_dotenv
import datetime

# Constants
CATEGORY_COLUMNS = ['Impact', 'Subscription_ID', 'Resource_Group', 'Type', 'Resource_name', 'Description', 'Links']
OVERVIEW_COLUMNS = ['Category', 'High', 'Medium', 'Low']
RESOURCE_URI_PARTS = 9
MS_SEARCH_URL = "https://learn.microsoft.com/en-us/search/?terms="

# Load environment variables from .env file
def load_env_variables():
    '''This function loads the environment variables from the .env file.'''
    try:
        load_dotenv()
    except Exception as e:
        print(f"Error loading environment variables: {e}")
        return None
    return True

def authenticate_with_azure():
    '''This function authenticates with Azure using the ClientSecretCredential.'''
    try:
        credential = ClientSecretCredential(
            tenant_id=os.environ['AZURE_TENANT_ID'],
            client_id=os.environ['AZURE_CLIENT_ID'],
            client_secret=os.environ['AZURE_CLIENT_SECRET'],
        )
    except Exception as e:
        print(f"Error authenticating with Azure: {e}")
        return None
    return credential


def initialize_advisor_client(credential):
    '''This function initializes the Advisor client.'''
    try:
        advisor_client = AdvisorManagementClient(credential, os.environ['AZURE_SUBSCRIPTION_ID'])
    except Exception as e:
        print(f"Error initializing Advisor client: {e}")
        return False
    return advisor_client

# Fetch the recommendations
def fetch_recommendations(advisor_client):
    '''This function fetches the recommendations from Azure Advisor.'''
    try:
        recommendations = advisor_client.recommendations.list()
    except Exception as e:
        print(f"Error fetching recommendations: {e}")
        recommendations = []
    return recommendations


def process_recommendations(recommendations):
    '''This function processes the recommendations and returns a dataframe for each category and an overview dataframe.'''
    # Initialize an empty dictionary to hold the dataframes for each category
    dfs = {}

    # Initialize a dataframe for the overview
    overview_df = pd.DataFrame(columns=['Category', 'High', 'Medium', 'Low'])

    # Iterate over the recommendations
    for recommendation in recommendations:
        # Get the category
        category = recommendation.category

        # If the category is not in the dictionary, initialize it with an empty dataframe
        if category not in dfs:
            dfs[category] = pd.DataFrame(columns=CATEGORY_COLUMNS)
            new_row = pd.DataFrame([{'Category': category, 'High': 0, 'Medium': 0, 'Low': 0}], columns=['Category', 'High', 'Medium', 'Low'])
            overview_df = pd.concat([overview_df, new_row], ignore_index=True)

        # Extract the resource URI from the recommendation's id
        resource_uri = recommendation.id.split('/providers/Microsoft.Advisor/recommendations/')[0]

        # Split the resource URI into its components
        resource_uri_parts = resource_uri.split('/')
        subscription_id = resource_uri_parts[2] if len(resource_uri_parts) > 2 else None
        resource_group = resource_uri_parts[4] if len(resource_uri_parts) > 4 else None
        resource_type = "/".join(resource_uri_parts[6:8]) if len(resource_uri_parts) > 7 else None
        resource_name = resource_uri_parts[8] if len(resource_uri_parts) > 8 else None

        # Append the recommendation to the appropriate dataframe
        new_row = pd.DataFrame({
            'Impact': [recommendation.impact],
            'Subscription_ID': [subscription_id],
            'Resource_Group': [resource_group],
            'Type': [resource_type],
            'Resource_name': [resource_name],
            'Description': [recommendation.short_description.problem],
            'Links': ['=HYPERLINK("' + MS_SEARCH_URL + recommendation.short_description.problem.replace(' ', '+') + '", "How To")']
        })
        dfs[category] = pd.concat([dfs[category], new_row], ignore_index=True)

        # Update the overview dataframe
        overview_df.loc[overview_df['Category'] == category, recommendation.impact] += 1

        # Calculate the percentages in the overview dataframe
        overview_df['High'] = (overview_df['High'] / (overview_df['High'] + overview_df['Medium'] + overview_df['Low'])) * 100
        overview_df['Medium'] = (overview_df['Medium'] / (overview_df['High'] + overview_df['Medium'] + overview_df['Low'])) * 100
        overview_df['Low'] = (overview_df['Low'] / (overview_df['High'] + overview_df['Medium'] + overview_df['Low'])) * 100

        # Format the 'High', 'Medium', and 'Low' columns as percentages
        styled_overview_df = overview_df.style.format({
            'High': '{:.2f}%',
            'Medium': '{:.2f}%',
            'Low': '{:.2f}%'
        })
        
    return overview_df, dfs

def write_to_excel(overview_df, dfs):
    '''This function writes the dataframes to an Excel file.'''

    # Initialize a pandas Excel writer using 'xlsxwriter' engine and append current date to filename
    current_date = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    writer = pd.ExcelWriter(f'azure-advisor-excel-reporter_{current_date}.xlsx', engine='xlsxwriter')

    # Define a format for the header
    header_format = writer.book.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D9E1F2',  # Change the cell's background to a lighter "blueish" color
        'font_color': 'black',
        'border': 1,
        'font_name': 'Arial'  # Change the font to Arial
    })

    # Define formats for the impact levels
    impact_formats = {
        'High': writer.book.add_format({'bg_color': '#FFEBEB', 'border': 1, 'font_name': 'Arial'}),
        'Medium': writer.book.add_format({'bg_color': '#FFF2CC', 'border': 1, 'font_name': 'Arial'}),
        'Low': writer.book.add_format({'bg_color': '#D9EAD3', 'border': 1, 'font_name': 'Arial'})  # Light green for 'Low' impact
    }

    # Write the overview dataframe to the first sheet in the Excel file
    overview_df.to_excel(writer, sheet_name='Overview', index=False)
    worksheet = writer.sheets['Overview']
    for col_num, value in enumerate(overview_df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    for row_num, row_data in overview_df.iterrows():
        for col_num, value in enumerate(row_data):
            if col_num == 0:  # If this is the 'Category' column
                # Add a hyperlink to the corresponding sheet
                worksheet.write_url(row_num + 1, col_num, f"internal:'{value}'!A1", string=value)
            else:
                # Add the '%' sign to the 'High', 'Medium', and 'Low' columns
                worksheet.write(row_num + 1, col_num, f"{value:.2f}%")

    # Add autofilter
    worksheet.autofilter(0, 0, overview_df.shape[0], overview_df.shape[1] - 1)

    # Create a stacked column chart for all categories and impact
    chart = writer.book.add_chart({'type': 'column', 'subtype': 'stacked'})
    colors = ['#FFEBEB', '#FFF2CC', '#D9EAD3']  # Colors for High, Medium, Low
    for i, impact_level in enumerate(['High', 'Medium', 'Low'], start=1):
        chart.add_series({
            'name': impact_level,
            'categories': ['Overview', 1, 0, overview_df.shape[0], 0],
            'values': ['Overview', 1, i, overview_df.shape[0], i],
            'fill': {'color': colors[i-1]},
        })
    chart.set_title({'name': 'Azure-Advisor-Excel-Reporter by Category (%)'})
    chart.set_x_axis({'name': 'Category'})
    chart.set_y_axis({'name': 'Percentage'})
    worksheet.insert_chart('F2', chart , {'x_scale': 2, 'y_scale': 2})

    # Write each category dataframe to a separate sheet in the Excel file
    for category, df in dfs.items():
        if not df.empty:
            df.to_excel(writer, sheet_name=category, index=False)
            worksheet = writer.sheets[category]
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            for row_num, row_data in df.iterrows():
                for col_num, value in enumerate(row_data):
                    worksheet.write(row_num + 1, col_num, value, impact_formats[row_data['Impact']])
            # Add autofilter
            worksheet.autofilter(0, 0, df.shape[0], df.shape[1] - 1)

    # Save the Excel file
    writer.close()

def main():
    '''This is the main function.'''

    if not load_env_variables():
        print("Error loading environment variables.")
        return
    credential = authenticate_with_azure()
    if not credential:
        print("Error authenticating with Azure.")
        return
    advisor_client = initialize_advisor_client(credential)
    recommendations = fetch_recommendations(advisor_client)
    overview_df, dfs = process_recommendations(recommendations)
    write_to_excel(overview_df, dfs)

if __name__ == "__main__":
    main()
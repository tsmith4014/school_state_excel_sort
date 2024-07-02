import pandas as pd

def extract_unique_states(input_file, output_file):
    # Load the Excel file
    excel_data = pd.ExcelFile(input_file)

    # Load the data from the sheet
    df = pd.read_excel(input_file, sheet_name=excel_data.sheet_names[0])

    # Extracting the school names and student states
    schools = []
    current_school = None

    for index, row in df.iterrows():
        if pd.notna(row[0]) and 'Campus:' in row[0]:
            current_school = row[0].replace('Campus: ', '').strip()
        elif pd.notna(row[5]) and current_school:
            schools.append({'School': current_school, 'State': row[5]})

    # Creating a dataframe from the extracted data
    schools_df = pd.DataFrame(schools)

    # Finding unique states for each school
    unique_states_df = schools_df.groupby('School')['State'].unique().reset_index()
    unique_states_df['State'] = unique_states_df['State'].apply(lambda x: ', '.join(x))

    # Creating a new Excel file with the unique states
    with pd.ExcelWriter(output_file) as writer:
        unique_states_df.to_excel(writer, index=False, sheet_name='Unique States')

if __name__ == "__main__":
    input_file = 'path/to/your/input_file.xlsx'
    output_file = 'path/to/your/output_file.xlsx'
    extract_unique_states(input_file, output_file)
    print(f"Unique states by school have been extracted and saved to {output_file}.")
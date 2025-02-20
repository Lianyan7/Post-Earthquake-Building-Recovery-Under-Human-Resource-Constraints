import pandas as pd


def calculate_damage_ratio():
    """
    This function reads raw building sample data from an Excel file,
    performs one-hot encoding on the 'CapStatus' column, and calculates a
    normalised Damage Ratio within each 'CapStatus_Undercap' group. The updated
    DataFrame is then saved to an output Excel file.
    """
    # File paths (adjust as necessary)
    input_file = 'BuildingSamples.xlsx' # Update your excel document
    output_file = 'BuildingSamples_with_damage_ratio.xlsx'

    # Load the Excel file containing raw building samples data
    df = pd.read_excel(input_file)

    # Perform one-hot encoding on the 'CapStatus' column to create indicator variables
    df = pd.get_dummies(df, columns=['CapStatus'], prefix='CapStatus')

    # Calculate the Damage Ratio by normalising 'Total Building Paid Incl GST'
    # within each group defined by 'CapStatus_Undercap'
    df['Damage Ratio'] = df.groupby('CapStatus_Undercap')['Total Building Paid Incl GST'] \
        .transform(lambda x: (x - x.min()) / (x.max() - x.min()))

    # Save the updated DataFrame with Damage Ratio to a new Excel file
    df.to_excel(output_file, index=False)

    print(f"Damage ratio calculation completed. Results have been saved to {output_file}")
    return df


def prioritise_buildings():
    """
    This function reads experimental building data from an Excel file, checks for
    necessary columns, normalises specific columns (Repair Cost and Policy Preference),
    and computes a PRI index for each building. The resulting DataFrame is sorted by the
    PRI index in descending order and saved to an output Excel file.
    """
    # File path and sheet name (adjust as necessary)
    input_file = 'Data for experiment 1.xlsx'
    sheet_name = 'data'

    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # Verify that the required columns exist in the DataFrame
    required_columns = ['Damage Ratio', 'Repair Cost', 'Importance Level', 'Policy Preference']
    for column in required_columns:
        if column not in df.columns:
            raise ValueError(f"Column '{column}' not found in the DataFrame")

    # Normalise 'Repair Cost' and 'Policy Preference'
    df['Repair Cost_normalized'] = (df['Repair Cost'] - df['Repair Cost'].min()) / \
                                   (df['Repair Cost'].max() - df['Repair Cost'].min())
    df['Policy Preference_normalized'] = (df['Policy Preference'] - df['Policy Preference'].min()) / \
                                         (df['Policy Preference'].max() - df['Policy Preference'].min())

    # Calculate the PRI index, giving equal weight to Damage Ratio, Repair Cost, and Policy Preference
    df['PRI'] = 0.25 * df['Damage Ratio'] + 0.25 * df['Repair Cost_normalized'] + \
                0.25 * df['Policy Preference_normalized']

    # Sort the DataFrame by PRI in descending order so that the highest priority buildings appear first
    df_sorted = df.sort_values(by='PRI', ascending=False)

    # Save the sorted DataFrame to an Excel file
    output_file = 'Option 1-Results_Ranked_Buildings.xlsx'
    df_sorted.to_excel(output_file, sheet_name=sheet_name, index=False)

    print(f"Building prioritisation completed. Sorted data has been saved to {output_file}")
    print(df_sorted)
    return df_sorted


def main():
    """
    Main function that orchestrates the damage score determination and building prioritisation.
    The process consists of two sequential tasks:
      1. Determining the Damage Ratio for each building sample.
      2. Prioritising buildings based on a computed PRI index.
    """
    # Step 1: Calculate Damage Ratio from raw building samples
    calculate_damage_ratio()

    # Step 2: Prioritise buildings using the experimental data
    prioritise_buildings()


if __name__ == '__main__':
    main()

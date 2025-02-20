import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Set Times New Roman as the font for all matplotlib plots
plt.rcParams['font.family'] = 'Times New Roman'


##############################
# STEP 1: DAMAGE RATIO CALCULATION
##############################

def calculate_damage_ratio():
    """
    Reads raw building sample data from an Excel file,
    performs one-hot encoding on the 'CapStatus' column, and calculates a
    normalised Damage Ratio within each 'CapStatus_Undercap' group.
    The updated DataFrame is then saved to an output Excel file.
    """
    # File paths (adjust as necessary)
    input_file = 'BuildingSamples.xlsx'  # Update your Excel document
    output_file = 'BuildingSamples_with_damage_ratio.xlsx'

    # Load the Excel file containing raw building sample data
    df = pd.read_excel(input_file)

    # Perform one-hot encoding on the 'CapStatus' column to create indicator variables
    df = pd.get_dummies(df, columns=['CapStatus'], prefix='CapStatus')

    # Calculate the Damage Ratio by normalising 'Total Building Paid Incl GST'
    # within each group defined by 'CapStatus_Undercap'
    df['Damage Ratio'] = df.groupby('CapStatus_Undercap')['Repair Cost'] \
        .transform(lambda x: (x - x.min()) / (x.max() - x.min()))

    # Save the updated DataFrame with Damage Ratio to a new Excel file
    df.to_excel(output_file, index=False)

    print(f"Damage ratio calculation completed. Results have been saved to {output_file}")
    return df


##############################
# STEP 2: BUILDING PRIORITISATION
##############################

def prioritise_buildings():
    """
    Reads experimental building data from an Excel file, verifies necessary columns,
    normalises specific columns (Repair Cost and Policy Preference), and computes a PRI index.
    The DataFrame is then sorted by the PRI index in descending order and saved to an output Excel file.
    """
    # File path and sheet name (adjust as necessary)
    input_file = 'BuildingSamples_with_damage_ratio.xlsx'
    sheet_name = 'Sheet1'

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

    # Sort the DataFrame by PRI in descending order
    df_sorted = df.sort_values(by='PRI', ascending=False)

    # Save the sorted DataFrame to an Excel file
    output_file = 'Option 1-Results_Ranked_Buildings.xlsx'
    df_sorted.to_excel(output_file, sheet_name=sheet_name, index=False)

    print(f"Building prioritisation completed. Sorted data has been saved to {output_file}")
    print(df_sorted)
    return df_sorted


##############################
# STEP 3: RESOURCE ALLOCATION AND RECOVERY ANALYSIS
##############################

def allocate_resources():
    """
    Reads building ranking data from an Excel file, allocates resources based on
    scenario-specific mobilisation factors, computes waiting and recovery times
    for each building, and writes the results for each scenario to separate sheets
    in an integrated Excel file.
    """
    input_file = 'Option 1-Results_Ranked_Buildings.xlsx'
    input_sheet = 'Sheet1'
    output_file = 'Integrated_Updated_Data_rank_buildings.xlsx'

    # Initial resource pool at time t = 0
    R_0 = 40
  
    # Define scenarios with corresponding mobilisation factors (informed by local context)
    scenarios = {
        'S1': 1,  # Baseline scenario
        'S2': 2,  # Optimistic scenario (Mitigation factor)
        'S3': 0.5  # Pessimistic scenario (Amplification factor)
    }

    def calculate_RM_t(t, factor):
        # This is informed by dynamic resource mobilisation patterns based on New Zealand immigration data.
        return (0.8194 * t - 2.1569) * factor

    # Create an Excel writer to store scenario outputs in one file
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for scenario_name, factor in scenarios.items():
            df = pd.read_excel(input_file, sheet_name=input_sheet)
            t = 0
            R_t = (R_0 + calculate_RM_t(t, factor)) * 0.7  # turnover rate is 0.3

            # Initialise previous building variables
            W_ID_prev = 0
            T_ID_prev = 0
            R_ID_prev = 0

            waiting_times = []
            recovery_times = []

            for index, building in df.iterrows():
                building_ID = building['Building ID']
                R_ID = building['Required Resources']
                RT_ID = building['Repair time']
                W_ID = 0  # initialise current waiting time

                if R_t >= R_ID:  # Sufficient resources available
                    W_ID = 0
                    t += W_ID
                    R_t -= R_ID
                    T_ID = W_ID + RT_ID
                else:  # Insufficient resources available
                    if t > T_ID_prev:
                        # Release resources from previous building
                        R_t += R_ID_prev
                        if R_t >= R_ID:
                            W_ID = W_ID_prev
                            T_ID = RT_ID + W_ID
                            R_t -= R_ID
                            t += W_ID
                        else:
                            t_req_ID = ((R_ID - R_t) + 2.5169 * factor) / (0.8194 * factor)
                            W_ID = t_req_ID + W_ID_prev
                            R_t -= R_ID
                            t = W_ID
                            T_ID = RT_ID + W_ID
                    else:
                        t_req_ID = ((R_ID - R_t) + 2.5169 * factor) / (0.8194 * factor)
                        W_ID = t_req_ID + W_ID_prev
                        R_t -= R_ID
                        t = W_ID
                        T_ID = RT_ID + W_ID

                waiting_times.append(W_ID)
                recovery_times.append(T_ID)

                # Update previous building values for the next iteration
                R_ID_prev = R_ID
                W_ID_prev = W_ID
                T_ID_prev = T_ID

                print(f"Allocate resources to Building ID {building_ID} under {scenario_name}:")
                print(f"At time {t}:")
                print(f"Waiting time (W_ID): {W_ID} days")
                print(f"Recovery time (T_ID): {T_ID} days\n")

            df['Waiting Time'] = waiting_times
            df['Recovery Time'] = recovery_times

            # Create a Rank column if one does not exist (assume the order reflects PRI order)
            if 'Rank' not in df.columns:
                df['Rank'] = np.arange(1, len(df) + 1)

            df.to_excel(writer, sheet_name=scenario_name, index=False)

    print("Resource allocation results have been saved to the integrated Excel file for all scenarios.")
    return output_file


##############################
# STEP 4: VISUALISATION
##############################

def plot_gantt_charts():
    """
    Reads the integrated resource allocation results and generates horizontal Gantt charts
    for building repair and waiting times for each scenario.
    """
    input_file = 'Integrated_Updated_Data_rank_buildings.xlsx'
    sheet_names = ['S1', 'S2', 'S3']
    dataframes = [pd.read_excel(input_file, sheet_name=sheet) for sheet in sheet_names]

    fig, axes = plt.subplots(nrows=1, ncols=3, figsize=(24, 10))
    subplot_labels = ['(a)', '(b)', '(c)']

    for idx, (ax, dataframe, sheet_name) in enumerate(zip(axes, dataframes, sheet_names)):
        if 'Rank' in dataframe.columns:
            dataframe = dataframe.sort_values('Rank')
        else:
            dataframe['Rank'] = np.arange(1, len(dataframe) + 1)

        for index, row in dataframe.iterrows():
            ax.barh(row['Rank'], row['Repair time'], left=row['Waiting Time'],
                    color='orange', edgecolor='black',
                    label='Repair time' if index == 0 else "")
            ax.barh(row['Rank'], row['Waiting Time'], color='grey', edgecolor='black',
                    label='Waiting time' if index == 0 else "")

        ax.set_xlabel('Time (days)', fontsize=22)
        ax.set_title(f'{subplot_labels[idx]} Building Repair and Recovery Gantt Chart - {sheet_name}', fontsize=22,
                     pad=20)

        if idx == 0:
            ax.legend(loc='lower right', fontsize=18)
            ax.set_ylabel('Building ID by PRI order', fontsize=22)
            ax.set_yticks(dataframe['Rank'])
            ax.set_yticklabels(dataframe['Building ID'], fontsize=21)
        else:
            ax.set_yticks([])

        ax.tick_params(axis='x', labelsize=21)

    plt.tight_layout()
    plt.show()


def plot_recovery_ecdf(integrated_file):
    """
    For each scenario, sorts the building recovery times by rank, then plots the recovery trajectory
    using an ECDF-based method. The area under the ECDF curve (enclosed by y=0, y=1 and a fixed t value)
    is computed using a step ECDF method.
    """
    sheet_names = ['S1', 'S2', 'S3']
    plt.figure(figsize=(12, 8))

    colors = ['orange', 'blue', 'green']
    labels = ['S1', 'S2', 'S3']
    areas = []

    # Define ECDF functions as provided
    def ecdf(data):
        n = len(data)
        x = np.sort(data)
        y = np.arange(1, n + 1) / n
        return x, y

    def step_ecdf(x, y):
        x_step = np.repeat(x, 2)[1:]
        y_step = np.repeat(y, 2)[:-1]
        x_step = np.insert(x_step, 0, x[0])
        y_step = np.insert(y_step, 0, 0)
        return x_step, y_step

    # Process each scenario sheet
    for sheet, color, label in zip(sheet_names, colors, labels):
        df = pd.read_excel(integrated_file, sheet_name=sheet)
        # Ensure ordering by Rank
        if 'Rank' in df.columns:
            df = df.sort_values('Rank')
        else:
            df['Rank'] = np.arange(1, len(df) + 1)

        # Use the 'Recovery Time' column as the data for ECDF
        recovery_time = df['Recovery Time'].values

        # Compute the ECDF and then create the step representation
        x, y = ecdf(recovery_time)
        x_step, y_step = step_ecdf(x, y)

        plt.plot(x_step, y_step, color=color, label=label, linewidth=3)

        # Calculate area enclosed by the ECDF curve, the x-axis, and vertical lines at x=0 and x=t_max.
        # The provided method uses t_max = 5152; adjust if necessary.
        t_max = 5152  # This is informed by the max values of the recovery time.
        x_closed = np.concatenate(([0], x_step, [t_max, t_max, 0]))
        y_closed = np.concatenate(([0], y_step, [1, 0, 0]))
        area = np.trapz(y_closed, x_closed)
        areas.append(area)

    plt.xlabel('Recovery Time (days)', fontsize=22)
    plt.ylabel('Recovery Level', fontsize=22)
    plt.xticks(fontsize=18)
    plt.yticks(fontsize=18)
    plt.legend(loc='lower right', fontsize=18)

    # Display computed area values on the plot
    area_text = '\n'.join([f'{label} Area: {area:.2f}' for label, area in zip(labels, areas)])
    plt.text(0.05, 0.95, area_text, transform=plt.gca().transAxes, fontsize=19,
             verticalalignment='top', bbox=dict(facecolor='white', alpha=0.5))

    plt.grid(True)
    plt.tight_layout()
    plt.show()


##############################
# MAIN WORKFLOW
##############################

def main():
    """
    Orchestrates the integrated workflow:
      1. Compute Damage Ratio for raw building samples.
      2. Prioritise buildings based on the computed PRI index.
      3. Allocate resources and compute waiting and recovery times.
      4. Generate Gantt charts for building repair and recovery schedules.
      5. Plot the recovery trajectory using an ECDF-based method.
    """
    # Step 1: Calculate Damage Ratio (if needed, this output can be used for further analysis)
    calculate_damage_ratio()

    # Step 2: Prioritise buildings using experimental data (assumes the Damage Ratio is included)
    prioritise_buildings()

    # Step 3: Allocate resources and compute recovery times (uses output from prioritisation)
    integrated_file = allocate_resources()

    # Step 4: Generate Gantt charts
    plot_gantt_charts()

    # Step 5: Plot recovery trajectory ECDF and compute area under the curve
    plot_recovery_ecdf(integrated_file)


if __name__ == '__main__':
    main()


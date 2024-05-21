import pandas as pd
import matplotlib.pyplot as plt
import os

# Function to check if all required columns exist in the DataFrame
def check_columns_existence(dataframe, columns):
    for col in columns:
        if col not in dataframe.columns:
            print(f"Error: Column '{col}' not found in the DataFrame.")
            return False
    return True

# Load the Excel file
file_path = input("Insert file path:")
dataframe = pd.read_excel(file_path)

# Series of data we want to plot
first_series = input("Enter the name of the first series you want to plot:")
second_series = input("Enter the name of the second series you want to plot:")

# Check if the columns exist before using them
columns_to_check = ['Image', 'Nucleus: Single band FR K Tania mean', 'Nucleus: Single band Green K Tania mean']
if not check_columns_existence(dataframe, columns_to_check):
    exit()

# Check if the entered series names exist in the 'Image' column
unique_images = dataframe['Image'].unique()
if first_series not in unique_images or second_series not in unique_images:
    print("Error: One or both series names not found in the 'Image' column.")
    exit()

# Set graph name
graph_name = f"Scatter plot of Edu and P21 nuclear intensity\n{first_series} vs {second_series} transfected cells"

# Selecting columns based on certain column names and associated data
selected_columns = dataframe[dataframe['Image'] == first_series][['Image', 'Nucleus: Single band FR K Tania mean', 'Nucleus: Single band Green K Tania mean']]
selected_columns2 = dataframe[dataframe['Image'] == second_series][['Image', 'Nucleus: Single band FR K Tania mean', 'Nucleus: Single band Green K Tania mean']]

# Define thresholds
threshold_edu = 200
threshold_p21 = 190

# Calculate median value for P21 positive cells
median_p21_positive_cells = selected_columns[selected_columns['Nucleus: Single band FR K Tania mean'] > threshold_p21]['Nucleus: Single band FR K Tania mean'].median()

# Plotting the data
plt.figure(figsize=(8, 6))  # Set the size of the figure

# Convert columns to numeric data type
dataframe['Nucleus: Single band FR K Tania mean'] = pd.to_numeric(dataframe['Nucleus: Single band FR K Tania mean'], errors='coerce')
dataframe['Nucleus: Single band Green K Tania mean'] = pd.to_numeric(dataframe['Nucleus: Single band Green K Tania mean'], errors='coerce')

# Now find the maximum value
max_value = max(dataframe['Nucleus: Single band FR K Tania mean'].dropna().max(), dataframe['Nucleus: Single band Green K Tania mean'].dropna().max())


# Set the same scale and maximum values for both axes
max_value = max(dataframe['Nucleus: Single band FR K Tania mean'].max(), dataframe['Nucleus: Single band Green K Tania mean'].max())
plt.xlim(0, max_value)
plt.ylim(0, max_value)

# Plotting data from selected_columns in blue color
plt.scatter(selected_columns['Nucleus: Single band Green K Tania mean'],
            selected_columns['Nucleus: Single band FR K Tania mean'],
            color='blue', s=20, label=first_series)

# Plotting data from selected_columns2 in red color
plt.scatter(selected_columns2['Nucleus: Single band Green K Tania mean'],
            selected_columns2['Nucleus: Single band FR K Tania mean'],
            color='orange', s=20, label=second_series)

# Adding labels, title, and legend
plt.xlabel('EdU staining nuclear intensity')
plt.ylabel('P21 staining nuclear intensity')
plt.title(graph_name)
plt.legend()  # Show legend

# Adding thresholds
plt.axhline(y=threshold_p21, color='red', linestyle='--', label=f'P21 Threshold ({threshold_p21})')
plt.axvline(x=threshold_edu, color='green', linestyle='--', label=f'Edu Threshold ({threshold_edu})')

# Create quadrants
plt.axhline(y=threshold_p21, color='gray', linestyle='--', alpha=0.5)
plt.axvline(x=threshold_edu, color='gray', linestyle='--', alpha=0.5)

# Calculate percentages for the first series
total_cells_first_series = len(selected_columns)
quadrant1_first_series = len(selected_columns[(selected_columns['Nucleus: Single band Green K Tania mean'] > threshold_edu) & (selected_columns['Nucleus: Single band FR K Tania mean'] > threshold_p21)])
quadrant2_first_series = len(selected_columns[(selected_columns['Nucleus: Single band Green K Tania mean'] > threshold_edu) & (selected_columns['Nucleus: Single band FR K Tania mean'] <= threshold_p21)])
quadrant3_first_series = len(selected_columns[(selected_columns['Nucleus: Single band Green K Tania mean'] <= threshold_edu) & (selected_columns['Nucleus: Single band FR K Tania mean'] > threshold_p21)])
quadrant4_first_series = len(selected_columns[(selected_columns['Nucleus: Single band Green K Tania mean'] <= threshold_edu) & (selected_columns['Nucleus: Single band FR K Tania mean'] <= threshold_p21)])
percentage_quadrant1_first_series = (quadrant1_first_series / total_cells_first_series) * 100
percentage_quadrant2_first_series = (quadrant2_first_series / total_cells_first_series) * 100
percentage_quadrant3_first_series = (quadrant3_first_series / total_cells_first_series) * 100
percentage_quadrant4_first_series = (quadrant4_first_series / total_cells_first_series) * 100

# Add quadrant annotations for the first series
plt.text(0.4, 0.3, f"Q3: {percentage_quadrant1_first_series:.2f}%", transform=plt.gcf().transFigure, fontsize=10,
         color='blue', bbox=dict(facecolor='white', alpha=0.9, edgecolor='white', boxstyle='round'))
plt.text(0.4, 0.13, f"Q4: {percentage_quadrant2_first_series:.2f}%", transform=plt.gcf().transFigure, fontsize=10,
         color='blue', bbox=dict(facecolor='white', alpha=0.9, edgecolor='white', boxstyle='round'))
plt.text(0.14, 0.3, f"Q1:\n{percentage_quadrant3_first_series:.2f}%", transform=plt.gcf().transFigure, fontsize=10,
         color='blue', bbox=dict(facecolor='white', alpha=0.9, edgecolor='white', boxstyle='round'))
plt.text(0.14, 0.13, f"Q2: {percentage_quadrant4_first_series:.2f}%", transform=plt.gcf().transFigure, fontsize=10,
         color='blue', bbox=dict(facecolor='white', alpha=0.9, edgecolor='white', boxstyle='round'))

# Count points in each quadrant for the second series
total_cells_second_series = len(selected_columns2)
quadrant1_second_series = len(selected_columns2[(selected_columns2['Nucleus: Single band Green K Tania mean'] > threshold_edu) & (selected_columns2['Nucleus: Single band FR K Tania mean'] > threshold_p21)])
quadrant2_second_series = len(selected_columns2[(selected_columns2['Nucleus: Single band Green K Tania mean'] > threshold_edu) & (selected_columns2['Nucleus: Single band FR K Tania mean'] <= threshold_p21)])
quadrant3_second_series = len(selected_columns2[(selected_columns2['Nucleus: Single band Green K Tania mean'] <= threshold_edu) & (selected_columns2['Nucleus: Single band FR K Tania mean'] > threshold_p21)])
quadrant4_second_series = len(selected_columns2[(selected_columns2['Nucleus: Single band Green K Tania mean'] <= threshold_edu) & (selected_columns2['Nucleus: Single band FR K Tania mean'] <= threshold_p21)])
percentage_quadrant1_second_series = (quadrant1_second_series / total_cells_second_series) * 100
percentage_quadrant2_second_series = (quadrant2_second_series / total_cells_second_series) * 100
percentage_quadrant3_second_series = (quadrant3_second_series / total_cells_second_series) * 100
percentage_quadrant4_second_series = (quadrant4_second_series / total_cells_second_series) * 100

# Add quadrant annotations for the second series
plt.text(0.4, 0.25, f"Q3: {percentage_quadrant1_second_series:.2f}%", transform=plt.gcf().transFigure, fontsize=10,
         color='orange', bbox=dict(facecolor='white', alpha=0.9, edgecolor='white', boxstyle='round'))
plt.text(0.4, 0.08, f"Q4: {percentage_quadrant2_second_series:.2f}%", transform=plt.gcf().transFigure, fontsize=10,
         color='orange', bbox=dict(facecolor='white', alpha=0.9, edgecolor='white', boxstyle='round'))
plt.text(0.14, 0.25, f"Q1:\n{percentage_quadrant3_second_series:.2f}%", transform=plt.gcf().transFigure, fontsize=10,
         color='orange', bbox=dict(facecolor='white', alpha=0.9, edgecolor='white', boxstyle='round'))
plt.text(0.14, 0.08, f"Q2: {percentage_quadrant4_second_series:.2f}%", transform=plt.gcf().transFigure, fontsize=10,
         color='orange', bbox=dict(facecolor='white', alpha=0.9, edgecolor='white', boxstyle='round'))

plt.legend()  # Show legend

# Show the plot
plt.show()

# Define the directory where you want to save the files
output_directory = input("Enter the directory path where you want to save the files:")

# Define file names with series names
file_names = [f"{file_name}_{first_series}_{second_series}.xlsx" for file_name in ['Q1_data', 'Q2_data', 'Q3_data', 'Q4_data']]

# Define file paths
file_paths = [os.path.join(output_directory, file_name) for file_name in file_names]

# Define dataframes for each quadrant
quadrant_dataframes = [
    selected_columns[(selected_columns['Nucleus: Single band Green K Tania mean'] <= threshold_edu) &
                     (selected_columns['Nucleus: Single band FR K Tania mean'] <= threshold_p21)],
    selected_columns[(selected_columns['Nucleus: Single band Green K Tania mean'] > threshold_edu) &
                     (selected_columns['Nucleus: Single band FR K Tania mean'] <= threshold_p21)],
    selected_columns[(selected_columns['Nucleus: Single band Green K Tania mean'] <= threshold_edu) &
                     (selected_columns['Nucleus: Single band FR K Tania mean'] > threshold_p21)],
    selected_columns[(selected_columns['Nucleus: Single band Green K Tania mean'] > threshold_edu) &
                     (selected_columns['Nucleus: Single band FR K Tania mean'] > threshold_p21)]
]

# Export dataframes to Excel files
for i, dataframe in enumerate(quadrant_dataframes):
    file_path = file_paths[i]
    dataframe.to_excel(file_path, index=False)
    print(f"Data for Quadrant {i+1} has been exported to {file_path}")


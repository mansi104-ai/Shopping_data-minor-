import pandas as pd
import matplotlib.pyplot as plt
import tensorflow as tf
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image


# Load the CSV file
file_path = 'Raw_Data.csv'  # Replace with the actual file path
df = pd.read_csv(file_path)

# Data Preprocessing

# Drop rows with missing values
df.dropna(inplace=True)

# Convert 'Age' column to numeric (assuming it contains numerical values)
df['Age'] = pd.to_numeric(df['Age'], errors='coerce')

# Convert 'Amount' column to numeric
df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')

# Convert 'Quantity' column to numeric
df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce')

# Data Refining

# Remove duplicate rows
df.drop_duplicates(inplace=True)

# Remove rows with negative values in the 'Quantity' column
df = df[df['Qty'] >= 0]

# Remove rows with negative values in the 'Amount' column
df = df[df['Amount'] >= 0]

# Assuming 'Gender' column has categorical values, you can convert it to lowercase
df['Gender'] = df['Gender'].str.lower()

# Save the processed data to a new CSV file
output_file_path = 'Processed_Data.csv'  # Replace with the desired output file path and extension
df.to_csv(output_file_path, index=False)

print("Data preprocessing and refining completed. Processed data saved to", output_file_path)

# Create a TensorFlow model for generating graphs
class GraphGenerationModel(tf.Module):
    def __init__(self):
        self.file_path = 'Raw_Data.csv'  # Replace with the actual file path
        self.df = pd.read_csv(self.file_path)

    def preprocess_data(self):
        # Data Preprocessing for 'Amount' and 'Gender' columns
        self.df[['Amount', 'Gender']] = self.df[['Amount', 'Gender']].apply(lambda x: x.str.strip() if x.dtype == "O" else x)
        self.df[['Amount', 'Gender']] = self.df[['Amount', 'Gender']].apply(lambda x: x.str.lower() if x.dtype == "O" else x)
        self.df['Amount'] = pd.to_numeric(self.df['Amount'], errors='coerce')

        # Filter rows with valid amounts
        self.df_filtered = self.df.dropna(subset=['Amount', 'Gender'])

    def generate_pie_chart_gender(self):
        # Calculate the total amount for men and women
        total_amount_by_gender = self.df_filtered.groupby('Gender')['Amount'].sum()

        # Manually provided total amounts (replace with actual calculated values)
        total_amount_men = 7613604
        total_amount_women = 13562773

        # Pie Chart
        labels = ['Men', 'Women']
        sizes = [total_amount_men, total_amount_women]
        colors = ['blue', 'lightcoral']

        plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
        plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

        plt.title('Total Amount of Purchases by Gender')

        # Save the pie chart as an image
        image_path = 'Men_vs_Women.png'
        plt.savefig(image_path)

        # Show the pie chart (optional)
        plt.show()

        print("Pie chart saved as 'Men_vs_Women.png'")
        return image_path

    def generate_pie_chart_status(self):
        # Specify the 'Order Status' values of interest
        order_status_values = ['Cancelled', 'Delivered', 'Refunded', 'Returned']

        # Filter data for the specified 'Order Status' values
        filtered_data = self.df[self.df['Status'].isin(order_status_values)]

        # Group by 'Order Status' and calculate the number of orders for each status
        order_status_counts = filtered_data['Status'].value_counts()

        # Generate pie chart with legend
        labels = order_status_counts.index
        values = order_status_counts.values

        plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=90, pctdistance=0.85)
        plt.title('Distribution of Orders by Order Status')

        # Add legend on the right-hand side
        plt.legend(title='Status', bbox_to_anchor=(1, 0.5), loc="center left", borderaxespad=0.)

        # Save the pie chart as an image
        image_path = 'Status.png'
        plt.savefig(image_path)

        # Show the pie chart (optional)
        plt.show()

        print("Pie chart saved as 'Status.png'")
        return image_path

    def generate_bar_chart_top_states(self):
        # Group by 'ship state' and calculate the sum of amounts for each state
        state_amounts = self.df.groupby('ship-state')['Amount'].sum()

        # Find the top 5 states with the maximum sum of amounts
        top_5_states = state_amounts.nlargest(5)

        # Generate horizontal bar graph
        fig, ax = plt.subplots()

        top_5_states.plot(kind='barh', ax=ax, color='blue')

        # Add labels and title
        ax.set_xlabel('Sum of Amount')
        ax.set_ylabel('State')
        ax.set_title('Sales: Top 5 States')

        # Save the bar graph as an image
        image_path = 'Top_5_States.png'
        plt.savefig(image_path)

        # Show the bar graph (optional)
        plt.show()

        print("Top 5 states with the maximum sum of amounts:", top_5_states)
        print("Bar graph saved as 'Top_5_States.png'")
        return image_path

    def generate_pie_chart_channels(self):
        # Data Preprocessing for 'Channels' column
        self.df['Channel '] = self.df['Channel '].str.strip().str.lower()

        # Filter out rows with missing or empty 'Channels' values
        self.df = self.df.dropna(subset=['Channel '])

        # Count the number of orders for each channel
        channel_counts = self.df['Channel '].value_counts()

        # Calculate percentages
        channel_percentages = channel_counts / channel_counts.sum() * 100

        # Generate a list of distinct colors for each division
        colors = ['red', 'orange', 'lightcoral', 'skyblue', 'limegreen', 'mediumorchid']  # Add more colors as needed

        # Generate pie chart
        labels = channel_percentages.index
        sizes = channel_percentages.values

        plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=colors)
        plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

        plt.title('Percentage of Orders from Each Channel')
        plt.show()

    def generate_all_graphs(self):

        # Create an Excel workbook
        workbook = Workbook()

    # Add preprocessed data to the first sheet
        sheet = workbook.active
        for record in dataframe_to_rows(self.df, index=False, header=True):
            sheet.append(record)

    # Create individual sheets for each graph
        sheet_names = ['PieChart_Gender', 'PieChart_Status', 'BarChart_TopStates', 'PieChart_Channels']
        image_paths = [
        self.generate_pie_chart_gender(),
        self.generate_pie_chart_status(),
        self.generate_bar_chart_top_states(),
        self.generate_pie_chart_channels()
    ]

        for name, path in zip(sheet_names, image_paths):
            sheet = workbook.create_sheet(title=name)
        
        # Add the image to the sheet
            img = Image(path)
            sheet.add_image(img, 'A1')

    # Save the workbook
        workbook.save('Graphs_and_Data.xlsx')


# Instantiate the model
model = GraphGenerationModel()

# Preprocess data
model.preprocess_data()

# Generate all graphs and save to Excel workbook
model.generate_all_graphs()

import os
import pandas as pd

folder_path = "ur path goes here"
output_folder = 'output pat'
name_prefix = "network"

# Create the output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Loop through all files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith(".txt"):
        # Your existing code to read the data from the text file
        file_path = os.path.join(folder_path, file_name)

        name = file_name

        results = []  # List to store the results for each line
        with open(file_path, 'r') as file:    
            for line in file:
                results.append(line.split())
                break

        with open(file_path, 'r') as file:
            next(file)
            for line in file:
                output = []  # List to store the output for the current line
                network = line[:(line.rfind(")")) + 1]
                output.append(network)
                rest = line[(line.rfind(")")) + 1:].split()
                output.extend(rest)
                results.append(output)

        # Convert the list of lists to a DataFrame
        df = pd.DataFrame(results)

        # Create a Pandas Excel writer using XlsxWriter as the engine
        excel_file_path = os.path.join(output_folder, f'output_{name}.xlsx')
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, header=False, sheet_name="data")

            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets["data"]

            # Set the column width for each column
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).apply(len).max(), len(str(col))) + 2
                worksheet.set_column(i, i, max_len)

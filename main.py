import pandas as pd
import json 
# # # Read Excel file into pandas DataFrame using openpyxl engine
df = pd.read_excel('convert.xlsx', engine='openpyxl')
df.to_csv("output_csv_file.csv", index=False)


# Read CSV into a DataFrame
df = pd.read_csv("output_csv_file.csv")

# Convert DataFrame to 2D list
array_2d = df.values.tolist()

# Initialize variables
json_data = []
data = None
table = []

# Iterate through the rows of the 2D array
for row in array_2d:
    # Check if the first column is not NaN
    if pd.notna(row[0]):
        # If data is not None, append it to json_data
        if data is not None:
            # Append table to data if it's not empty
            if table:
                data["tableData"] = table
                json_data.append(data)
        # Reset data and table for a new criterion
        data = {
            "criterion": row[0],
            "activeButton": row[1],
            "keyIndicator": row[2],
            "tableData": []
        }
        table = []
        temp = {
            "metricNo": row[3],
            "description": row[4],
            "link": "" if pd.isna(row[5]) else row[5]  # Replace NaN link with empty string
        }
        table.append(temp)
    else:
        temp = {
            "metricNo": row[3],
            "link": "" if pd.isna(row[5]) else row[5]  # Replace NaN link with empty string
        }
        # If description is not NaN, append it to table
        if pd.notna(row[4]):
            temp["description"] = row[4]
        table.append(temp)


# Append the last data to json_data
if data is not None:
    if table:
        data["tableData"] = table
        json_data.append(data)

# Write JSON data to file


def parse_data(input_data):
    parsed_data = []

    for entry in input_data:
        parsed_entry = {
            "criterion": entry["criterion"],
            "activeButton": str(entry["activeButton"]),
            "keyIndicator": entry["keyIndicator"],
            "tableData": []
        }

        table_data = []
        current_metric = None
        current_description = []
        
        for item in entry["tableData"]:
            metric = item["metricNo"]
            
            # Check if the metric number is not NaN
            if pd.notna(metric):
                # If there is a current metric and description,
                # add them to the table data
                if current_metric is not None and current_description:
                    parsed_item = {
                        "metricNo": current_metric,
                        "description": current_description
                    }
                    # Add the parsed item to the table data
                    if "link" in item:
                        parsed_item["link"] = item["link"]
                    table_data.append(parsed_item)
                
                # Reset the current metric and description
                current_metric = metric
                current_description = [item["description"]]
            else:
                # Append the description to the current description list
                current_description.append(item["description"])

        # Add the last item to table data if there's any
        if current_metric is not None and current_description:
            parsed_item = {
                "metricNo": current_metric,
                "description": current_description
            }
            if "link" in item:
                parsed_item["link"] = item["link"]
            table_data.append(parsed_item)

        parsed_entry["tableData"] = table_data
        parsed_data.append(parsed_entry)

    return parsed_data
final_data=parse_data(json_data)

with open('data.json', 'w') as json_file:
    json.dump(final_data, json_file)

print("Data successfully written to 'data.json' file.")
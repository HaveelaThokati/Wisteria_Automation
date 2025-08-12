import pandas as pd
import requests
import json

# Excel File path
file_path = "C:/Users/Haveela/Downloads/Response_Generation.xlsx"

# API URL
API_URL = "https://connectedenterprise.vassardigital.ai/wisteria/api/gorgias/chat_response_generation_testing"
API_HEADERS = {"Content-Type": "application/json"}

# Reading the Excel, Test Case, User Query Columns
df = pd.read_excel(file_path)

# Checking for Scenario Title
if df.shape[1] < 3:
    raise ValueError(
        "Excel file must have at least three columns: Column A (Scenario Title), Column B (Test Case), Column C (User Query).")

# Preparing lists to hold output data for columns D (Converted_JSON), E (Result), F (Status_Code)
converted_json_list = []
api_result_list = []
status_code_list = []

# Processing each row:
for idx, row in df.iterrows():
    scenario_title = str(row.iloc[0]).strip()  # Column A
    test_case = row.iloc[1]  # Column B
    user_query = row.iloc[2]  # Column C : the input to API which will be converted into JSON

    if pd.isna(user_query):
        converted_json_list.append("")
        api_result_list.append("")
        status_code_list.append("")
        continue

    # Cleaning the query text - removing 'Query:'
    clean_query = str(user_query).strip()
    if clean_query.lower().startswith("query:"):
        clean_query = clean_query[len("query:"):].strip()

    # Building the JSON payload from Column C (user_query)
    payload = {
        "conversation": [
            {
                "human": "Hi",
                "AI": "Hi there! How can I assist you today?"
            }
        ],
        "query": clean_query
    }

    # Converting payload to JSON string -to save in Excel column D
    json_str = json.dumps(payload, ensure_ascii=False)
    converted_json_list.append(json_str)

    # Calling the API
    try:
        response = requests.post(API_URL, headers=API_HEADERS, data=json_str)
        if response.status_code == 200:
            data = response.json()
            api_result = data.get("result", "")
            api_result_list.append(api_result)
            status_code_list.append(data.get("code", response.status_code))

            print(f"\nScenario: {scenario_title}")
            print("Query:", clean_query)
            print("Bot Response:", api_result)
        else:
            api_result_list.append("")
            status_code_list.append(response.status_code)
            print(f"\nScenario: {scenario_title}")
            print("Query:", clean_query)
            print(f"Bot Response: [No response, Status Code: {response.status_code}]")

    except Exception as e:
        api_result_list.append("")
        status_code_list.append(f"Error: {e}")
        print(f"\nScenario: {scenario_title}")
        print("Query:", clean_query)
        print(f"Bot Response: [Error: {e}]")

# Saving results to Excel:
# Putting the data into columns: D=Converted_JSON, E=Result, F=Status_Code
# Column D
df["Converted_JSON"] = converted_json_list
# Column E
df["Result"] = api_result_list
# Column F
df["Status_Code"] = status_code_list

df.to_excel(file_path, index=False)
print(f"\nAPI calls completed. Results saved in {file_path}")

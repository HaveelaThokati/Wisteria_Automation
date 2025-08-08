import json
import openpyxl
from langchain.chains import LLMChain
from langchain_community.chat_models import AzureChatOpenAI
from langchain.prompts import ChatPromptTemplate, HumanMessagePromptTemplate, SystemMessagePromptTemplate


def initialize_llm():
    return AzureChatOpenAI(
        deployment_name="vassar-4o-mini",
        model_name="gpt-4o-mini",
        azure_endpoint="https://vassar-openai.openai.azure.com/",
        openai_api_key="DWDh8ZelTMEab6XAys0wo0d4p71l2i58QHe8GTiWOwYo5bHzkTisJQQJ99ALACYeBjFXJ3w3AAABACOGJVPB",
        openai_api_version="2024-02-15-preview",
        temperature=0.0
    )


EXTRACTION_RULES = {
"SYSTEM": """You are a detail extraction bot. Your task is to extract SKU ID(s) and product names from the USER CURRENT MESSAGE, refer to HISTORY when necessary.
    a) SKU ID(s) Extraction:
        1. Extract the SKU IDs is found in USER CURRENT MESSAGE.
        2. Identify SKU IDs that follow these patterns:
            - A numerical sequence, an  hyphen, and additional alphanumeric characters (e.g., 123-1010198-NAT, 106-111293-BLUE-King)
            - A numerical sequence followed by an  hyphen and additional alphanumeric characters (e.g., 108-4567890)
        3. If no SKU ID is found, return an empty list [].
    b) Product Name Extraction:
        1. Extract the product name or product names from the USER CURRENT MESSAGE if explicitly mentioned or partially mentioned (even if uncertain or approximate).
        2. if no product name is mentioned, return an empty list [].
    c) When the user mentions a product by its position from a list in the HISTORY, locate the list and extract the corresponding product details.
""",
"CONTEXT": """
    HISTORY:
     {previous_conversations}
    USER CURRENT MESSAGE:
     {question}
    """,
"DISPLAY": """Ensure that the output is in the following JSON format exactly as shown:
    {{
        "SKU_IDs": ["List of SKU IDs"],
        "product_names": ["List of product names"]
    }}
    """,
"REMEMBER": """
    - First, extract details from the USER CURRENT MESSAGE.
    - HISTORY only as a supplementary source for context.
    - Ensure all extracted details are explicitly mentioned."""


}


def extract_order_details(email_history, current_email):
    llm = initialize_llm()
    prompt = ChatPromptTemplate(
        messages=[
            SystemMessagePromptTemplate.from_template(EXTRACTION_RULES["SYSTEM"]),
            HumanMessagePromptTemplate.from_template(EXTRACTION_RULES["CONTEXT"]),
            SystemMessagePromptTemplate.from_template(EXTRACTION_RULES["DISPLAY"]),
            SystemMessagePromptTemplate.from_template(EXTRACTION_RULES["REMEMBER"])
        ]
    )
    llm_chain = LLMChain(llm=llm, prompt=prompt, verbose=True)
    return llm_chain.run({"previous_conversations": email_history, "question": current_email})

def process_excel_data(sheet):
    for row in range(2, sheet.max_row + 1):
        if not sheet[f'A{row}'].value:
            break
        cell_value = sheet[f'B{row}'].value
        if cell_value and "Users_Current_Email:" in cell_value:
            split_values = cell_value.split("Users_Current_Email:")
            email_history = split_values[0].strip() if len(split_values) > 0 else ""
            user_current_email = split_values[1].strip() if len(split_values) > 1 else ""

            if email_history and user_current_email:
                output = extract_order_details(email_history, user_current_email)
                sku_ids, product_names = [], []

                if output:
                    try:
                        parsed_output = json.loads(output)
                        sku_ids = parsed_output.get("SKU_IDs", [])
                        product_names = parsed_output.get("product_names", [])
                    except json.JSONDecodeError:
                        print(f"JSON parsing error in row {row}")

                    print(f"Writing SKU IDs: {sku_ids} and Product Names: {product_names} to row {row}")
                else:
                    print(f"Empty response for row {row}. Email: {user_current_email}")
                    sku_ids, product_names = ["No Output"], ["No Output"]

                sheet[f'C{row}'] = ", ".join(sku_ids)
                sheet[f'D{row}'] = ", ".join(product_names)
            else:
                sheet[f'C{row}'] = "Invalid Input"
                sheet[f'D{row}'] = "Invalid Input"

def process_excel(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        process_excel_data(sheet)
        updated_file_path = file_path.replace(".xlsx", "_updated.xlsx")
        workbook.save(updated_file_path)
        print("Updated file saved at:", updated_file_path)
    except (FileNotFoundError, PermissionError) as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    file_path = "C:/Users/Haveela/Downloads/ID_Extraction_Automation.xlsx"
    process_excel(file_path)

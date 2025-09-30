import pandas as pd
from docx import Document
import os

DATA_DIR = "../data"

def load_excel(file_name):
    path = os.path.join(DATA_DIR, file_name)
    df = pd.read_excel(path)
    return df

def load_word(file_name):
    path = os.path.join(DATA_DIR, file_name)
    doc = Document(path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return paragraphs

def extract_tool_metadata():
    # Load Excel files
    cleanup_df = load_excel("Tool Database Cleanup_as of 081825.xlsx")
    els_df = load_excel("Tools in Online eLS Database.xlsx")

    # Load Word file
    prompt_paragraphs = load_word("AI Tool Page Prompt_082025.docx")

    # Merge and clean data
    tools = {}
    for _, row in cleanup_df.iterrows():
        name = row.get("Tool Name", "").strip()
        if name:
            tools[name] = {
                "Status": row.get("Status", ""),
                "Purpose": row.get("Purpose", "").split(","),
                "Usage": row.get("Usage", "").split(","),
                "Cost": row.get("Cost", ""),
                "Support": row.get("Support", "")
            }

    # Add narrative prompts if available
    for para in prompt_paragraphs:
        for tool in tools:
            if tool.lower() in para.lower():
                tools[tool]["Overview"] = para

    return tools

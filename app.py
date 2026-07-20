import streamlit as st
import pandas as pd
import google.generativeai as genai
from docx import Document
import json
import time
import io
import xlsxwriter

# ==========================================
# 1. PAGE CONFIGURATION & UI
# ==========================================
st.set_page_config(page_title="Privacy Policy Extractor", layout="wide")
st.title("⚖️ Privacy Policy Metadata Extractor")
st.markdown("""
This tool uses AI to strictly extract 16 metadata points from privacy policy `.docx` files. 
It operates on a **Zero-Inference Guarantee** and supports dual-sheet Excel exports (Binary & Verbatim).
""")

# Sidebar for API Key configuration
st.sidebar.header("Configuration")
api_key = st.sidebar.text_input("Enter Google Gemini API Key:", type="password")
st.sidebar.markdown("[Get your free API key here](https://aistudio.google.com/app/apikey)")


# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================
def extract_text_from_docx(file):
    """Reads the uploaded DOCX file and extracts all text rows cleanly."""
    doc = Document(file)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text.strip())
    return "\n".join(full_text)


def analyze_policy(text, api_key):
    """Sends raw text to Gemini API and uses a brace-counting algorithm to extract pure JSON data."""
    genai.configure(api_key=api_key)

    # Utilizing the state-of-the-art 3.5 Flash architecture
    model = genai.GenerativeModel('gemini-3.5-flash')

    system_prompt = """
        You are an expert in legal privacy policy analysis. 
        Analyze the provided privacy policy text and extract the 16 required metadata fields based on the STRICT definitions provided below.

        FIELD DEFINITIONS:
        1. PII: Information that can potentially identify a person (e.g., Name, email address, phone number, SSN).
        2. NonPII: Information that cannot identify a person on its own (e.g., IP address, location, device ID, OS).
        3. NonAffiliatedThirdParty: Whether PII/Non-PII is shared with third parties for purposes like marketing, business, etc.
        4. PurposeDisclosure: Does the notice indicate that the purpose for disclosure to third parties is to facilitate direct marketing, personalized ads, or "for any purpose"?
        5. CategoriesOfInfo: Categories of non-affiliated third parties with whom the information will be shared (e.g., Marketers, Advertisers, Data Analytics, Sponsors).
        6. ChoiceOfOptOut: Whether the user has the choice to opt-out of the sharing of information by the data collector.
        7. ParentConsent: Whether the data collector collects/stores children's personal data only upon parent/legal guardian's consent.
        8. CCPA_Compliance: Whether the data collector is CCPA compliant.
        9. CCPA_PICategories: Description of the categories of personal information collected under CCPA.
        10. CCPA_NonAffiliatedThirdParty: The purpose for which the Data Collector will use the information, specifically including statements that information will be shared with unaffiliated third parties under CCPA.
        11. CCPA_CategoriesThirdParty: Categories of non-affiliated third parties with whom information is shared under CCPA.
        12. CCPA_DataCollectorContact: How the consumer can contact the data collector to learn about collection/disclosure.
        13. CCPA_DoNotSellPILink: Mentions of a "Do not sell my personal information" link.
        14. CCPA_OptOutDisc: Opt-out disclosure in the privacy notice or California-specific description of privacy rights.
        15. CCPA_OptOutLink: Link to a CCPA opt-out page.
        16. GDPR_Compliance: Whether the data collector is GDPR compliant.

        CRITICAL RULES:
        1. Zero-Inference Guarantee: Never infer or assume text that isn't there.
        2. Value Logic (Strict) - YOU MUST BALANCE YES, NO, AND NA:
            - Output "Yes" if the policy explicitly affirms this or provides the required information.
            - Output "No" ONLY if the policy explicitly denies this or states they DO NOT do it (e.g., "We do not sell personal information", "We do not share data with non-affiliated third parties", "We do not knowingly collect data from children").
            - Output "NA" if the topic, regulation, or link is SIMPLY NOT MENTIONED anywhere in the document (e.g., if GDPR or a specific CCPA link is completely missing from the text, it is "NA").
        3. Description logic: 
            - If "Yes" or "No": Copy the EXACT VERBATIM text from the document. Never summarize or paraphrase.
            - If "NA": Description must be exactly "NA".
        4. If evidence exists in multiple locations, copy all relevant text and separate using a pipe "|".
        5. JSON ESCAPING (MANDATORY): You MUST properly escape all double quotes (\\") and newlines (\\n) inside your verbatim descriptions.
    
        Respond ONLY with a valid JSON object matching this exact structure:
        {
            "PII": "Yes/No/NA", "PII_Desc": "...",
            "NonPII": "Yes/No/NA", "NonPII_Desc": "...",
            "NonAffiliatedThirdParty": "Yes/No/NA", "NonAffiliatedThirdParty_Desc": "...",
            "ChoiceOfOptOut": "Yes/No/NA", "ChoiceOfOptOut_Desc": "...",
            "PurposeDisclosure": "Yes/No/NA", "PurposeDisclosure_Desc": "...",
            "CategoriesOfInfo": "Yes/No/NA", "CategoriesOfInfo_Desc": "...",
            "ParentConsent": "Yes/No/NA", "ParentConsent_Desc": "...",
            "CCPA_Compliance": "Yes/No/NA", "CCPA_Compliance_Desc": "...",
            "CCPA_PICategories": "Yes/No/NA", "CCPA_PICategories_Desc": "...",
            "CCPA_NonAffiliatedThirdParty": "Yes/No/NA", "CCPA_NonAffiliatedThirdParty_Desc": "...",
            "CCPA_CategoriesThirdParty": "Yes/No/NA", "CCPA_CategoriesThirdParty_Desc": "...",
            "CCPA_DataCollectorContact": "Yes/No/NA", "CCPA_DataCollectorContact_Desc": "...",
            "CCPA_DoNotSellPILink": "Yes/No/NA", "CCPA_DoNotSellPILink_Desc": "...",
            "CCPA_OptOutDisc": "Yes/No/NA", "CCPA_OptOutDisc_Desc": "...",
            "CCPA_OptOutLink": "Yes/No/NA", "CCPA_OptOutLink_Desc": "...",
            "GDPR_Compliance": "Yes/No/NA", "GDPR_Compliance_Desc": "..."
        }
    """

    response = model.generate_content(
        f"{system_prompt}\n\nPRIVACY POLICY TEXT TO ANALYZE:\n{text}",
        generation_config=genai.GenerationConfig(
            response_mime_type="application/json"
        )
    )

    # --- THE ULTIMATE JSON CLEANER (BRACE-COUNTING ALGORITHM) ---
    raw_response = response.text.strip()

    try:
        return json.loads(raw_response)
    except json.JSONDecodeError:
        pass

    start_idx = raw_response.find('{')
    if start_idx == -1:
        raise ValueError("The AI response did not contain a valid JSON object structure.")

    brace_count = 0
    end_idx = -1

    for i, char in enumerate(raw_response[start_idx:]):
        if char == '{':
            brace_count += 1
        elif char == '}':
            brace_count -= 1
            if brace_count == 0:
                end_idx = start_idx + i
                break

    if end_idx != -1:
        clean_json_str = raw_response[start_idx:end_idx + 1]
        try:
            return json.loads(clean_json_str)
        except json.JSONDecodeError as e:
            raise ValueError(f"AI provided invalid JSON (likely unescaped quotes). Raw error: {str(e)}")
    else:
        raise ValueError("The AI response had unmatched braces and could not be parsed.")


# ==========================================
# 3. MAIN APPLICATION LOGIC
# ==========================================
uploaded_files = st.file_uploader("Upload Privacy Policy Documents (.docx)", type=["docx"], accept_multiple_files=True)

if st.button("Extract Metadata", type="primary"):
    if not api_key:
        st.error("⚠️ Please enter your Gemini API Key in the sidebar.")
    elif not uploaded_files:
        st.warning("⚠️ Please upload at least one .docx file.")
    else:
        results = []
        progress_bar = st.progress(0)

        with st.spinner('Analyzing documents... Anti-spam pacing and auto-retries active.'):
            for i, file in enumerate(uploaded_files):
                try:
                    base_url = file.name.lower().replace(".docx", "")
                    raw_text = extract_text_from_docx(file)

                    # Request API Extraction with AUTOMATIC RETRIES
                    max_retries = 3
                    ai_data = None

                    for attempt in range(max_retries):
                        try:
                            ai_data = analyze_policy(raw_text, api_key)
                            break
                        except Exception as e:
                            if "429" in str(e) or "Quota" in str(e):
                                if attempt < max_retries - 1:
                                    time.sleep(10)
                                else:
                                    raise Exception("Rate limit max retries exceeded. Try again later.")
                            else:
                                raise e

                    row_data = {"BASE_URL": base_url, "URL": f"https://www.{base_url}/privacy-policy"}
                    row_data.update(ai_data)
                    results.append(row_data)

                    # 8-second delay to stay safely under Free Tier RPM limits
                    if i < len(uploaded_files) - 1:
                        time.sleep(8)

                except Exception as e:
                    st.error(f"Error processing {file.name}: {str(e)}")

                progress_bar.progress((i + 1) / len(uploaded_files))

        # ==========================================
        # 4. DATA TRANSFORMATION & EXCEL EXPORT
        # ==========================================
        if results:
            st.success("✅ Extraction Complete!")

            # Create master DataFrame
            df_raw = pd.DataFrame(results)

            # ------------------------------------------
            # Sheet 1: Output-Binary (0s and 1s)
            # ------------------------------------------
            binary_columns = [
                'BASE_URL', 'URL', 'PII', 'NonPII', 'NonAffiliatedThirdParty', 'ChoiceOfOptOut',
                'PurposeDisclosure', 'CategoriesOfInfo', 'ParentConsent', 'CCPA_Compliance',
                'CCPA_PICategories', 'CCPA_NonAffiliatedThirdParty', 'CCPA_CategoriesThirdParty',
                'CCPA_DataCollectorContact', 'CCPA_DoNotSellPILink', 'CCPA_OptOutDisc',
                'CCPA_OptOutLink', 'GDPR_Compliance'
            ]
            df_binary = df_raw[binary_columns].copy()

            # Map Yes->1, No->0, NA->0
            mapping_dict = {'Yes': 1, 'No': 0, 'NA': 0}
            for col in binary_columns[2:]:  # Skip BASE_URL and URL
                df_binary[col] = df_binary[col].map(mapping_dict).fillna(0).astype(int)

            # ------------------------------------------
            # Sheet 2: Output- Yes|No (Verbatim text)
            # ------------------------------------------
            rename_mapping = {
                'PII_Desc': 'Description',
                'NonPII_Desc': 'Description.1',
                'NonAffiliatedThirdParty_Desc': 'Description.2',
                'ChoiceOfOptOut_Desc': 'Description.3',
                'PurposeDisclosure_Desc': 'Description.4',
                'CategoriesOfInfo_Desc': 'Description.5',
                'ParentConsent_Desc': 'Description.6',
                'CCPA_Compliance_Desc': 'Description.7',
                'CCPA_PICategories_Desc': 'Description.8',
                'CCPA_NonAffiliatedThirdParty_Desc': 'Description.9',
                'CCPA_CategoriesThirdParty_Desc': 'Description.10',
                'CCPA_DataCollectorContact_Desc': 'Description.11',
                'CCPA_DoNotSellPILink_Desc': 'Description.12',
                'CCPA_OptOutDisc_Desc': 'Description.13',
                'CCPA_OptOutLink_Desc': 'Description.14',
                'GDPR_Compliance_Desc': 'Description.15'
            }

            # Reorder columns to match Excel pattern exactly (Value, Description, Value, Description...)
            ordered_columns = ['BASE_URL', 'URL']
            for col in binary_columns[2:]:
                ordered_columns.append(col)
                ordered_columns.append(col + '_Desc')

            df_yes_no = df_raw[ordered_columns].rename(columns=rename_mapping)

            # ------------------------------------------
            # Write to Excel Buffer
            # ------------------------------------------
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                df_binary.to_excel(writer, index=False, sheet_name='Output-Binary')
                df_yes_no.to_excel(writer, index=False, sheet_name='Output- Yes|No')

                # Format Binary Sheet
                worksheet1 = writer.sheets['Output-Binary']
                for i, col in enumerate(df_binary.columns):
                    worksheet1.set_column(i, i, max(len(col), 15))

                # Format Yes|No Sheet (Cap description column widths for readability)
                worksheet2 = writer.sheets['Output- Yes|No']
                for i, col in enumerate(df_yes_no.columns):
                    if 'Description' in col:
                        worksheet2.set_column(i, i, 60)  # Wrap wide text columns
                    else:
                        worksheet2.set_column(i, i, max(len(col), 15))

            excel_data = excel_buffer.getvalue()

            # UI Display
            st.markdown("### Preview: Binary Output")
            st.dataframe(df_binary.head())

            st.download_button(
                label="📊 Download Final Excel Report (.xlsx)",
                data=excel_data,
                file_name="Privacy_Policy_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
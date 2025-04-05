import streamlit as st
from PyPDF2 import PdfReader
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="PDF Form Extractor", page_icon="📝")

def extract_form_field_pairs(pdf_file):
    """Extract label-value pairs from fillable PDF form fields"""
    reader = PdfReader(pdf_file)
    fields = reader.get_fields() or {}
    
    results = []
    for field_name, field in fields.items():
        value = field.get('/V', '').strip()
        field_label = field.get('/T', field_name).strip()
        
        if '(' in field_label and ')' in field_label:
            field_label = field_label.split('(')[0].strip()
        
        if value:
            if value.startswith('/'):
                value = value[1:]
            results.append({
                'question': field_label,
                'answer': value
            })
    return results

# Predefined index list
indexes = [
    "0.0", "0.1", "I.1", "I.2", "I.3", "I.4", "I.5", "I.6", "I.6.a", "I.6.b",
    "II.1", "II.2", "II.3", "II.3a", "II.4", "II.5", "II.6", "II.7", "II.8",
    "II.9.1", "II.9.2", "II.10", "II.11", "III.1", "III.1", "III.1.1", "III.2",
    "III.3", "III.4", "3"
]

# Streamlit UI
st.title("📝 PDF Form to Excel Converter")
st.markdown("""
Upload a fillable PDF form to extract all completed fields and download as Excel.
""")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    with st.spinner("Extracting form data..."):
        try:
            # Extract data
            form_data = extract_form_field_pairs(uploaded_file)
            
            if not form_data:
                st.warning("No form fields with values found in this PDF.")
                st.stop()
            
            # Create DataFrame
            df = pd.DataFrame({
                'Index': indexes[:len(form_data)],
                'Question': [item['question'] for item in form_data],
                'Answer': [item['answer'] for item in form_data]
            })
            
            # Show preview
            st.success(f"Extracted {len(form_data)} form fields!")
            st.dataframe(df.head(10))
            
            # Create download link
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='FormData')
            
            st.download_button(
                label="📥 Download Excel File",
                data=output.getvalue(),
                file_name="form_data_extracted.xlsx",
                mime="application/vnd.ms-excel"
            )
            
            # Show stats
            st.info(f"""
            - **Total questions exported**: {len(form_data)}
            - **Indexes used**: {len(indexes)} ({(len(form_data)/len(indexes))*100:.1f}% utilization)
            """)
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
else:
    st.info("Please upload a PDF file to get started")

# Add some instructions
st.markdown("""
### How to use:
1. Upload a fillable PDF form that has been completed
2. Preview and download the extracted data
""")
import streamlit as st
import pandas as pd
import openai
from io import BytesIO

# Set the OpenAI API key securely from Streamlit secrets
openai.api_key = st.secrets["openai"]["api_key"]

st.set_page_config(page_title="Email Reviewer", layout="wide")
st.title("Khotwa Email Reviewer")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Ensure Email column is present
    if 'Email' not in df.columns:
        st.error("The uploaded file must contain an 'Email' column.")
        st.stop()

    # Fill NaNs
    df['Email'] = df['Email'].fillna('')

    # Function to review one email and extract multiple student records
    def review_email(email_text):
        prompt = f"""
You are an email reviewer.

Given the email description:

Email:
{email_text}

TASK:
Your job is to review the email and provide the following columns as an output:

Date of Email - This column will have the first date of the email as provided in the email.
Date of Close of Email Thread - The latest date of the email in the thread.
PS Number - This column will have the PS number of the student mentioned in the email. NA incase of no PS Number
Student Name - This column will have the student name of the student mentioned in the email.
Mentor - This column will highlight the name who has the email address in the format "@mindbase.education".
Issue - This column will highlight the issue in the email. It has to choosen from one among "Housing Issue, Salary Issue, Tuition Issue, TWIMC (To Whoever it may concern), Academic Achievement, National Service, Transcript Not Submitted, Poor academic performance, Inconsistent Communication".
Brief - This column will display the brief summary of the body of the email. The email should begin with "Mindbase mentor, <Name of the Sender>, emaied ADEK Advisor, <Name of the Receiver>," and then followed by a brief summary. The summary should be more than 30 words.
Tier Classification - This column will highlight the concern. It chooses one among "Tier 1: Safety & Behavioral Concerns, Tier 2: Academic Concern, Tier 3: Accommodation Concern, Other Issue".
Sent to - This column will display the name of an ADEK advisor. The email address of an ADEK advisor is in the format "@adek.gov.ae".
Handover Items - This column will highlight the details of the issue in the email. It has to be choosen from one among "Housing Payment, Housing Updates, TWIMC Letter, Pending Tuition Fees, Salary Issue, Nothing".

Please note that for the Handover Items column, if 2 or more issues collide in an email, then the priority will be given to the most important issue. For example, if the student has to be provided with the TWIMC letter for renting the apartment, then Housing Tems for that particular student should be "Housing Updates". 
The Mentor and Sent to columns will only have the names of the person and not their email addresses.

Instructions:
1. Treat the entire email chain as one communication block. Focus on the most recent intent or resolution.
2. If the email mentions multiple students, return one set of details per student.
3. Shared fields like Date of Email, Mentor, and Sent To should remain the same for all rows.
4. Issue, Brief, Tier Classification, and Handover Items must be extracted separately per student.
5. Always return results in the below format:

Date of Email: <value>
Date of Close of Email Thread: <value>  
PS Number: <value>  
Student Name: <value>  
Mentor: <value>  
Issue: <value>  
Brief: <value>  
Tier Classification: <value>  
Sent to: <value>  
Handover Items: <value>

Separate each student's result with a blank line.
Only output values without additional commentary.
        """

        try:
            response = openai.ChatCompletion.create(
                model="gpt-4-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful reviewer."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0
            )
            result = response['choices'][0]['message']['content']

            # Split into individual student blocks
            blocks = [b.strip() for b in result.strip().split("\n\n") if b.strip()]
            student_rows = []

            for block in blocks:
                fields = {}
                lines = block.split("\n")
                for line in lines:
                    if ":" in line:
                        key, value = line.split(":", 1)
                        fields[key.strip()] = value.strip()
                row = [
                    fields.get("Date of Email", ""),
                    fields.get("Date of Close of Email Thread", ""),
                    fields.get("PS Number", ""),
                    fields.get("Student Name", ""),
                    fields.get("Mentor", ""),
                    fields.get("Issue", ""),
                    fields.get("Brief", ""),
                    fields.get("Tier Classification", ""),
                    fields.get("Sent to", ""),
                    fields.get("Handover Items", "")
                ]
                student_rows.append(row)

            return student_rows

        except Exception as e:
            st.error(f"API call failed: {e}")
            return [["Error"] * 10]

    if st.button("🔍 Perform Review"):
        with st.spinner("Reviewing... Please wait."):

            all_rows = []

            for _, row in df.iterrows():
                email_text = row["Email"]
                result_rows = review_email(email_text)
                for r in result_rows:
                    all_rows.append({
                        "Original Email": email_text,
                        "Date of Email": r[0],
                        "Date of Close of Email Thread": r[1],
                        "PS Number": r[2],
                        "Student Name": r[3],
                        "Mentor": r[4],
                        "Issue": r[5],
                        "Brief": r[6],
                        "Tier Classification": r[7],
                        "Sent to": r[8],
                        "Handover Items": r[9]
                    })

            final_df = pd.DataFrame(all_rows)

            # Export to Excel
            output = BytesIO()
            final_df.to_excel(output, index=False, engine='openpyxl')
            st.success("✅ Review Complete!")

            st.download_button(
                label="📥 Download Reviewed File",
                data=output.getvalue(),
                file_name="Email_Reviewed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

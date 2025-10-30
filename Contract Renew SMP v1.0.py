import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.message import EmailMessage

def load_file():
    uploaded_file = st.sidebar.file_uploader("Upload Excel or CSV", type=['xlsx', 'csv'])
    if uploaded_file:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
        else:
            df = pd.read_csv(uploaded_file)
        return df
    return None

def display_contract_status(df):
    st.header("Contract Status by Subscription Status")
    df = df[['Subscription Status', 'SOLDTO_NAME', 'FINANCIAL_MATERIAL_NUMBER', 'Material Name', 'Validfrom', 'Validuntil']]
    df = df.dropna(subset=['Subscription Status'])

    for status in df['Subscription Status'].unique():
        status_lower = status.lower()
        if status_lower == 'expired':
            color = '#f44336'  # Red
        elif status_lower == 'active':
            color = '#4CAF50'  # Green
        else:
            color = '#2196F3'  # Blue for others

        st.markdown(f'''
            <div style="background-color:{color}; padding: 8px; border-radius: 5px; margin-top: 10px; margin-bottom: 5px;">
                <h4 style="color:white; margin:0;">Subscription Status: {status}</h4>
            </div>
        ''', unsafe_allow_html=True)

        filtered_df = df[df['Subscription Status'] == status]
        st.dataframe(filtered_df, use_container_width=True)

def get_upcoming_renewals(df):
    today = datetime.today()
    df = df[['SOLDTO_NAME', 'FINANCIAL_MATERIAL_NUMBER', 'Material Name', 'Validfrom', 'Validuntil']]
    df['Validuntil'] = pd.to_datetime(df['Validuntil'], errors='coerce')
    df = df.dropna(subset=['Validuntil'])
    three_months_later = today + timedelta(days=90)
    upcoming = df[(df['Validuntil'] >= today) & (df['Validuntil'] <= three_months_later)]
    return upcoming

def display_upcoming_renewals(df):
    st.header("Upcoming Renewals (within 3 months)")
    upcoming = get_upcoming_renewals(df)
    st.dataframe(upcoming, use_container_width=True)
    return upcoming

def send_email(sender_email, sender_password, recipients, subject, upcoming_df, expired_df):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = ', '.join(recipients)

    upcoming_html = upcoming_df.to_html(index=False)
    expired_html = expired_df.to_html(index=False)
    html_content = f"""
    <html>
        <body>
            <h2>Upcoming Renewals</h2>
            {upcoming_html}
            <br><br>
            <h2>Expired Contracts</h2>
            {expired_html}
        </body>
    </html>
    """

    msg.add_alternative(html_content, subtype='html')

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, sender_password)
        smtp.send_message(msg)

def main():
    st.sidebar.title("Upload your data and email credentials")
    df = load_file()

    sender_email = st.sidebar.text_input("Sender Email")
    sender_password = st.sidebar.text_input("Sender Email Password", type="password")

    recipient_emails = ['samhita.bhakta@roche.com']

    if df is not None:
        tabs = st.tabs(["Contract Status", "Upcoming Renewals"])
        with tabs[0]:
            display_contract_status(df)
        with tabs[1]:
            upcoming = display_upcoming_renewals(df)

            expired = df[df['Subscription Status'].str.lower() == 'expired'][['SOLDTO_NAME', 'FINANCIAL_MATERIAL_NUMBER', 'Material Name', 'Validfrom', 'Validuntil']]

            if st.button("Send Email with Upcoming Renewals and Expired"):
                confirm = st.radio("Do you want to send the list to the recipients?", ("Yes", "No"))
                if confirm == "Yes":
                    try:
                        send_email(
                            sender_email,
                            sender_password,
                            recipient_emails,
                            "Upcoming Renewals and Expired Contracts Report",
                            upcoming,
                            expired
                        )
                        st.success("Email sent successfully!")
                    except Exception as e:
                        st.error(f"Failed to send email: {e}")
                else:
                    st.info("Email sending cancelled.")
    else:
        st.info("Please upload an Excel or CSV file to proceed.")

if __name__ == "__main__":
    main()

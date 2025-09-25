import streamlit as st
import pandas as pd
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# ----------------------
# CONFIG
# ----------------------
EXCEL_FILE = "OSID DATA.xlsx"
TARGET_EMAIL = "akhilmp@myg.in"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "jasil@myg.in"
SENDER_PASSWORD = "vurw qnwv ynys xkrf"   # ⚠️ Use app password, not Gmail password

WEB_APP_URL = "https://script.google.com/macros/s/AKfycby48-irQy37Eq_SQKJSpv70xiBFyajtR5ScIBfeRclnvYqAMv4eVCtJLZ87QUJADqXt/exec"

# ----------------------
# LOAD EXCEL DATA
# ----------------------
@st.cache_data
def load_data():
    df = pd.read_excel(EXCEL_FILE)
    df.columns = df.columns.str.strip().str.replace("\u00A0", " ").str.lower()
    return df

df = load_data()

# ----------------------
# STREAMLIT TABS
# ----------------------
tab1, tab2 = st.tabs(["Submit Warranty Claim", "Claim Status"])

with tab1:
    st.title("📌 Submit Warranty Claim")
    mobile_no_input = st.text_input("Enter Customer Mobile No")

    if mobile_no_input:
        customer_data = df[df["mobile no"].astype(str) == mobile_no_input.strip()]

        if not customer_data.empty:
            customer_name = customer_data["customer"].iloc[0]

            st.subheader("Customer Details")
            st.write(f"**Customer Name:** {customer_name}")
            st.write(f"**Mobile:** {mobile_no_input.strip()}")

            customer_address = st.text_area("Enter Customer Address")
            issue_description = st.text_area("Describe the Issue")

            st.subheader("Purchased Products")
            customer_data["product display"] = (
                "Invoice: " + customer_data["invoice no"].astype(str) +
                " | Model: " + customer_data["model"].astype(str) +
                " | Serial No: " + customer_data["serial no"].astype(str) +
                " | OSID: " + customer_data["osid"].astype(str)
            )

            product_choices = st.multiselect(
                "Select Product(s) for Claim",
                options=customer_data["product display"].tolist()
            )

            uploaded_file = st.file_uploader("Upload Invoice / Supporting Document", type=["pdf", "jpg", "png"])

            if st.button("Submit Claim"):
                if not product_choices:
                    st.warning("Please select at least one product.")
                elif not customer_address.strip():
                    st.warning("Please enter customer address.")
                elif not issue_description.strip():
                    st.warning("Please describe the issue.")
                else:
                    selected_products = customer_data[customer_data["product display"].isin(product_choices)]

                    # ----------------------
                    # EMAIL BODY (HTML with bold headings)
                    # ----------------------
                    product_info = "<br><br>".join([
                        f"Invoice  : {row['invoice no']}<br>"
                        f"Model    : {row['model']}<br>"
                        f"Serial No: {row['serial no']}<br>"
                        f"OSID     : {row['osid']}"
                        for _, row in selected_products.iterrows()
                    ])

                    body = f"""
<p>Dear Shyla,</p>

<p>We have received a warranty claim for the products purchased by our customer. Please find the details below:</p>

<hr>
<p><b>Customer Information</b></p>
Name       : {customer_name}<br>
Mobile No  : {mobile_no_input.strip()}<br>
Address    : {customer_address}
<hr>

<p><b>Product(s) Details</b></p>
{product_info}
<hr>

<p><b>Issue Description</b></p>
{issue_description}
<hr>

<p>We request your team to review and process this claim at the earliest convenience. Kindly update the claim status once processed.</p>
"""

                    # ----------------------
                    # SEND EMAIL
                    # ----------------------
                    msg = MIMEMultipart()
                    msg["From"] = SENDER_EMAIL
                    msg["To"] = TARGET_EMAIL
                    msg["Subject"] = f"Warranty Claim Submission – {customer_name}"
                    msg.attach(MIMEText(body, "html"))  # <-- HTML email now

                    if uploaded_file is not None:
                        file_attachment = MIMEApplication(uploaded_file.read(), Name=uploaded_file.name)
                        file_attachment['Content-Disposition'] = f'attachment; filename="{uploaded_file.name}"'
                        msg.attach(file_attachment)

                    try:
                        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                            server.starttls()
                            server.login(SENDER_EMAIL, SENDER_PASSWORD)
                            server.sendmail(SENDER_EMAIL, TARGET_EMAIL, msg.as_string())

                        payload = {
                            "customer_name": customer_name,
                            "mobile_no": mobile_no_input.strip(),
                            "address": customer_address,
                            "products": "; ".join(product_choices),
                            "issue_description": issue_description,
                            "status": "Pending"
                        }
                        response = requests.post(WEB_APP_URL, json=payload)
                        if response.status_code == 200:
                            st.success("✅ Claim submitted successfully, email sent, and saved to Google Sheets.")
                        else:
                            st.error(f"❌ Failed to submit to Google Sheets: {response.text}")

                    except Exception as e:
                        st.error(f"❌ Error sending email: {e}")

        else:
            st.warning("No products found for this mobile number.")

with tab2:
    st.title("📌 Warranty Claim Status")
    try:
        response = requests.get(WEB_APP_URL)
        all_claims = pd.DataFrame(response.json())
        if all_claims.empty:
            st.info("No claims submitted yet.")
        else:
            st.dataframe(all_claims)
            st.info("Update the 'Status' column directly in Google Sheets; it will reflect here automatically.")
    except Exception as e:
        st.error(f"❌ Failed to fetch claims: {e}")

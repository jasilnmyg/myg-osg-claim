
import streamlit as st
import pandas as pd
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
import pytz
import time

# ----------------------
# PAGE CONFIGURATION
# ----------------------
st.set_page_config(
    page_title="Warranty Claim Management",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------
# HELPER FUNCTION FOR IST
# ----------------------
def get_ist_datetime():
    """Get current datetime in Indian Standard Time"""
    ist = pytz.timezone('Asia/Kolkata')
    return datetime.now(ist)

def format_ist_datetime(dt_str):
    """Format datetime string to IST display format"""
    try:
        # Parse the datetime string (assuming it's in ISO format)
        dt = pd.to_datetime(dt_str)
        # If timezone-naive, assume it's already IST
        if dt.tz is None:
            ist = pytz.timezone('Asia/Kolkata')
            dt = ist.localize(dt)
        else:
            # Convert to IST if it has timezone info
            ist = pytz.timezone('Asia/Kolkata')
            dt = dt.astimezone(ist)
        return dt.strftime('%Y-%m-%d %H:%M:%S IST')
    except:
        return str(dt_str)

# ----------------------
# CUSTOM CSS STYLING
# ----------------------
st.markdown("""
<style>
    /* Main theme colors */
    :root {
        --primary-color: #2E86C1;
        --secondary-color: #F8F9FA;
        --accent-color: #28A745;
        --warning-color: #FFC107;
        --danger-color: #DC3545;
        --text-color: #2C3E50;
        --border-color: #E9ECEF;
    }

    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Main container styling */
    .main-header {
        background: linear-gradient(135deg, #2E86C1 0%, #5DADE2 100%);
        padding: 1.6rem;
        border-radius: 10px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.06);
    }

    .main-header h1 {
        color: white;
        font-size: 1.8rem;
        font-weight: 700;
        margin: 0;
        text-align: center;
    }

    .main-header p {
        color: rgba(255, 255, 255, 0.95);
        font-size: 0.95rem;
        text-align: center;
        margin: 0.3rem 0 0 0;
    }

    /* Form containers */
    .form-container {
        background: white;
        padding: 1.25rem;
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.06);
        border: 1px solid #E9ECEF;
        margin-bottom: 1.5rem;
    }

    .section-header {
        color: var(--primary-color);
        font-size: 1.05rem;
        font-weight: 600;
        margin-bottom: 0.75rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .info-box, .warning-box, .success-box {
        padding: 0.85rem;
        border-radius: 8px;
        margin: 0.85rem 0;
    }

    .info-box { background: #F0F8FF; border: 1px solid #2E86C1; }
    .warning-box { background: #FFFBF0; border: 1px solid #FFC107; }
    .success-box { background: #F0FFF4; border: 1px solid #28A745; }

    .customer-details { padding: 0.9rem; border-radius: 8px; border: 1px solid #DEE2E6; margin: 0.75rem 0; }

    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #2E86C1 0%, #5DADE2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.25rem;
        font-weight: 600;
    }

    .stFileUploader > div { border: 2px dashed #2E86C1; padding: 1rem; border-radius: 8px; background: #F0F8FF; }

    .metric-container { padding: 0.9rem; border-radius: 8px; border: 1px solid #E9ECEF; background: white; }

    /* Dataframe container fix and larger text */
    .dataframe { 
        border-radius: 8px; 
        overflow: hidden; 
    }
    .dataframe table {
        font-size: 1.2rem !important; /* Increase table text size */
        line-height: 1.5 !important;
    }
    .dataframe th, .dataframe td {
        padding: 0.75rem !important; /* Add padding for better readability */
    }
</style>
""", unsafe_allow_html=True)

# ----------------------
# CONFIG
# ----------------------
EXCEL_FILE = "OSID DATA.xlsx"
TARGET_EMAIL = "shyla.mariadhasan@onsite.co.in"
CC_EMAILS = ["shine.at@onsite.co.in", "akhilmp@myg.in","sachin.kadam@onsite.co.in","shanmugaraja.a@onsite.co.in"]
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "jasil@myg.in"
SENDER_PASSWORD = "vurw qnwv ynys xkrf"
WEB_APP_URL = "https://script.google.com/macros/s/AKfycby48-irQy37Eq_SQKJSpv70xiBFyajtR5ScIBfeRclnvYqAMv4eVCtJLZ87QUJADqXt/exec"

# ----------------------
# SIDEBAR
# ----------------------
with st.sidebar:
    st.markdown("### ℹ️ Help & Support")
    st.markdown("**Contact Information:**")
    st.markdown("📧 Email: jasil@myg.in")
    st.markdown("📱 Phone: +91 8589852747")
    st.markdown("🕒 Business Hours: 9 AM - 6 PM")

# ----------------------
# MAIN HEADER
# ----------------------
st.markdown("""
<div class="main-header">
    <h1>🛡️ Warranty Claim Management System</h1>
    <p>Professional warranty claim submission and tracking platform</p>
</div>
""", unsafe_allow_html=True)

# ----------------------
# LOAD EXCEL DATA
# ----------------------
@st.cache_data(ttl=300)
def load_data():
    """Loads the Excel data and normalizes column names. Returns empty DataFrame on failure."""
    try:
        df = pd.read_excel(EXCEL_FILE)
        # Normalize headers to simple keys to avoid KeyError
        df.columns = (
            df.columns.astype(str)
            .str.strip()
            .str.replace("\u00A0", " ")
            .str.lower()
            .str.replace(r"\s+", " ", regex=True)
        )
        return df
    except Exception as e:
        st.warning(f"Could not load Excel file: {e}")
        return pd.DataFrame()

# load once
df = load_data()

# helper: safe column lookup
def col(df, name_variants):
    """Return the first matching column name from the dataframe for a list of variants."""
    for n in name_variants:
        if n in df.columns:
            return n
    return None

# Common column name variants
mobile_col = col(df, ["mobile no", "mobile", "mobile_no", "mobile no rf"]) or "mobile no"
invoice_col = col(df, ["invoice no", "invoice", "invoice_no"]) or "invoice no"
model_col = col(df, ["model"]) or "model"
serial_col = col(df, ["serial no", "serialno", "serial_no"]) or "serial no"
osid_col = col(df, ["osid"]) or "osid"
customer_col = col(df, ["customer", "customer name"]) or "customer"

# ----------------------
# MAIN TABS
# ----------------------
tab1, tab2 = st.tabs(["🔧 Submit New Claim", "📋 Track Claims"]) 

# ----------------------
# TAB 1: Submit Warranty Claim
# ----------------------
with tab1:
    st.markdown('<div class="form-container">', unsafe_allow_html=True)

    st.markdown('<div class="section-header">👤 Customer Information</div>', unsafe_allow_html=True)

    mobile_no_input = st.text_input(
        "📱 Customer Mobile Number *",
        placeholder="Enter 10-digit mobile number",
        help="Enter the mobile number registered with the purchase"
    ).strip()

    # Only proceed to lookup when valid 10-digit number entered
    customer_data = pd.DataFrame()
    if mobile_no_input and len(mobile_no_input) == 10 and mobile_no_input.isdigit():
        if not df.empty:
            # Ensure comparison on string values
            try:
                customer_data = df[df[mobile_col].astype(str).str.strip() == mobile_no_input]
            except Exception:
                # fallback: try any numeric match
                customer_data = df[df[mobile_col].astype(str).str.contains(mobile_no_input, na=False)]

        if customer_data.empty:
            st.markdown("""
            <div class="warning-box">
                ⚠️ <strong>No products found</strong><br>
                No products are registered under this mobile number. Please verify the number or contact support.
            </div>
            """, unsafe_allow_html=True)
        else:
            # pick first readable fields safely
            customer_name = str(customer_data.get(customer_col, pd.Series(["Unknown"])).iloc[0])

            st.markdown(f"""
            <div class="customer-details">
                <h4 style="color: var(--primary-color);">✅ Customer Found</h4>
                <div><strong>Customer Name:</strong> {customer_name}</div>
                <div><strong>Mobile Number:</strong> {mobile_no_input}</div>
                <div><strong>Total Products:</strong> {len(customer_data)} item(s)</div>
            </div>
            """, unsafe_allow_html=True)

            # Address & issue
            customer_address = st.text_area(
                "📍 Complete Service Address *",
                placeholder="Enter complete address including pincode",
                help="This address will be used for service appointment",
                height=100
            ).strip()

            issue_description = st.text_area(
                "📝 Describe the Issue *",
                placeholder="Please describe the problem you're experiencing with the product(s)",
                help="Provide detailed information about the issue for faster resolution",
                height=120
            ).strip()

            # Build display string for products
            def make_display(r):
                inv = r.get(invoice_col, "")
                mod = r.get(model_col, "")
                ser = r.get(serial_col, "")
                osid = r.get(osid_col, "")
                return f"Invoice: {inv} | Model: {mod} | Serial: {ser} | OSID: {osid}"

            customer_data = customer_data.copy()
            customer_data["product_display"] = customer_data.apply(make_display, axis=1)

            product_choices = st.multiselect(
                "Select Product(s) for Warranty Claim *",
                options=customer_data["product_display"].tolist(),
                help="You can select multiple products if the issue affects more than one item"
            )

            uploaded_file = st.file_uploader(
                "📄 Upload Invoice / Supporting Document *",
                type=["pdf", "jpg", "jpeg", "png"],
                help="Upload invoice, receipt, or any other supporting document"
            )

            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                submit_button = st.button("🚀 Submit Warranty Claim", use_container_width=True)

            if submit_button:
                # server-side validation
                errors = []
                if not mobile_no_input:
                    errors.append("Mobile number is required")
                elif len(mobile_no_input) != 10 or not mobile_no_input.isdigit():
                    errors.append("Please enter a valid 10-digit mobile number")

                if not customer_address:
                    errors.append("Customer address is required")

                if not issue_description:
                    errors.append("Issue description is required")

                if not product_choices:
                    errors.append("Please select at least one product")

                if uploaded_file is None:
                    errors.append("Please upload a supporting document")

                if errors:
                    for error in errors:
                        st.error(f"❌ {error}")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    try:
                        status_text.text("📧 Preparing email...")
                        progress_bar.progress(10)

                        selected_products = customer_data[customer_data["product_display"].isin(product_choices)]

                        product_info = "<br><br>".join([
                            f"Invoice  : {row.get(invoice_col, '')}<br>"
                            f"Model    : {row.get(model_col, '')}<br>"
                            f"Serial No: {row.get(serial_col, '')}<br>"
                            f"OSID     : {row.get(osid_col, '')}"
                            for _, row in selected_products.iterrows()
                        ])

                        # Get current IST time
                        ist_time = get_ist_datetime()
                        ist_formatted = ist_time.strftime('%Y-%m-%d %H:%M:%S IST')

                        body = f"""
<div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto;">
    <div style="background: linear-gradient(135deg, #2E86C1 0%, #5DADE2 100%); color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
        <h2 style="margin: 0;">🛡️ Warranty Claim Submission</h2>
        <p style="margin: 5px 0 0 0;">New claim received from customer</p>
    </div>
    <div style="background: #f8f9fa; padding: 20px; border-radius: 0 0 10px 10px;">
        <p>Dear Shyla,</p>
        <p>We have received a warranty claim for the products purchased by our customer. Please find the details below:</p>
        <div style="background: white; padding: 15px; border-radius: 8px; margin: 12px 0; border-left: 4px solid #2E86C1;">
            <h3 style="color: #2E86C1; margin-top: 0;">👤 Customer Information</h3>
            <p><strong>Name:</strong> {customer_name}<br>
            <strong>Mobile No:</strong> {mobile_no_input}<br>
            <strong>Address:</strong> {customer_address}</p>
        </div>
        <div style="background: white; padding: 15px; border-radius: 8px; margin: 12px 0; border-left: 4px solid #28A745;">
            <h3 style="color: #28A745; margin-top: 0;">📦 Product(s) Details</h3>
            <div style="font-family: monospace; font-size: 14px;">{product_info}</div>
        </div>
        <div style="background: white; padding: 15px; border-radius: 8px; margin: 12px 0; border-left: 4px solid #FFC107;">
            <h3 style="color: #FFC107; margin-top: 0;">🔍 Issue Description</h3>
            <p style="background: #f8f9fa; padding: 10px; border-radius: 5px; font-style: italic;">"{issue_description}"</p>
        </div>
        <div style="background: #e7f3ff; padding: 12px; border-radius: 8px; margin: 12px 0;">
            <p><strong>📅 Submitted:</strong> {ist_formatted}</p>
            <p style="margin-bottom: 0;">We request your team to review and process this claim at the earliest convenience.</p>
        </div>
        <div style="text-align: center; margin-top: 14px; padding-top: 10px; border-top: 1px solid #e9ecef;">
            <p style="margin: 0;"><strong>Best Regards,</strong><br>
            <strong>JASIL N</strong><br>
            📞 +918589852747</p>
        </div>
    </div>
</div>
"""

                        status_text.text("📨 Sending email...")
                        progress_bar.progress(40)

                        # Email assembly
                        osid_list = selected_products[osid_col].astype(str).unique().tolist()
                        osid_str = ", ".join(osid_list)

                        msg = MIMEMultipart()
                        msg["From"] = SENDER_EMAIL
                        msg["To"] = TARGET_EMAIL
                        msg["Cc"] = ", ".join(CC_EMAILS)
                        msg["Subject"] = f"🛡️ Warranty Claim Submission – OSID: {osid_str} – {customer_name}"
                        msg.attach(MIMEText(body, "html"))

                        # Attach uploaded file (seek to start in case stream consumed)
                        uploaded_file.seek(0)
                        file_attachment = MIMEApplication(uploaded_file.read(), Name=uploaded_file.name)
                        file_attachment['Content-Disposition'] = f'attachment; filename="{uploaded_file.name}"'
                        msg.attach(file_attachment)

                        recipients = [TARGET_EMAIL] + CC_EMAILS

                        # Send email
                        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
                            server.starttls()
                            server.login(SENDER_EMAIL, SENDER_PASSWORD)
                            server.sendmail(SENDER_EMAIL, recipients, msg.as_string())

                        status_text.text("💾 Saving to database...")
                        progress_bar.progress(70)

                        # Submit to Google Sheets (or other endpoint)
                        payload = {
                            "customer_name": customer_name,
                            "mobile_no": mobile_no_input,
                            "address": customer_address,
                            "products": "; ".join(product_choices),
                            "issue_description": issue_description,
                            "status": "Pending",
                            "submitted_date": ist_time.isoformat()
                        }

                        try:
                            response = requests.post(WEB_APP_URL, json=payload, timeout=8)
                            post_ok = (response.status_code == 200)
                        except Exception:
                            post_ok = False

                        progress_bar.progress(100)
                        status_text.text("✅ Completed!")
                        time.sleep(0.8)
                        progress_bar.empty()
                        status_text.empty()

                        if post_ok:
                            st.markdown(f"""
                            <div class="success-box">
                                <h3 style="margin-top: 0;">🎉 Claim Submitted Successfully!</h3>
                                <p><strong>✅ Submission Date & Time:</strong> {ist_formatted}</p>
                                <p><strong>✅ Email sent to warranty team with CC to relevant stakeholders</strong></p>
                                <p><strong>✅ Claim saved to tracking system</strong></p>
                                <p><strong>✅ Supporting documents attached</strong></p>
                                <hr style="border-color: #28A745;">
                                <p style="margin-bottom: 0;"><strong>Next Steps:</strong> Our warranty team will review your claim and contact you within 24-48 hours. You can track the status in the "Track Claims" tab.</p>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.error("❌ Failed to submit to tracking system. Email was sent but saving failed.")

                    except Exception as e:
                        progress_bar.empty()
                        status_text.empty()
                        st.error(f"❌ Error processing claim: {e}")

    elif mobile_no_input and (len(mobile_no_input) != 10 or not mobile_no_input.isdigit()):
        st.markdown("""
        <div class="warning-box">
            ⚠️ <strong>Invalid mobile number</strong><br>
            Please enter a valid 10-digit mobile number.
        </div>
        """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ----------------------
# TAB 2: Claim Status
# ----------------------
with tab2:
    st.markdown('<div class="form-container">', unsafe_allow_html=True)

    st.markdown('<div class="section-header">🔍 Search & Filter Claims</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([3, 1])
    with col1:
        search_mobile = st.text_input(
            "📱 Filter by Mobile Number",
            placeholder="Enter mobile number to filter claims (leave blank for all)",
            help="Enter mobile number to view specific customer claims"
        ).strip()

    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        refresh_button = st.button("🔄 Refresh Data", use_container_width=True)

    # Fetch claims each time so refresh works
    try:
        response = requests.get(WEB_APP_URL, timeout=8)
        if response.status_code == 200:
            all_claims = pd.DataFrame(response.json())

            if not all_claims.empty:
                # Normalize columns
                all_claims.columns = all_claims.columns.str.strip().str.lower().str.replace(" ", "_")

                if search_mobile:
                    filtered_claims = all_claims[all_claims.get('mobile_no', all_claims.columns[0]).astype(str).str.strip() == search_mobile]
                else:
                    filtered_claims = all_claims

                if filtered_claims.empty:
                    if search_mobile:
                        st.markdown(f"""
                        <div class="info-box">
                            ℹ️ <strong>No claims found</strong><br>
                            No claims found for mobile number: <strong>{search_mobile}</strong>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown("""
                        <div class="info-box">
                            ℹ️ <strong>No claims submitted yet</strong><br>
                            No warranty claims have been submitted yet.
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="section-header">📊 Claims Overview ({len(filtered_claims)} records)</div>
                    """, unsafe_allow_html=True)

                    # Pretty card view
                    for idx, claim in filtered_claims.iterrows():
                        status = str(claim.get('status', 'Unknown')).lower()
                        status_label = claim.get('status', 'Unknown')

                        st.markdown(f"""
                        <div style="background: white; border: 1px solid #e9ecef; border-radius: 10px; padding: 1rem; margin: 0.6rem 0;">
                            <div style="display:flex; justify-content:space-between; align-items:center;">
                                <h4 style="color: var(--primary-color); margin: 0;">{claim.get('customer_name', 'N/A')}</h4>
                                <div style="font-weight:700;">{status_label}</div>
                            </div>
                            <div style="display:grid; grid-template-columns: 1fr 1fr; gap: 0.5rem; margin-top: 0.6rem;">
                                <div>
                                    <strong>📱 Mobile:</strong> {claim.get('mobile_no', 'N/A')}<br>
                                    <strong>📍 Address:</strong> {str(claim.get('address', 'N/A'))[:60]}{'...' if len(str(claim.get('address', 'N/A'))) > 60 else ''}
                                </div>
                                <div>
                                    <strong>📦 Products:</strong> {str(claim.get('products', 'N/A'))[:60]}{'...' if len(str(claim.get('products', 'N/A'))) > 60 else ''}<br>
                                    <strong>🔍 Issue:</strong> {str(claim.get('issue_description', 'N/A'))[:60]}{'...' if len(str(claim.get('issue_description', 'N/A'))) > 60 else ''}
                                </div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                    # Tabular view
                    st.markdown('<div class="section-header">📋 Detailed Claims Table</div>', unsafe_allow_html=True)

                    display_df = filtered_claims.copy()
                    # format dates with IST
                    if 'submitted_date' in display_df.columns:
                        display_df['submitted_date'] = display_df['submitted_date'].apply(
                            lambda x: format_ist_datetime(x) if pd.notna(x) else 'N/A'
                        )
                    elif 'timestamp' in display_df.columns:
                        display_df['timestamp'] = display_df['timestamp'].apply(
                            lambda x: format_ist_datetime(x) if pd.notna(x) else 'N/A'
                        )

                    # Rename for nicer column names
                    rename_map = {
                        'customer_name': 'Customer Name',
                        'mobile_no': 'Mobile No',
                        'address': 'Address',
                        'products': 'Products',
                        'issue_description': 'Issue Description',
                        'status': 'Status',
                        'submitted_date': 'Submitted Date (IST)',
                        'timestamp': 'Timestamp (IST)'
                    }
                    display_df = display_df.rename(columns={k: v for k, v in rename_map.items() if k in display_df.columns})

                    st.dataframe(display_df, use_container_width=True, hide_index=True)

            else:
                st.markdown("""
                <div class="info-box">
                    ℹ️ <strong>No claims found</strong><br>
                    No warranty claims have been submitted yet. Start by submitting a new claim in the "Submit New Claim" tab.
                </div>
                """, unsafe_allow_html=True)
        else:
            st.error(f"❌ Failed to fetch claims data. Status code: {response.status_code}")

    except requests.exceptions.RequestException as e:
        st.error(f"❌ Network error while fetching claims: {e}")
    except Exception as e:
        st.error(f"❌ Unexpected error while fetching claims: {e}")

    st.markdown("""
    <div class="info-box">
        <h4 style="color: var(--primary-color); margin-top: 0;">📝 How Status Updates Work</h4>
        <ul style="margin-bottom: 0;">
            <li><strong>Pending:</strong> Claim submitted and under review</li>
            <li><strong>Approved:</strong> Claim approved, service will be scheduled</li>
            <li><strong>In Progress:</strong> Service technician assigned</li>
            <li><strong>Completed:</strong> Service completed successfully</li>
            <li><strong>Rejected:</strong> Claim not covered under warranty</li>
        </ul>
        <hr>
        <p style="margin-bottom: 0;"><strong>Note:</strong> Status updates are managed by the warranty team through Google Sheets and will reflect here automatically when you refresh the data. All timestamps are displayed in Indian Standard Time (IST).</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ----------------------
# FOOTER
# ----------------------
st.markdown("""
---
<div style="text-align: center; padding: 1rem; background: #f8f9fa; border-radius: 10px; margin-top: 1rem;">
    <h4 style="color: var(--primary-color); margin-bottom: 0.3rem;">🛡️ Warranty Claim Management System</h4>
    <p style="margin: 0; color: #6c757d;">Powered by Loyalty Operation | 📧 <a href="mailto:jasil@myg.in" style="color: var(--primary-color);">jasil@myg.in</a> | 📱 <a href="tel:+918589852747" style="color: var(--primary-color);">+91 8589852747</a></p>
</div>
""", unsafe_allow_html=True)

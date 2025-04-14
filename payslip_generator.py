import os
import pandas as pd
from fpdf import FPDF
import yagmail
from dotenv import load_dotenv

# Load email credentials from .env
load_dotenv()
EMAIL_USER = os.getenv("brandontembo78gmail.com")
EMAIL_PASS = os.getenv("uitg jpqa qsvl eiqe")

# Create output folder
os.makedirs("payslips", exist_ok=True)

# Load employee data
try:
    df = pd.read_excel("employees.xlsx")
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

# Validate required columns
required_columns = ['Employee ID', 'Name', 'Email', 'Basic Salary', 'Allowances', 'Deductions']
if not all(col in df.columns for col in required_columns):
    print("Excel file is missing one or more required columns.")
    exit()

# Calculate Net Salary
df["Net Salary"] = df["Basic Salary"] + df["Allowances"] - df["Deductions"]

# Generate and send payslips
for _, row in df.iterrows():
    try:
        # Extract employee info
        emp_id = row['Employee ID']
        name = row['Name']
        email = row['Email']
        basic = row['Basic Salary']
        allow = row['Allowances']
        deduct = row['Deductions']
        net = row['Net Salary']

        # Create PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        pdf.cell(200, 10, txt=f"Payslip for {name} (ID: {emp_id})", ln=True, align='C')
        pdf.ln(10)
        pdf.cell(200, 10, txt=f"Basic Salary:     ${basic:.2f}", ln=True)
        pdf.cell(200, 10, txt=f"Allowances:       ${allow:.2f}", ln=True)
        pdf.cell(200, 10, txt=f"Deductions:       ${deduct:.2f}", ln=True)
        pdf.cell(200, 10, txt=f"Net Salary:       ${net:.2f}", ln=True)
        pdf.ln(10)
        pdf.cell(200, 10, txt="Thank you for your service.", ln=True)

        # Save PDF
        pdf_path = f"payslips/{emp_id}.pdf"
        pdf.output(pdf_path)

        # Send email with payslip attached
        yag = yagmail.SMTP("brandontembo78@gmail.com", "uitg jpqa qsvl eiqe")
        subject = "Your Payslip for This Month"
        body = f"Hi {name},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nHR Team"
        yag.send(to=email, subject=subject, contents=body, attachments=pdf_path)

        print(f"Payslip sent to {name} ({email})")

    except Exception as e:
        print(f"Error processing employee ID {row['Employee ID']}: {e}")

import os
from datetime import datetime
from flask import Flask, request, send_file, render_template_string
import pandas as pd
from fpdf import FPDF
import zipfile

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'receipts'
FONT_PATH_REG = 'DejaVuSans.ttf'
FONT_PATH_BOLD = 'DejaVuSans-Bold.ttf'
LOGO_PATH = 'WhatsApp Image 2025-05-20 at 16.18.57.jpeg'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

class PDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font("DejaVu", "", FONT_PATH_REG, uni=True)
        self.add_font("DejaVu", "B", FONT_PATH_BOLD, uni=True)
        self.set_font("DejaVu", size=12)

    def header(self):
        self.image(LOGO_PATH, x=10, y=8, w=30)
        self.set_y(20)
        self.set_font("DejaVu", size=15)
        self.cell(0, 10, "INTEGRITY COURT RESIDENTS ASSOCIATION (LANDLORD)", ln=True, align="R")
        self.ln(10)

@app.route('/', methods=['GET', 'POST'])
def index():
    html = '''
    <!doctype html>
    <title>Upload Excel File</title>
    <h1>Upload Excel to Generate Receipts</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file><br><br>
      <input type=submit value=Upload>
    </form>
    {% if download_link %}
      <h2>Receipts generated successfully.</h2>
      <a href="{{ download_link }}">Download ZIP</a>
    {% endif %}
    '''

    if request.method == 'POST':
        file = request.files['file']
        if file:
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            df = pd.read_excel(filepath, sheet_name='FULL VIEW')
            df = df.iloc[2:].copy()
            df.columns = [
                "House", "Landlord", "Total Bill Due", "Fence Dues", "Painting", "Generator Due",
                "Prev Ground Rents", "2022 Ground Rent", "2023 Ground Rent", "2024 Ground Rent",
                "2025 Ground Rent", "CofO Payment", "Total Paid", "Total Outstanding"
            ]
            df = df[df["House"].notna()]
            df.fillna(0, inplace=True)

            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            receipt_paths = []

            for _, row in df.iterrows():
                house = str(row["House"]).strip()
                receipt_text = f"""
{timestamp}
Landlord: {row['Landlord']}
House {house}
{'='*40}
Total Bill Due: ₦{int(row['Total Bill Due']):,}
Breakdown:
  - Fence Dues: ₦{int(row['Fence Dues']):,}
  - Painting: ₦{int(row['Painting']):,}
  - Generator Due: ₦{int(row['Generator Due']):,}
  - Previous Ground Rents till 2021: ₦{int(row['Prev Ground Rents']):,}
  - 2022 Ground Rent: ₦{int(row['2022 Ground Rent']):,}
  - 2023 Ground Rent: ₦{int(row['2023 Ground Rent']):,}
  - 2024 Ground Rent: ₦{int(row['2024 Ground Rent']):,}
  - 2025 Ground Rent: ₦{int(row['2025 Ground Rent']):,}
  - CofO Payment: ₦{int(row['CofO Payment']):,}

Payments:
  - Total Paid: ₦{int(row['Total Paid']):,}
  - Total Outstanding: ₦{int(row['Total Outstanding']):,}

{'-'*30}
Authorized Signatory
"""
                pdf = PDF()
                pdf.add_page()
                page_width = pdf.w - 2 * pdf.l_margin
                naira_x = 130
                lines = receipt_text.split("\n")

                for idx, line in enumerate(lines):
                    if line.strip() == "Breakdown:":
                        pdf.set_font("DejaVu", "B", size=12)
                        pdf.set_x(pdf.l_margin)
                        pdf.cell(0, 10, text=line, ln=True)
                        pdf.set_font("DejaVu", style="", size=12)
                        continue
                    if idx == 4:
                        line_width = pdf.get_string_width(line)
                        x_position = (pdf.w - line_width) / 2
                        pdf.set_x(x_position)
                        pdf.cell(line_width, 10, text=line, ln=True)
                    elif "₦" in line and ": ₦" in line:
                        label, amount_str = line.split(": ₦")
                        amount = amount_str.strip()
                        pdf.set_x(pdf.l_margin)
                        pdf.cell(naira_x - pdf.l_margin - 2, 10, text=label + ":", ln=False)
                        pdf.set_x(naira_x)
                        pdf.cell(5, 10, text="₦", ln=False)
                        amount_width = pdf.get_string_width(amount)
                        pdf.set_x(naira_x + 5 + (40 - amount_width))
                        pdf.cell(amount_width, 10, text=amount, ln=True)
                    elif idx == 1:
                        line_width = pdf.get_string_width(line)
                        x_position = pdf.w - pdf.r_margin - line_width
                        pdf.set_x(x_position)
                        pdf.cell(line_width, 10, text=line, ln=True)
                    else:
                        pdf.set_x(pdf.l_margin)
                        pdf.cell(0, 10, text=line, ln=True)

                receipt_filename = f"Receipt_{house}.pdf"
                receipt_path = os.path.join(OUTPUT_FOLDER, receipt_filename)
                pdf.output(receipt_path)
                receipt_paths.append(receipt_path)

            zip_path = os.path.join(OUTPUT_FOLDER, "all_receipts.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for file in receipt_paths:
                    zipf.write(file, os.path.basename(file))

            return render_template_string(html, download_link=f"/download")

    return render_template_string(html, download_link=None)

@app.route('/download')
def download_zip():
    return send_file(os.path.join(OUTPUT_FOLDER, "all_receipts.zip"), as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

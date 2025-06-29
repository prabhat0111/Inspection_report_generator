from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os
import csv
from generate_report import generate_report_and_pdf
from config import EMAIL_USER, EMAIL_PASS
import smtplib
from email.message import EmailMessage

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['REPORT_FOLDER'] = 'reports'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORT_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        data = {
            "insured_name": request.form['insured_name'],
            "address": request.form['address'],
            "insurer": request.form['insurer'],
            "claim_number": request.form['claim_number'],
            "date_of_inspection": request.form['date_of_inspection'],
            "date_of_loss": request.form['date_of_loss'],
            "date_of_report": request.form['date_of_report'],
            "type_of_loss": request.form['type_of_loss'],
            "cause_of_loss": request.form['cause_of_loss'],
            "indemnity_work": request.form['indemnity_work'],
            "listing_pricing_reserve": request.form['listing_pricing_reserve'],
            "contents_loss_reserve": request.form['contents_loss_reserve'],
            "email": request.form['email']
        }

        # Save photos
        photo_folder = os.path.join(app.config['UPLOAD_FOLDER'], data["claim_number"])
        os.makedirs(photo_folder, exist_ok=True)
        for file in request.files.getlist('photos'):
            if file:
                file.save(os.path.join(photo_folder, secure_filename(file.filename)))

        word_path, pdf_path = generate_report_and_pdf(data, photo_folder, app.config['REPORT_FOLDER'])

        # Log to CSV
        with open('report_log.csv', 'a', newline='') as f:
            writer = csv.writer(f)
            if f.tell() == 0:
                writer.writerow(data.keys())
            writer.writerow(data.values())

        # Email report
        msg = EmailMessage()
        msg['Subject'] = f"Inspection Report - {data['claim_number']}"
        msg['From'] = EMAIL_USER
        msg['To'] = data['email']
        msg.set_content(f"Dear {data['insured_name']},\n\nPlease find your fire inspection report attached.")

        with open(pdf_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(pdf_path))

        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_USER, EMAIL_PASS)
                smtp.send_message(msg)
        except Exception as e:
            print("Email sending failed:", e)

        return send_file(pdf_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

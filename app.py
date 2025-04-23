from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Inches
import openai
import os
from dotenv import load_dotenv

# Load environment variables (including OPENAI_API_KEY)
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Get form data
        GP_practice = request.form.get("GP_practice")
        patient_name = request.form.get("patient_name")
        nhs_number = request.form.get("nhs_number")
        date_of_clinic = request.form.get("date_of_clinic")
        store_name = request.form.get("store_name")
        store_phone_number = request.form.get("store_phone_number")
        audiologist = request.form.get("audiologist")
        clinical_notes = request.form.get("clinical_notes")

        # Craft the prompt
        prompt = (
            "You are an audiologist generating clinical letter content for a referring GP. "
            "Use UK English. Do NOT include headers, addresses, greetings, or closings. "
            "Start with: 'The patient was seen at [store_name] today for an audiology assessment.' "
            "Do not say 'Thank you for referring...' or add extra commentary. "
            "Do not suggest referrals or actions unless stated in the original text. "
            "Only proofread, correct grammar, and condense the clinical notes below into a clear, concise summary:\n\n"
            f"{clinical_notes}"
        )

        # Replace placeholder with actual store name
        processed_prompt = prompt.replace("[store_name]", store_name)

        # Call OpenAI to get the proofread and refined text
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {
                    "role": "system",
                    "content": "You are a UK audiologist writing professional, concise clinical summaries."
                },
                {
                    "role": "user",
                    "content": processed_prompt
                }
            ],
            temperature=0.5
        )

        generated_letter = response['choices'][0]['message']['content']

        # Create Word document
        doc = Document()

        # Add logo
        logo_paragraph = doc.add_paragraph()
        logo_paragraph.alignment = 1  # Center
        logo_run = logo_paragraph.add_run()
        logo_run.add_picture('images/Specsaversaudio-removebg-preview.png', width=Inches(2.5))

        # Store info
        store_info = doc.add_paragraph()
        store_info.alignment = 1
        store_info.add_run(f"{store_name}\n{store_phone_number}").bold = True

        doc.add_paragraph("\n")

        # Patient details
        doc.add_heading('Patient Details', level=2)
        doc.add_paragraph(f"Patient Name: {patient_name}")
        doc.add_paragraph(f"NHS Number: {nhs_number}")
        doc.add_paragraph(f"GP Practice: {GP_practice}")
        doc.add_paragraph(f"Clinic Date: {date_of_clinic}")

        doc.add_paragraph("\n")

        # Insert refined clinic letter
        doc.add_heading('Clinic Letter', level=2)
        doc.add_paragraph(generated_letter)

        doc.add_paragraph("\n")
        doc.add_paragraph(f"Please contact the service on {store_phone_number} if you need any further information.")
        doc.add_paragraph("\n")

        # Signature
        doc.add_paragraph(f"Yours sincerely,\n\n{audiologist}\nAudiologist")

        # Save and return the Word document
        output_filename = "clinic_letter.docx"
        doc.save(output_filename)
        return send_file(output_filename, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    if not os.path.exists('images'):
        os.makedirs('images')
    app.run(debug=True)



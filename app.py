from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import FileResponse, HTMLResponse
import shutil
import os
from evaluate_application import main as eval_main, create_excel_report, read_application_text, evaluate_application, EVALUATION_QUESTIONS, EVALUATION_QUESTIONS_OPPSTART_1
import pandas as pd
import re

app = FastAPI()

UPLOAD_DIR = "uploads"
RESULT_DIR = "results"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(RESULT_DIR, exist_ok=True)

@app.get("/", response_class=HTMLResponse)
def index():
    return """
    <html>
    <head>
    <title>Innovasjon Norge Søknadsevaluering</title>
    <meta name='viewport' content='width=device-width, initial-scale=1'>
    <style>
      body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f7fa; margin: 0; padding: 0; }
      .container { max-width: 420px; margin: 40px auto; background: #fff; border-radius: 12px; box-shadow: 0 2px 16px #0001; padding: 32px 28px; }
      h2 { text-align: center; color: #1F4E79; margin-bottom: 24px; }
      label { font-weight: 500; color: #1F4E79; }
      select, input[type='file'] { width: 100%; margin: 8px 0 18px 0; }
      button { width: 100%; background: #1F4E79; color: #fff; border: none; border-radius: 6px; padding: 12px; font-size: 1.1em; font-weight: 600; cursor: pointer; transition: background 0.2s; }
      button:hover { background: #366092; }
      .dropzone { border: 2px dashed #366092; border-radius: 8px; background: #f0f4fa; color: #366092; text-align: center; padding: 32px 10px; margin-bottom: 18px; transition: border 0.2s, background 0.2s; cursor: pointer; }
      .dropzone.dragover { border-color: #1F4E79; background: #e3eaf5; }
      .file-info { margin: 8px 0 18px 0; color: #1F4E79; font-size: 0.98em; }
    </style>
    </head>
    <body>
    <div class='container'>
      <h2>Innovasjon Norge<br>Søknadsevaluering</h2>
      <form id='evalForm' action='/evaluate/' method='post' enctype='multipart/form-data'>
        <div class='dropzone' id='dropzone'>
          Slipp PDF-søknad her eller klikk for å velge fil
          <input type='file' id='fileInput' name='file' accept='.pdf' style='display:none;' required />
        </div>
        <div class='file-info' id='fileInfo'></div>
        <label>Velg oppstartstype:</label>
        <select name='oppstartstype'>
          <option value='Oppstart 1'>Oppstart 1</option>
          <option value='Oppstart 2'>Oppstart 2</option>
          <option value='Oppstart 3'>Oppstart 3</option>
          <option value='NIC'>NIC Klyngeevaluering</option>
        </select>
        <button type='submit'>Evaluer og last ned rapport</button>
      </form>
    </div>
    <script>
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    dropzone.addEventListener('click', () => fileInput.click());
    dropzone.addEventListener('dragover', e => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', e => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
    });
    dropzone.addEventListener('drop', e => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files.length) {
        fileInput.files = e.dataTransfer.files;
        updateFileInfo();
      }
    });
    fileInput.addEventListener('change', updateFileInfo);
    function updateFileInfo() {
      if (fileInput.files.length) {
        fileInfo.textContent = 'Valgt fil: ' + fileInput.files[0].name;
      } else {
        fileInfo.textContent = '';
      }
    }
    // Nedlasting av fil etter submit
    document.getElementById('evalForm').onsubmit = async function(e) {
      e.preventDefault();
      const formData = new FormData(this);
      const btn = this.querySelector('button');
      btn.disabled = true; btn.textContent = 'Vurderer...';
      try {
        const response = await fetch('/evaluate/', { method: 'POST', body: formData });
        if (!response.ok) throw new Error('Noe gikk galt under evalueringen.');
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'evaluering_resultat.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
      } catch (err) {
        alert(err.message);
      } finally {
        btn.disabled = false; btn.textContent = 'Evaluer og last ned rapport';
      }
    };
    </script>
    </body>
    </html>
    """

@app.post("/evaluate/")
def evaluate(file: UploadFile = File(...), oppstartstype: str = Form(...)):
    # Lagre PDF midlertidig
    pdf_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Les søknadstekst
    application_text, selected_pdf = read_application_text(pdf_path)
    pdf_base_name = re.sub(r'[^\w\-_]', '', file.filename.replace('.pdf', '').replace(' ', '_'))
    
    if oppstartstype == "NIC":
        excel_filename = f"nic_evaluering_resultat_{pdf_base_name}.xlsx"
        excel_path = os.path.join(RESULT_DIR, excel_filename)
        from evaluate_nic_application import evaluate_nic_application, create_nic_excel_report
        results_df = evaluate_nic_application(application_text, selected_pdf)
        create_nic_excel_report(results_df, selected_pdf, excel_path)
        return FileResponse(excel_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=excel_filename)
    
    # Velg riktige spørsmål
    if oppstartstype == "Oppstart 1":
        evaluation_questions = EVALUATION_QUESTIONS_OPPSTART_1
    else:
        evaluation_questions = EVALUATION_QUESTIONS

    excel_filename = f"evaluering_resultat_{pdf_base_name}.xlsx"
    excel_path = os.path.join(RESULT_DIR, excel_filename)
    # Evaluer søknad
    results_df = evaluate_application(application_text, selected_pdf, evaluation_questions)
    create_excel_report(results_df, selected_pdf, excel_path, oppstartstype)
    return FileResponse(excel_path, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=excel_filename) 
from flask import Flask, request, render_template, send_file, redirect, url_for, flash, jsonify
from datetime import datetime
from paths import PathManager
from data_processor import DataProcessor
import os
import time
import csv

app = Flask(__name__)
app.secret_key = 'your_secret_key'

# Pfadmanager für dynamische Pfade
path_manager = PathManager()
paths = path_manager.get_paths()

def ensure_directory_exists(directory, retries=5, delay=1):
    """Erstelle das Verzeichnis, wenn es nicht existiert, mit wiederholter Prüfung."""
    for _ in range(retries):
        if not os.path.exists(directory):
            try:
                os.makedirs(directory)
                print(f"[INFO] Das Verzeichnis '{directory}' wurde erstellt.")
            except Exception as e:
                print(f"[ERROR] Fehler beim Erstellen des Verzeichnisses '{directory}': {e}")
        if os.path.exists(directory):
            return True
        time.sleep(delay)
    raise Exception(f"[ERROR] Das Verzeichnis '{directory}' konnte nach mehreren Versuchen nicht erstellt werden.")

# Sicherstellen, dass das Upload-Verzeichnis beim Start der Anwendung existiert
ensure_directory_exists(paths["input_dir"])
ensure_directory_exists(os.path.dirname(paths["output_file"]))

@app.route("/", methods=["GET", "POST"])
def upload_file():
    ensure_directory_exists(paths["input_dir"])  # Sicherstellen, dass das Verzeichnis existiert

    if request.method == "POST":
        if 'file' not in request.files:
            flash('Keine Datei ausgewählt.', 'error')
            return redirect(request.url)

        files = request.files.getlist('file')

        if len(files) == 0 or (len(files) == 1 and files[0].filename == ''):
            flash('Keine Datei ausgewählt.', 'error')
            return redirect(request.url)

        try:
            saved_files = []
            for file in files:
                if file.filename != '':
                    upload_path = os.path.join(paths["input_dir"], file.filename)
                    file.save(upload_path)

                    if os.path.exists(upload_path) and os.path.getsize(upload_path) > 0:
                        saved_files.append(upload_path)
                        flash(f"Datei {file.filename} wurde erfolgreich hochgeladen.", 'success')
                    else:
                        flash(f"Fehler: Datei {file.filename} wurde nicht korrekt hochgeladen.", 'error')

            if not saved_files:
                flash("Keine Dateien konnten erfolgreich gespeichert werden.", 'error')
                return redirect(request.url)

            flash("Die Dateien wurden erfolgreich hochgeladen.", 'success')
            return redirect(url_for('upload_file'))
        except Exception as e:
            flash(f"Fehler beim Hochladen der Dateien: {str(e)}", 'error')
            return redirect(request.url)

    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process_files():
    ensure_directory_exists(paths["input_dir"])  # Sicherstellen, dass das Verzeichnis existiert

    try:
        data_processor = DataProcessor(paths, datetime.now().year)
        input_files = os.listdir(paths["input_dir"])
        input_file_paths = [os.path.join(paths["input_dir"], file) for file in input_files if file.endswith('.xlsx')]

        if not input_file_paths:
            flash("Keine gültigen Dateien zum Verarbeiten gefunden.", 'error')
            return redirect(url_for('upload_file'))

        data_processor.process_files(input_file_paths)

        flash("Bestellliste wurde erfolgreich erstellt.", 'success')
        return redirect(url_for('download_file'))
    except Exception as e:
        flash(f"Fehler bei der Verarbeitung der Bestellliste: {str(e)}", 'error')
        return redirect(url_for('upload_file'))

@app.route("/download")
def download_file():
    output_file_path = paths["output_file"]
    if os.path.exists(output_file_path):
        return send_file(output_file_path, as_attachment=True)
    else:
        flash("Die Ergebnisdatei existiert nicht. Bitte laden Sie eine Datei hoch und führen Sie die Verarbeitung durch.", 'error')
        return redirect(url_for('upload_file'))

@app.route("/get_uploaded_files")
def get_uploaded_files():
    ensure_directory_exists(paths["input_dir"])  # Sicherstellen, dass das Verzeichnis existiert

    try:
        uploaded_files = os.listdir(paths["input_dir"])
        return jsonify(uploaded_files)
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/delete_file/<filename>", methods=["DELETE"])
def delete_file(filename):
    try:
        file_path = os.path.join(paths["input_dir"], filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            return f"Datei {filename} wurde erfolgreich gelöscht.", 200
        else:
            return f"Datei {filename} wurde nicht gefunden.", 404
    except Exception as e:
        return f"Fehler beim Löschen der Datei: {str(e)}", 500

@app.route("/clear_files", methods=["DELETE"])
def clear_files():
    try:
        for filename in os.listdir(paths["input_dir"]):
            file_path = os.path.join(paths["input_dir"], filename)
            if os.path.isfile(file_path):
                os.remove(file_path)

        ensure_directory_exists(paths["input_dir"])

        flash("Alle Dateien wurden erfolgreich gelöscht.", 'success')
        return "Alle Dateien wurden erfolgreich gelöscht.", 200
    except Exception as e:
        return f"Fehler beim Löschen der Dateien: {str(e)}", 500

# Neue Route zum Hinzufügen eines neuen Mappings zur Datei mapping_949v.csv
@app.route("/add_mapping", methods=["POST"])
def add_mapping():
    try:
        # Extrahiere die JSON-Daten aus der Anfrage
        data = request.get_json()
        print(f"Erhaltene Daten: {data}")  # Debugging-Info

        verlag = data.get("verlag")
        lieferant = data.get("lieferant")

        if not verlag or not lieferant:
            return jsonify({"error": "Beide Felder, Verlag und Lieferant, sind erforderlich."}), 400

        mapping_file_path = paths["csv_mapping_949v"]
        print(f"Mapping-Dateipfad: {mapping_file_path}")  # Debugging-Info

        # Überprüfen, ob die Datei existiert und sicherstellen, dass sie existiert
        if not os.path.exists(mapping_file_path):
            print("Die Datei mapping_949v.csv existiert nicht. Wird nun erstellt.")  # Debugging-Info
            # Erstelle die Datei und füge die Header hinzu, wenn sie nicht existiert
            with open(mapping_file_path, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(['264$b', '949$v'])  # Header hinzufügen

        # Überprüfen, ob das Mapping bereits existiert (Dublettencheck)
        mapping_exists = False
        with open(mapping_file_path, mode='r', encoding='utf-8-sig') as file:
            reader = csv.DictReader(file)
            header = reader.fieldnames  # Header lesen
            # Normalize header (e.g., to avoid whitespace or character issues)
            header = [h.strip() for h in header]
            if '264$b' not in header:
                raise ValueError(f"Header '264$b' konnte in der Datei nicht gefunden werden. Gefundene Header: {header}")

            for row in reader:
                if len(row) > 0 and row['264$b'].strip().lower() == verlag.strip().lower():
                    mapping_exists = True
                    break

        if mapping_exists:
            return jsonify({"error": "Das Mapping für diesen Verlag existiert bereits."}), 400

        # Neues Mapping hinzufügen
        with open(mapping_file_path, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow([verlag, lieferant])

        return jsonify({"message": "Mapping erfolgreich hinzugefügt."}), 200

    except Exception as e:
        print(f"Fehler: {e}")  # Debugging-Info
        return jsonify({"error": f"Fehler beim Hinzufügen des Mappings: {str(e)}"}), 500

if __name__ == "__main__":
    app.run(debug=True)

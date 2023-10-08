from flask import Flask, request, render_template, send_from_directory, redirect, url_for
import os
import zipfile
import io
from pydub import AudioSegment

app = Flask(__name__)
port = int(os.environ.get("PORT", 8080))

# Function to clear files in a folder
def clear_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")

@app.route('/')
def index():
    mp3_files = [f for f in os.listdir('downloads') if f.endswith('.mp3')]
    mp3_files.sort()  # Sort the list of MP3 files alphabetically
    return render_template('index.html', mp3_files=mp3_files)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']
    if file.filename == '':
        return "No selected file"

    if file:
        # Clear the downloads folder before converting
        clear_folder('downloads')

        # Save the uploaded PowerPoint file
        file.save(os.path.join('uploads', 'presentation.pptx'))

        # Extract audio files from the ZIP
        audio_files = []
        with zipfile.ZipFile(os.path.join('uploads', 'presentation.pptx'), 'r') as zip_ref:
            for file_info in zip_ref.infolist():
                if file_info.filename.startswith('ppt/media/') and (file_info.filename.endswith('.mp3') or file_info.filename.endswith('.m4a')):
                    audio_data = zip_ref.read(file_info)
                    audio_extension = os.path.splitext(file_info.filename)[-1].lower()
                    audio_files.append((file_info.filename, audio_data, audio_extension))

        # Convert audio files to MP3
        for i, (filename, data, audio_extension) in enumerate(audio_files):
            audio = AudioSegment.from_file(io.BytesIO(data), format=audio_extension[1:])
            mp3_filename = f'{file.filename[:-5]}_{i + 1}.mp3'
            audio.export(os.path.join('downloads', mp3_filename), format="mp3")

        return redirect(url_for('conversion_complete'))  # Redirect to conversion_complete page

@app.route('/conversion_complete')
def conversion_complete():
    return render_template('conversion_complete.html')

@app.route('/play/<filename>', methods=['GET'])
def play(filename):
    return send_from_directory('downloads', filename, as_attachment=True)

@app.route('/clear_all', methods=['POST'])
def clear_all():
    clear_folder('uploads')
    clear_folder('downloads')
    return redirect(url_for('index'))

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('downloads', exist_ok=True)
    app.run(debug=True, port=port)

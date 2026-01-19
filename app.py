from flask import Flask, render_template, send_from_directory
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static/images'  # folder foto stan

# Route utama → menampilkan halaman visualisasi
@app.route('/')
def index():
    return render_template('visualisasi_data.html')

# Route untuk menampilkan foto dari folder static/images
@app.route('/images/<filename>')
def serve_image(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == "__main__":
    app.run(debug=True)

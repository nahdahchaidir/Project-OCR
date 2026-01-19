#!/usr/bin/env python3
"""
Server sederhana untuk melayani file foto dan Excel
Jalankan: python server_simple.py
"""

import http.server
import socketserver
import os
from pathlib import Path

PORT = 8000

class SimpleHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    
    def end_headers(self):
        # Tambahkan CORS headers
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        super().end_headers()
    
    def do_GET(self):
        # Handle permintaan khusus untuk file foto
        if self.path.startswith('/foto/'):
            file_path = self.path[6:]  # Hapus '/foto/'
            
            # Cari file di folder-folder foto
            folders_to_check = ['2_images', '3_scan_output', '.']
            
            for folder in folders_to_check:
                full_path = os.path.join(folder, file_path)
                if os.path.exists(full_path):
                    self.serve_file(full_path)
                    return
            
            self.send_error(404, f"File not found: {file_path}")
            return
        
        # Serve file HTML dan lainnya seperti biasa
        super().do_GET()
    
    def serve_file(self, file_path):
        try:
            with open(file_path, 'rb') as f:
                content = f.read()
            
            self.send_response(200)
            
            # Tentukan Content-Type berdasarkan ekstensi
            if file_path.endswith('.jpg') or file_path.endswith('.jpeg'):
                self.send_header('Content-Type', 'image/jpeg')
            elif file_path.endswith('.png'):
                self.send_header('Content-Type', 'image/png')
            elif file_path.endswith('.gif'):
                self.send_header('Content-Type', 'image/gif')
            else:
                self.send_header('Content-Type', 'application/octet-stream')
            
            self.end_headers()
            self.wfile.write(content)
            
        except Exception as e:
            self.send_error(500, f"Internal server error: {str(e)}")

def main():
    # Buat folder jika belum ada
    for folder in ['2_images', '3_scan_output']:
        Path(folder).mkdir(exist_ok=True)
        print(f"✓ Folder '{folder}' siap")
    
    # Jalankan server
    with socketserver.TCPServer(("", PORT), SimpleHTTPRequestHandler) as httpd:
        print(f"\n🚀 Server berjalan di http://localhost:{PORT}")
        print(f"📁 Website: http://localhost:{PORT}/5 - Fix - Visualisasi_data.html")
        print(f"\n📁 Struktur folder:")
        print(f"  • 2_images/       - Foto asli dari download")
        print(f"  • 3_scan_output/  - Hasil scan NEG")
        print(f"\n📸 Contoh URL foto:")
        print(f"  http://localhost:{PORT}/foto/321500871587_202601_1.jpg")
        print(f"\nTekan Ctrl+C untuk menghentikan server")
        print("=" * 60)
        
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n\nServer dihentikan")

if __name__ == "__main__":
    main()
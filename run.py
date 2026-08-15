from app import app
from waitress import serve
import os

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"✅ Servidor iniciado en http://0.0.0.0:{port}")
    serve(app, host='0.0.0.0', port=port, threads=8, channel_timeout=120)
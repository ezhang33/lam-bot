from flask import Flask
from threading import Thread
import os
import time
import urllib.request
import urllib.error

app = Flask('')

@app.route('/')
def home():
    return "Hello, LamBot is alive!", 200

@app.route('/health')
def health():
    return "OK", 200

@app.route('/ping')
def ping():
    return "pong", 200

def run():
    # Render requires using the PORT environment variable
    port = int(os.getenv('PORT', 8080))
    print(f"🌐 Starting webserver on port {port}...")
    try:
        app.run(host='0.0.0.0', port=port, use_reloader=False, debug=False, threaded=True)
    except Exception as e:
        print(f"❌ Webserver error: {e}")

def ping_self():
    """Periodically ping the health endpoint to keep the service alive"""
    def ping_loop():
        time.sleep(10)  # Wait for server to start
        port = int(os.getenv('PORT', 8080))
        url = f"http://localhost:{port}/health"
        
        while True:
            try:
                with urllib.request.urlopen(url, timeout=5) as response:
                    if response.getcode() == 200:
                        print("💓 Health check ping successful")
                    else:
                        print(f"⚠️ Health check returned {response.getcode()}")
            except Exception as e:
                print(f"⚠️ Health check ping failed: {e}")
            
            # Ping every 5 minutes to keep service alive (Render free tier sleeps after 15 min)
            time.sleep(300)  # 5 minutes
    
    ping_thread = Thread(target=ping_loop, daemon=True)
    ping_thread.start()
    print("✅ Self-ping thread started")

def keep_alive():
    """Start the webserver in a separate thread and ping it periodically"""
    t = Thread(target=run, daemon=True)
    t.start()
    print("✅ Webserver thread started")
    
    # Start self-pinging to keep service alive
    ping_self()
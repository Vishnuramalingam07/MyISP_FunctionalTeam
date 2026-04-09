"""
MyISP Internal Tools - Production Server
Uses Waitress WSGI server for better performance and reliability
"""
from waitress import serve
from app import app
import os

if __name__ == '__main__':
    host = '0.0.0.0'  # Listen on all interfaces
    port = 8000
    
    print("\n" + "="*80)
    print("🌐 MyISP Internal Tools Server - PRODUCTION MODE")
    print("="*80)
    print(f"\n✓ Server starting with Waitress WSGI...")
    print(f"✓ Listening on: {host}:{port}")
    print(f"✓ Local access:  http://localhost:{port}")
    print(f"✓ Team access:   http://192.168.1.2:{port}")
    print(f"\n⚠️  Running in PRODUCTION mode - optimized for 24/7 operation")
    print(f"✓ Multi-threaded (4 threads)")
    print(f"✓ Better error handling")
    print(f"✓ Improved performance")
    print("\n" + "="*80 + "\n")
    
    # Serve the Flask app using Waitress
    serve(
        app,
        host=host,
        port=port,
        threads=4,  # Number of worker threads
        url_scheme='http',
        ident='MyISP-Tools'
    )

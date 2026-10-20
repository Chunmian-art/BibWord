import ssl
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
import os
os.chdir(Path(__file__).resolve().parent)
server=ThreadingHTTPServer(('127.0.0.1',8765),SimpleHTTPRequestHandler)
ctx=ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
ctx.load_cert_chain('localhost.pem','localhost-key.pem')
server.socket=ctx.wrap_socket(server.socket,server_side=True)
print('BibWord running at https://localhost:8765',flush=True)
server.serve_forever()

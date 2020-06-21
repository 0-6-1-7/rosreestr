from http.server import HTTPServer, CGIHTTPRequestHandler
from socketserver import ThreadingMixIn
import threading

class ThreadingSimpleServer(ThreadingMixIn, HTTPServer):
    pass

CaptchaServerMultiThreaded = ThreadingSimpleServer(("", 8111), CGIHTTPRequestHandler)
CaptchaServerMultiThreaded.serve_forever()

from http.server import HTTPServer, CGIHTTPRequestHandler

server_address = ("", 8111)
CaptchaServer = HTTPServer(server_address, CGIHTTPRequestHandler)
CaptchaServer.serve_forever()

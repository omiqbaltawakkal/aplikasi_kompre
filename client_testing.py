
import socket

server_ip = "10.233.79.96"
server_port = 54318


socketObject = socket.socket()
socketObject.connect((server_ip, server_port))
print("Connected to localhost")

HTTPMessage = "[MESSAGE HERE]"
bytes = str.encode(HTTPMessage)

socketObject.sendall(bytes)

while(True):
    data = socketObject.recv(1024)
    print(data)
    print("Connection closed")
    break

socketObject.close()

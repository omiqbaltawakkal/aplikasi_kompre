import socket
import json
import subprocess
from base64 import b64encode
import struct

image_filepath_list = ['C:/Users/iqbal/Downloads/task manager.JPG', 'C:/Users/iqbal/Downloads/delhivery.jpg', 'C:/Users/iqbal/Downloads/aramex big red.png']
report_collection = dict()

report_collection.update({"name":"iqbal"})

def upload_file(collection_key, image_file_path):
    f = open(image_file_path, "rb")
    im_bytes = f.read()
    base64_bytes = b64encode(im_bytes)
    my_file = base64_bytes.decode("utf-8")
    report_collection.update({collection_key: my_file})

# def saveImagePath():
    # f_types = [('Jpg Files', '*.jpg'),
    #            ('PNG Files', '*.png')]
    # filename = fd.askopenfilename(multiple=False)
# f = open('taskmanager.jpg', "rb")
# im_bytes = f.read()
# base64_bytes = b64encode(im_bytes)
# myFile = base64_bytes.decode("utf-8")
# report_collection.update({"image": myFile})


# saveImagePath()
for image_files in image_filepath_list:
    key = "image_"+(str(image_files).rsplit("/", maxsplit=1)[-1]).split(".")[0]
    upload_file(key, image_files)
    
print(report_collection.keys())


server_ip = "127.0.0.1"
server_port = 1000

# file_data = json.dumps(report_collection).encode("utf-8")

# socketObject = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
# socketObject.connect((server_ip, server_port))
# socketObject.sendall(struct.pack(">I", len(file_data)))
# socketObject.sendall(file_data)
#     # file_data = file.read(file)

# while True:
#     data = socketObject.recv(1024)
#     print(data)
#     print("Connection closed")
#     break

# socketObject.close()

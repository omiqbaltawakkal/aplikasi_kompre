from tkinter import *   # use lower case tkinter for python3
from threading import Thread
from time import sleep,time
import socket

class Gui():
    def __init__(self, master):
        self.running = 0    #not listening
        self.addr = None
        self.conn = None

        self.frame = Frame(master)
        self.frame.pack(side = LEFT, anchor = N)
        self.startb = Button(self.frame, text = "Start", command = self.startc)
        self.startb.pack(side = LEFT, anchor = N)
        self.stopb = Button(self.frame, text = "Stop", command = self.stopc)
        self.stopb.pack(side = LEFT, anchor = N)
        self.connectionl = Label(self.frame, text = "not connected")
        self.connectionl.pack(side = LEFT, anchor = SW)

    def socket_thread(self):
        print("thread started..")
        self.ls = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        port = 9999
        self.ls.bind(('', port))
        print("Server listening on port %s" %port)
        self.ls.listen(1)
        self.ls.settimeout(5)
        self.conn=None
        while self.running != 0:
            if self.conn is None:
                try:
                    (self.conn, self.addr) = self.ls.accept()
                    print("client is at", self.addr[0], "on port", self.addr[1])
                    self.connectionl.configure(text="CONNECTED!")
                except Exception as E:
                    print("Connect exception: "+str(E) )

            if self.conn != None:
                print ("connected to "+str(self.conn)+","+str(self.addr))
                self.conn.settimeout(5)
                self.rc = ""
                connect_start = time() # actually, I use this for a timeout timer
                while self.rc != "done":
                    self.rc=''
                    try:
                        self.rc = self.conn.recv(1000).decode('utf-8')

                    except Exception as E:
                        # we can wait on the line if desired
                        print ("socket error: "+repr(e))

                    if len(self.rc):
                        print("got data", self.rc)
                        self.conn.send("got data.\n")
                        connect_start=time()  # reset timeout time
                    elif (self.running==0) or (time()-connect_start > 30):
                        print ("Tired of waiting on connection!")
                        self.rc = "done"

                print ("closing connection")
                self.connectionl.configure(text="not connected.")
                self.conn.close()
                self.conn=None
                print ("connection closed.")

        print ("closing listener...")
        # self running became 0
        self.ls.close()

    def startc(self):
        if self.running == 0:
            print ("Starting thread")
            self.running = 1
            self.thread=Thread(target=self.socket_thread)
            self.thread.start()
        else:
            print ("thread already started.")

    def stopc(self):
        if self.running:
            print ("stopping thread...")
            self.running = 0
            self.thread.join()
        else:
            print ("thread not running")   

root = Tk()
gui = Gui(root)
root.mainloop()
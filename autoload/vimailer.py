#! /usr/bin/python3
# -*- coding: utf-8 -*-

import sys
import json
import os

from PyQt5.QtWidgets import *
from PyQt5 import uic

import imaplib
from imapclient import imap_utf7
import email

import subprocess
import re

def decode_header(msg, property_name):
    content = email.header.decode_header(msg.get(property_name))[0][0]
    encoding = email.header.decode_header(msg.get(property_name))[0][1]

    if encoding is None:
        return content 
    else:
        return content.decode(encoding,'replace')

class MyMainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        uic.loadUi("./mainwindow.ui", self)
        self.show()

        self.quitButton.clicked.connect(app.quit)
        self.connectButton.clicked.connect(self.connect_to_server)
        self.mailboxes.itemClicked.connect(self.get_message_list)
        self.statusbar.showMessage("Click \"Connect\"")
        self.message_per_page = 50
        self.messages.cellClicked.connect(self.get_message_body)

    def connect_to_server(self):

        self.statusbar.showMessage("Parsing account info...")
        with open('account.json') as f:
            account = json.load(f)
        try:
            passwd = subprocess.check_output(account["passwdcmd"], shell=True)
            passwd = passwd.strip().decode()
        except Exception as e:
            self.statusbar.showMessage("Decoding Password failed.")
            import traceback; traceback.print_exc()
            return

        # Connect to the server
        self.statusbar.showMessage("Connecting to server...")
        try:
            self.imap_connection = imaplib.IMAP4_SSL(account["server"], account["port"])
        except Exception as e:
            self.statusbar.showMessage("Connection to server failed.")
            import traceback; traceback.print_exc()
            return

        # Login to the server
        self.statusbar.showMessage("Authenticating...")
        try:
            typ, data = self.imap_connection.login(account["user"], passwd)
            if typ != 'OK':
                self.statusbar.showMessage("Authentication failed.")
                return
        except Exception as e:
            self.statusbar.showMessage("Authentication failed.")
            import traceback; traceback.print_exc()
            return

        self.statusbar.showMessage("Authentication succeeded.")
        mailbox_list = self.get_mailbox_list()

        # Make list of mailboxes and register to QListWidget
        for item in mailbox_list:
            decoded_name = imap_utf7.decode(item["name"].encode())
            item = QListWidgetItem(decoded_name)
            self.mailboxes.addItem(item)

        return

    def get_mailbox_list(self):
        self.statusbar.showMessage("Fetching list of mailboxes...")
        typ, tmp_mailbox_list = self.imap_connection.list()
        if typ != 'OK':
            self.statusbar.showMessage("Fetching mailbox list was failed.")
            return []
 
        mailbox_list = []
        for i, mailbox in enumerate(tmp_mailbox_list):
            flags, name = mailbox.decode().split(" \"/\" ")
            name = name.replace("\"","")
            mailbox_list.append({
                "flags": flags,
                "name" : name
                })

        self.statusbar.showMessage("Finished fetching list of mailboxes!")

        return mailbox_list

    def get_message_list(self, item):
        mailbox = imap_utf7.encode(item.text())
        self.statusbar.showMessage("Selecting mailbox...")
        typ, num_msg = self.imap_connection.select(mailbox)
        if typ != 'OK':
            self.statusbar.showMessage("Selecting mailbox was failed.")
            return

        num_msg = int(num_msg[0])
        self.statusbar.showMessage("Searching in mailbox...")
        typ, msg_ids = self.imap_connection.search(None, 'ALL')
        if typ != 'OK':
            self.statusbar.showMessage("Searching in mailbox was failed.")
            return
        if msg_ids[0] == b'':
            self.statusbar.showMessage("Mailbox is empty")
            return

        # Get message id's list as newest is first
        msg_ids = msg_ids[0].decode().split()
        num_get_msg = min(len(msg_ids), self.message_per_page)
        self.statusbar.showMessage("Fetching messages...")
        typ, tmp_data_set = self.imap_connection.fetch(msg_ids[-1]+":"+msg_ids[-num_get_msg],"(UID RFC822.HEADER)")
        if typ != 'OK':
            self.statusbar.showMessage("Getting messages failed.")
            return 

        raw_msg_data_list = []
        for i in range(num_get_msg):
            raw_msg_data_list.append(tmp_data_set[i*2])

        self.statusbar.showMessage("Finished fetching messages!")

        msg_list = []
        # parse to get uid
        regex = re.compile('UID\ [0-9]+')
        for i in range(num_get_msg):
            tmp = raw_msg_data_list[i][0].decode() # '17502 (UID 17505 RFC822.HEADER {6198}'
            tmp = regex.search(tmp).group() # 'UID 17505'
            uid = int(tmp.split()[1]) # '17505'
            raw_msg_contents = raw_msg_data_list[i][1]

            msg_list.append({
                "uid":uid,
                "contents":email.message_from_string(raw_msg_contents.decode())
            })

        mail_property = {}

        self.messages.setRowCount(self.message_per_page)

        for idx, msg in enumerate(msg_list):
            mail_property["from"] = decode_header(msg["contents"],'From')
            mail_property["subject"] = decode_header(msg["contents"],'Subject')
            self.messages.setItem(idx, 0, QTableWidgetItem(str(msg["uid"])))
            self.messages.setItem(idx, 1, QTableWidgetItem(mail_property["from"]))
            self.messages.setItem(idx, 2, QTableWidgetItem(mail_property["subject"]))
            self.messages.setHorizontalHeaderLabels(["uid","from","subject"])

    def get_message_body(self, row, column):
        print("Row %d and Column %d was clicked" % (row, column))
        item_uid = self.messages.item(row, 0).text()
        typ, data = self.imap_connection.uid('FETCH',item_uid,'RFC822')
        msg = email.message_from_string(data[0][1].decode())
        self.show_mail(msg)

    def show_mail(self, maildata):
        mail_value = {}
        mail_value["from"] = decode_header(maildata,'From')
        mail_value["subject"] = decode_header(maildata,'Subject')
        print("From:", mail_value["from"])
        print("Subject:", mail_value["subject"])

        for part in maildata.walk():
            print(part.get_content_type())

        for part in maildata.walk():
            print
            if not part.is_multipart():
                mail_value["charset"] = part.get_content_charset()
                if mail_value["charset"] == None:
                    mail_value["body"] = part.get_payload(decode=False)
                else:
                    mail_value["body"] = part.get_payload(decode=True).decode(mail_value["charset"],'replace')

                print("Content-Type:", part.get_content_type())
                print("Content-Transfer-Encoding:", part.get("Content-Transfer-Encoding"))
                print("Body:\n", mail_value["body"])

                if part.get_content_type() == "text/html":
                    print("Write HTML")
                    with open("tmp.html", "wb") as file:
                        file.write(mail_value["body"].encode())
                    self.browser.setText(mail_value["body"])
                    #  if os.name == 'nt':
                        #  os.system('explorer tmp.html')
                    #  else:
                        #  os.system('firefox tmp.html')
                print("\n")



if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = MyMainWindow()

    sys.exit(app.exec())

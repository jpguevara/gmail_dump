'''Script to download all the emails from an specific imap folder.
Was created to download  emails from gmail as .eml with attached files.
Usage:
python gmail_dump.py -u name@gmail.com

then password will be asked
'''

#!/usr/bin/env python
#-*- coding:utf-8 -*-

import imaplib
import email
from email.generator import Generator
from email.header import decode_header
import os
import getpass
import argparse
import logging

def print_progress(total, current, msg=''):
    import sys
    sys.stdout.write('\r')
    # the exact output you're looking for:

    current += 1
    cp = (float(current) / float(total)) * 100
    max_len = 20

    #100 - 20
    # cp -  i``

    i = (cp * max_len) / 100

    i = int(i)
    format_string = '[%-{}s] %d%%'.format(max_len)

    if msg == '':
        msg = str(current) + '/' + str(total)
    line = format_string % ('=' * i, cp) + ' ' + msg

    sys.stdout.write(line)
    sys.stdout.flush()
    if cp == 100:
        print ('')
        print ('done!')


def prepare_folders(root_folder, email_address, mailbox_to_dump):
    print ('Preparing folders...')
    logging.info ('Preparing folders...')

    path_to_save = f'{root_folder}/{email_address}/{mailbox_to_dump}'
    # path_to_save = root_folder + '/' + email_address+'/'+mailbox_to_dump

    if not os.path.exists(email_address):
        os.makedirs(email_address)
    if not os.path.exists(path_to_save):
        os.makedirs(path_to_save)

    return path_to_save


def list_folders(conn):
    folders = get_mailboxes_names(conn)
    for folder in folders:
        print (folder)


def get_mailboxes_names(conn):
    rv, mailboxes = conn.list()
    
    if rv != 'OK':
      return []
    
    # print mailboxes
    mailbox_list = []
    for mailbox in mailboxes:
        mail_box = mailbox.split()
        # print mail_box
        mailbox_list.append(get_mailbox_name(mail_box))
    return mailbox_list


def get_mailbox_name(mb):
    mb_name = ''
    is_name = False
    for item in mb:
        itemAsUtf=item.decode('UTF-8')
        if  itemAsUtf == '"/"':
            is_name = True
            continue

        if is_name:
            mb_name += itemAsUtf

    return mb_name


def save_message(path, contentString):
    f = open(path, 'w')
    f.write(contentString)

def save_email_message(path, email_message):
  out = open(path, 'w')
  gen = Generator(out)
  
  gen.flatten(email_message,)


def process_message(id, msg_data, target_path):
    response_part = msg_data[0]
    # email_body = response_part[1].decode('UTF-8')
    emailMessage = email.message_from_bytes(response_part[1])
    subject = emailMessage['subject']
    fromEmail = emailMessage['from']
    toEmail = emailMessage['to']

    # bytes, encoding = decode_header(subject)[0]

    if subject == None:
      subject = 'NO SUBJECT'
        
    # if encoding != None:
    #   subject = bytes.decode(encoding).encode("UTF-8")
    
    subject = sanitize_subject(subject)
    filename= build_filename_for_email(id, subject)
    file_path = f'{target_path}/{filename}'

    try:
      save_email_message(file_path, emailMessage)
    except :
      logging.error (f'error saving email id:{id} - subject:{subject}')
    
def build_filename_for_email(id, subject):
  # filename format id 5 digits - from address - subject
  filename = f'{int(id):0>5}'
  # filename = f'{filename} -{from_address} - {subject}'
  filename = filename.replace("/", "")  # remove all the / chars
  filename = filename[:100] + '.eml'  # trim the filename to 100 characters
  return filename
  
def sanitize_subject(subjectString):
  if type(subjectString) is str:
      return subjectString.replace(':', '').replace('/', '')
  return 'No Subject'

def read_emails(conn, remote_folder, target_path, starting_id):
    print ('Reading messages...')
    logging.info ('Reading messages...')
    # <-- pass the name of a mailbox
    stat, msgCount = conn.select(remote_folder)
    stat, data = conn.search(None, 'All')

    if stat != 'OK':
      print ('Error connecting to the inbox.')
      logging.error ('Error connecting to the inbox.')
      return
        
    ids = data[0].split()
    totalIds = len(ids)

    count = 0
    for id in ids:

        print_progress(totalIds, count)
        if int(id) < starting_id:
            count += 1
            continue
        try:
            stat, msg_data = conn.fetch(id, '(RFC822)')
            process_message(id, msg_data, target_path)
            count += 1
        except Exception as ex:
            logging.error (f'Processing email id:{id}')
            logging.error (ex)


def main():
    argparser = argparse.ArgumentParser(
        description="Dump a Gmail folder into .eml files")
    argparser.add_argument('-u', dest='username',
                           help="Gmail username", required=True)
    argparser.add_argument(
        '-s', dest='host', help="Gmail host, like imap.gmail.com", default="imap.gmail.com")
    argparser.add_argument(
        '--port', dest='port', help="Gmail port, like imap.gmail.com 993", default=993)
    argparser.add_argument('-p', dest='password', help="Gmail password")
    argparser.add_argument('-r', dest='remote_folder',
                           help="Remote folder to download", default='[Gmail]/Todos')
    argparser.add_argument('-l', dest='local_folder',
                           help="Local folder where to save .eml files", default='.')
    argparser.add_argument('--list-folders', dest='list_folders',
                           help="List all folders", default=False, action='store_true')
    argparser.add_argument('--startingId', dest='startingId',
                           help="Starting Id", default=14651, type=int)

    args = argparser.parse_args()

    host = args.host
    port = args.port
    starting_id = args.startingId

    email_address = args.username
    local_folder = args.local_folder
    remote_folder = args.remote_folder
    password = args.password

    if password is None:
        password = getpass.getpass()

    logging.basicConfig(filename=F'{email_address}.log',level=logging.DEBUG)
  

    print (F'Connecting to {email_address} folder {remote_folder}')
    logging.info(F'Connecting to {email_address} folder {remote_folder}')
    logging.info(F'StartingId {starting_id}')
    conn = imaplib.IMAP4_SSL(host, port)
    conn.login(email_address, password)

    if args.list_folders:
        print ('Listing folders.')
        list_folders(conn)
    else:
        try:
            target_path = prepare_folders(
                local_folder, email_address, remote_folder)
            read_emails(conn, remote_folder, target_path, starting_id)
            conn.close()
        except Exception as ex:
            logging.error(ex)
    conn.logout()

if __name__ == '__main__':
    main()

'''
This library contains all tools to retrieve and send new e-mails
'''



from exchangelib import Credentials, Account
import yaml
import os
def get_new_forms(fname=999):
    '''
    get_new_forms(fname=int)
    Fetches new Emails from Server and stores new Orderdata in ./Formulare.
    After the Attachments got saved the respective emails get moved to the öffentlich folder in the inbox
    '''
    user, passwd=get_exchange_cred()
    credentials = Credentials(user, passwd)
    try: 
        account = Account('XXX@XXX.XX', credentials=credentials, autodiscover=True)
        unread = account.inbox.filter()   # returns all mails in inbox

    except Exception as e: print('\nEntweder sind die Login Credentials falsch, oder es besteht keine Internetverbindung\n Nachdem beheben der Probleme kann erneut ein Versuch gestartet werden: \n{}'.format(e))
    else:
        for msg in unread:
            hashlist= forms2hash()
            for attachment in msg.attachments:
                if attachment.name=='FormData_FormData.csv':
                    fpath = os.path.join('Formulare', '{}.csv'.format(fname))
                    with open(fpath, 'wb') as f:
                        f.write(attachment.content)
                    new_hash=get_hash(fpath)
                    if  new_hash in hashlist.keys():
                        print('Bestellung Nr: {} und {} sind Doppelt'.format(hashlist[new_hash],fname))
                        to_folder = account.inbox / '20 WS' / 'öffentlichDuplikate'
                        msg.move(to_folder)
                        os.rename(fpath, fpath[:10]+'dup_'+hashlist[new_hash][12:])
                    else:
                        to_folder = account.inbox / '20 WS' / 'öffentlich'
                        msg.move(to_folder)
                        fname+=1

def get_exchange_cred():
    '''
    Returns a Tuple (user, passwd) with username and password for access to XXX@XXX.XX
    '''
    try: 
        with open('./config/credentials.yaml', 'r') as file:
            conf = yaml.safe_load(file)
            user= conf['user']['user']
            passwd=conf['user']['passwd']
            return user, passwd
    except:print('Die credentials.conf Datei wurde nicht gefunden. \nBitte stellen Sie sicher, dass diese entsprechend der Anleitung erstellt wurde')


from exchangelib import Credentials, Account, Message, Mailbox
def create_mail(body, subject,  recipient):
    '''
    create_mail(body=str, subject= str, recipient=str)
    Returns exchangelib message object, which can be send by calling .send() on it
    '''
    user, passwd=get_exchange_cred()
    credentials = Credentials(user, passwd)
    account = Account('XXX@XXX.XX', credentials=credentials, autodiscover=True)
    m=Message(
    account=account,
    subject=subject,
    body=body,
    to_recipients=[
        Mailbox(email_address=recipient),
    ])
    return m

import hashlib
def get_hash(filepath):
    with open(filepath, 'rb') as inputfile:
        data = inputfile.read()
        return hashlib.md5(data).hexdigest()



import glob
def forms2hash(path='./Formulare/*.csv'):
    '''
    returns a dictionary with an hash as its key values and the corresponding Filename as its value
    '''
    filenames = glob.glob(path)
    hashlist={}
    for filename in filenames:
        hash=get_hash(filename)
        hashlist[hash]=filename
    return hashlist

def compose_orderconf_mail(mail_data):
    '''
    Does what the methodes name suggests.
    Returns a tuple (body, subject, recepient-adress)
    '''
    with open('./config/Bestellbestätigung.txt') as file:
        rows=file.readlines()
        body=''
        for row in rows:
            body=body+row
        anrede=mail_data[-1]+' '+mail_data[1]
        sano=mail_data[5]
        saer=mail_data[6]
        sono=mail_data[7]
        soer=mail_data[8]
        best_nr=mail_data[0]
        preis=(sano+sono)*15+(saer+soer)*10
        body=body.format(anrede=anrede,sano=sano,saer=saer,sono=sono,soer=soer,best_nr=best_nr,preis=preis)
        subject='Konzertkartenbestellung Hochschule München, {best_nr}'.format(best_nr=best_nr)
        mail=mail_data[3]


    return body, subject, mail

def compose_transferconf_mail(mail_data):
    with open('./config/Zahlungseingang.txt') as file:
        rows=file.readlines()
        body=''
        for row in rows:
            body=body+row
        anrede=mail_data[-1]+' '+mail_data[1]
        best_nr=mail_data[0]
        mail=mail_data[3]
        body=body.format(anrede=anrede,best_nr=best_nr, email=mail)
    with open('./Email-Zahlung/Zahlungseingang{}.txt'.format(best_nr),'w') as file:
            file.write(body)
 
    

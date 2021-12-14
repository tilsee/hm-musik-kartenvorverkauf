import config.okvv_mail_lib as okvv_mail
import config.okvv_excel_lib as okvv_excel
import os
import optparse
import time

def main(send_mail=False):
    #load excelsheet
    try:
        wb, fname=okvv_excel.load_excel()
        ws=wb['Kartenbestellungen']
    except: raise Exception("The Excelfile was not found. Template can be found in config directory")
    #set next order number
    next_order_nr=ws.max_row-2
    
    #variable to exit programm after beeing idle for some time
    iteration_counter=0
    
    #getting the path folder seperator for os paths
    dir_sep=okvv_excel.get_os_dir_sep()
    
    #fetching new order data from konzertkarten@hm.edu
    okvv_mail.get_new_forms(next_order_nr)
    
    #cache for emails
    messages_data=[]
    
    
    while True:
        if iteration_counter==10:
            break
            
        #processing of new forms, stored in ./Formulare
        if '{}.csv'.format(next_order_nr) in os.listdir('Formulare'):
            
            #extracting data from order form
            data=list(okvv_excel.data_extractor('{}.csv'.format(next_order_nr)).values())
            
            #generating greeting for emails
            if data[0]=='Frau':
                anrede='Sehr geehrte Frau'
            else: 
                anrede='Sehr geehrter Herr'
                
            #adding the ordernumber to the ordernumber
            data[0]='C'+'{}'.format(next_order_nr).zfill(3)
            #add the new order to excel
            okvv_excel.add_new_orders(data,ws)
            
            #adding data for generating the emails
            data.append(anrede)
            messages_data.append(data)
            
            
            #saving the changes made to excel
            wb.save(fname)
            
            #updating next ordernumber
            next_order_nr+=1
            
            #resetting idle counter
            iteration_counter=0
        
        iteration_counter+=1
        
    #Generating Email Text and sending mail
    for message_data in messages_data:
        body, subject, mail=okvv_mail.compose_orderconf_mail(message_data)

        if send_mail:
            okvv_mail.create_mail(body=body,subject=subject,recipient=mail).send()
            print('Email {} wurden versendet'.format(subject[-4:]))
        else:
            print('Emails werden derzeit nicht versandt. Zum Versenden muss das Script mit "python okvv_v2.py --send_mail" gestartet werden!')

    
        #Saving Mails for debuging purpose
        dir_sep=okvv_excel.get_os_dir_sep()
        with open('.{}Emails{}{}.txt'.format(dir_sep, dir_sep,message_data[0]),'w') as file:
            file.write(mail+'\n\n'+subject+'\n\n'+body)
        okvv_mail.compose_transferconf_mail(message_data)
    
if __name__=='__main__':
    print('Programm läuft')
    parser = optparse.OptionParser()
    parser.set_defaults(send_mail=False)
    parser.add_option('--send_mail', action='store_true', dest='send_mail')
    parser.set_defaults(auto=False)
    parser.add_option('--auto', action='store_true', dest='auto')
    parser.add_option('--xls', action='store_true', dest='xls')
    (options, args) = parser.parse_args()

    while True:
        main(options.send_mail)
        if options.auto != True:
            print('Soll erneut nach neuen Formularen geschaut werden mit "Ja" bestätigen oder mit "Nein" abbrechen:')
            
            if input().lower().startswith('j')==False:
                print('Programm beendet')
                break
        if options.auto:
            print("Run was successfull")
            time.sleep(300)
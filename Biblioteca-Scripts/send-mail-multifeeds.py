import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
import traceback
import pandas as pd
import os

def send_email_with_attachment(csv_files,mail,sujeto,tramo,copia):
    
    email_address = 'mail@mail.com'
    recipient_email = f'{mail}'
    cc_email = copia.split(';')


    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = recipient_email
    today = datetime.today().strftime("%d_%m_%Y")
    msg['Subject'] = f'{sujeto} ' + today

    if cc_email:
        msg['Cc'] = ', '.join(cc_email)

        

    body = f'''
    <html>
        <body>
        <a href="https://gobik.com/"><img src="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/Logo_Gobik.png" width="180" height="60"/></a></br></br>
        <p style="font-family: 'Arial Nova'; font-size: 18px"><mark>FEEDS PRODUCTS {tramo} - Archivos enviados automáticamente</mark><br></p>

                    <footer style="display: flex;align-items: center;max-width: 500px;font-size: 10px;color: gray;">
                        <p style="font-family: 'Gill Sans', 'Gill Sans MT', Calibri, 'Trebuchet MS', sans-serif;font-size: 10px;text-align: justify;">
                            <b>PROTECCIÓN DE DATOS: </b>Los datos personales que forman parte de este correo electrónico son tratados por GOBIK SPORT WEAR S.L.
                            con la finalidad de mantenimiento de contactos. Los datos se han obtenido con su consentimiento o como consecuencia de una relación jurídica previa.
                            Puede usted ejercitar los derechos de acceso, rectificación, cancelación y oposición, así como obtener más información solicitándolo al remitente de este correo electrónico.
                        Este mensaje y los ficheros anexos son confidenciales. Los mismos contienen información reservada de la empresa que no puede ser difundida. En caso de no 
                        ser el destinatario de esta información, por favor, rogamos nos lo comunique en la dirección del remitente para la eliminación de su dirección electrónica, 
                        no copiando ni entregando este mensaje a nadie más y procediendo a su destrucción.</p>
                    </footer>
        </body>
    </html>
    '''
    

    for csv_file_path, csv_file_name in csv_files:
        with open(csv_file_path, 'rb') as attachment:
            csv_part = MIMEApplication(attachment.read(), Name=csv_file_name)
            csv_part['Content-Disposition'] = f'attachment; filename={csv_file_name}'
            msg.attach(csv_part)


    smtp_server = 'smtp-es.securemail.pro'
    smtp_port = 465
    email_password = 'password-mail'
    msg.attach(MIMEText(body, 'html'))
    server = smtplib.SMTP_SSL(smtp_server, smtp_port)
    server.login(email_address, email_password)
    to_email = [recipient_email] + cc_email

    server.sendmail(email_address, to_email, msg.as_string())
    server.quit()

def csv_to_xlsx(base_path_T1, base_path_T2, base_path_T3):
    
    csv_files = [
        (rf'{base_path_T1}\PRODUCTS-T1.csv', rf'{base_path_T1}\PRODUCTS-T1.xlsx'),
        (rf'{base_path_T2}\DE-T2.csv', rf'{base_path_T2}\DE-T2.xlsx'),
        (rf'{base_path_T2}\EN-T2.csv', rf'{base_path_T2}\EN-T2.xlsx'),
        (rf'{base_path_T2}\ES-T2.csv', rf'{base_path_T2}\ES-T2.xlsx'),
        (rf'{base_path_T2}\FR-T2.csv', rf'{base_path_T2}\FR-T2.xlsx'),
        (rf'{base_path_T2}\IT-T2.csv', rf'{base_path_T2}\IT-T2.xlsx'),
        (rf'{base_path_T3}\EN-T3.csv', rf'{base_path_T3}\EN-T3.xlsx'),
        (rf'{base_path_T3}\ES-T3.csv', rf'{base_path_T3}\ES-T3.xlsx')
    ]
    
    for csv_name, xlsx_name in csv_files:
        # Construir las rutas completas de los archivos CSV y XLSX
        csv_path = os.path.join( csv_name)
        xlsx_path = os.path.join(xlsx_name)
        
        # Leer el archivo CSV y guardarlo como XLSX
        df = pd.read_csv(csv_path)
        df.to_excel(xlsx_path, index=False)
        

def main():

    today = datetime.today().strftime("%d_%m_%Y")
    
    base_path_T1 = r'\\IP\shopify_multifeeds\FEEDS_PRODUCTS_T1'
    base_path_T2 = r'\\IP\shopify_multifeeds\FEEDS_PRODUCTS_T2'
    base_path_T3 = r'\\IP\shopify_multifeeds\FEEDS_PRODUCTS_T3'

    csv_to_xlsx(base_path_T1, base_path_T2, base_path_T3)
    
    Excel_files_A = [
        (rf'{base_path_T2}\DE-T2.xlsx', 'DE-T2_'+today+'.xlsx'),
        (rf'{base_path_T2}\EN-T2.xlsx', 'EN-T2_'+today+'.xlsx'),
        (rf'{base_path_T2}\ES-T2.xlsx', 'ES-T2_'+today+'.xlsx'),
        (rf'{base_path_T2}\FR-T2.xlsx', 'FR-T2_'+today+'.xlsx'),
        (rf'{base_path_T2}\IT-T2.xlsx', 'IT-T2_'+today+'.xlsx'),
        (rf'{base_path_T3}\EN-T3.xlsx', 'EN-T3_'+today+'.xlsx'),
        (rf'{base_path_T3}\ES-T3.xlsx', 'ES-T3_'+today+'.xlsx')
    ]
    
    Excel_files_B = [(rf'{base_path_T1}\PRODUCTS-T1.xlsx', 'PRODUCTS-T1_'+today+'.xlsx')]
    
    
    mailA = 'mail-to' #TRADEINN
    cc_copiaAA = 'copy-mail-to'
    mailB = 'mail-to' #SERVIBIKE
    cc_copiaBB = 'copy-mail-to'
    sujetoA = 'Archivo CSV TradeInn - MultiFeeds'
    sujetoB = 'Archivo CSV ServiBike - MultiFeeds'
    tramoA = '(T2--T3)'
    tramoB = '(T1)'
    
    # VARIABLE DE TIEMPO PARA REALIZAR ARCHIVO DE LOG's
    fecha_hora_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    try:
        send_email_with_attachment(Excel_files_A,mailA,sujetoA,tramoA,cc_copiaAA)
        print(f'{fecha_hora_actual} - INFO - Mail Successfully Sent.\n')
    except Exception:
        print(f'{fecha_hora_actual} - ERROR - Mail Failed Sent.\n')
        with open('C:\Tareas Programadas\TAREAS MULTIFEEDS\event-ERROR_history_Multifeed_Email.log', 'a') as archivo:
            archivo.write(f'{fecha_hora_actual} - ERROR - Error al enviar mensaje Multifeeds A\n')
            archivo.write(traceback.format_exc())
        
        
    try:
        send_email_with_attachment(Excel_files_B,mailB,sujetoB,tramoB,cc_copiaBB)
        print(f'{fecha_hora_actual} - INFO - Mail Successfully Sent.\n')
    except Exception:
        print(f'{fecha_hora_actual} - ERROR - Mail Failed Sent.\n')
        with open('C:\Tareas Programadas\TAREAS MULTIFEEDS\event-ERROR_history_Multifeed_Email.log', 'a') as archivo:
            archivo.write(f'{fecha_hora_actual} - ERROR - Error al enviar mensaje Multifeeds B\n')
            archivo.write(traceback.format_exc())
        
                
if __name__ == '__main__':
    main()

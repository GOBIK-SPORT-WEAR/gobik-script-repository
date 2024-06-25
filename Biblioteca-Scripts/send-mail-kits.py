from hdbcli import dbapi
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')
import locale
locale.setlocale(locale.LC_ALL, 'es_ES')
import traceback



# DEFINICION PARA ENVIO DE CORREO
def send_email_with_attachment(doc, cantidad, trak, idioma, tiempo, country, e_mail_user, comercial, query):
    
    # DEFINO LOS DATOS PARA ENVIAR EL CORREO, REMITENTE, DESTINATARIO, EN COPIA...
    
    email_address = 'mail@mail.com'
    recipient_email = ''#f'{e_mail_user}'
    cc_email = ''#f'copy-mail-to;{comercial}'

    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = recipient_email
    if cc_email:
        msg['Cc'] = cc_email
        
    #ASUNTO
    msg['Subject'] = f'KIT GOBIK - {doc}'

    # AQUI ES DONDE METEMOS LA CONSULTA B EN UN DATAFRAME PARA PODER DARLE FORMATO TABLA Y DARLE UN ESTILO
    tabla_SQL_QueryB = pd.DataFrame(query)
    html_content = tabla_SQL_QueryB.to_html(index=False, justify='center')                      
    html_content = html_content.replace('<table border="1" class="dataframe">','<table style="border-collapse: collapse; border: 1px solid black;">')
    html_content = html_content.replace('<th>', '<th style="border: 1px solid black; padding: 5px; padding-right:10px; text-align: left;">')
    html_content = html_content.replace('<td>', '<td style="border: 1px solid black; padding: 5px; padding-right:10px; text-align: left;">')
    
    # CUERPO DEL MENSAJE, QUE SEGUN EL IDIOMA DEL CLIENTE APLICA UNO U OTRO
    if idioma == 23 and country == 'ES': #ESPAÑOL
        
        body = f'''
            <html>
                <body>
                    <a href="https://gobik.com/"><img src="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000LOGO_GOBIK_CUSTOM_WORKS.jpg?v=1707923377"/></a></br></br></br>


                    <p style="font-family: 'Arial Nova'; font-size: 18px">Hola,<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">En GOBIK tenemos buenas noticias para ti: tu <b>kit de tallas</b> ha sido enviado.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">El número de referencia GOBIK es: <b>{doc}</b>. Utilízalo para referenciar todas tus comunicaciones y la identificación para su devolución.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Puedes hacer el <b>seguimiento del envío</b> del mismo en este <a href="{trak}"><u>Link</u></a><br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px;">Recuerda que tienes hasta el próximo día: <b style="color:red;">{tiempo}</b> para realizar la devolución.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Por favor, revisa las instrucciones de devolución adjuntas en este link: <a href="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000GOBIK_DEVOLUCION_ES.pdf"><u>INSTRUCCIONES-GOBIK-DEVOLUCIÓN.pdf</u></a>.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">También tienes disponible nuestro <a href="https://gobik.com/pages/virtual-fitting-kit"><u>Kit de Tallas Virtual</u></a> disponible 24/7 donde podrás consultar tu talla mediante unas simples medidas.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px"><b>Los artículos que han sido enviados son los siguientes:</b><br></p>
                       
                    {html_content}
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">NOTA: Estándar (XS, S, M, L, XL y 2XL); Grandes (3XL y 4XL); Especiales (5XL y 6XL); Pequeñas (2XS)<br></p>              
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Queremos que sepas que el kit de tallas que has recibido tiene un valor estimado de <b>{cantidad}</b> €. 
                    Este servicio te lo ofrecemos <b>completamente gratuito</b>, con el objetivo de que puedas seleccionar las tallas perfectas sin preocupaciones. Consideramos importante 
                    que seas consciente del valor de los productos enviados, reflejo de nuestro compromiso con la calidad y tu satisfacción.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Si detectas que alguna de las prendas que figura en el listado anterior no se corresponde con las recibidas, 
                    te rogamos que nos lo notifiques lo antes posible a kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Recuerda que las prendas que componen este kit tienen una función exclusiva: ayudarte a escoger tu talla 
                    GOBIK de forma correcta; cualquier otro uso (entrenamientos, competición, etc.) <b>no está autorizado</b> y debes evitarlo.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Si tienes cualquier duda al respecto, no dudes en escribirnos al correo kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">¡Muchas gracias!<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">GOBIK<br></p>

                    
                    
                


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
    
    elif idioma == 23 and country != 'ES': #RESTO IDIOMA ESPAÑOL
        
        body = f'''
            <html>
                <body>
                    <a href="https://gobik.com/"><img src="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000LOGO_GOBIK_CUSTOM_WORKS.jpg?v=1707923377"/></a></br></br></br>

                    <p style="font-family: 'Arial Nova'; font-size: 18px">Hola,<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">En GOBIK tenemos buenas noticias para ti: tu <b>kit de tallas</b> ha sido enviado.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">El número de referencia GOBIK es: <b>{doc}</b>. Utilízalo para referenciar todas tus comunicaciones y la identificación para su devolución.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Puedes hacer el <b>seguimiento del envío</b> del mismo en este <a href="{trak}"><u>Link</u></a><br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px;">Recuerda que tienes hasta el próximo día: <b style="color:red;">{tiempo}</b> para realizar la devolución.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Por favor, revisa las instrucciones de devolución adjuntas en este link: <a href="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000GOBIK_DEVOLUCION_ES_EXT.pdf"><u>INSTRUCCIONES-GOBIK-DEVOLUCIÓN.pdf</u></a>.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">También tienes disponible nuestro <a href="https://gobik.com/pages/virtual-fitting-kit"><u>Kit de Tallas Virtual</u></a> disponible 24/7 donde podrás consultar tu talla mediante unas simples medidas.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px"><b>Los artículos que han sido enviados son los siguientes:</b><br></p>
                       
                    {html_content}
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">NOTA: Estándar (XS, S, M, L, XL y 2XL); Grandes (3XL y 4XL); Especiales (5XL y 6XL); Pequeñas (2XS)<br></p>              
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Queremos que sepas que el kit de tallas que has recibido tiene un valor estimado de <b>{cantidad}</b> €. 
                    Este servicio te lo ofrecemos <b>completamente gratuito</b>, con el objetivo de que puedas seleccionar las tallas perfectas sin preocupaciones. Consideramos importante 
                    que seas consciente del valor de los productos enviados, reflejo de nuestro compromiso con la calidad y tu satisfacción.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Si detectas que alguna de las prendas que figura en el listado anterior no se corresponde con las recibidas, 
                    te rogamos que nos lo notifiques lo antes posible a kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Recuerda que las prendas que componen este kit tienen una función exclusiva: ayudarte a escoger tu talla 
                    GOBIK de forma correcta; cualquier otro uso (entrenamientos, competición, etc.) <b>no está autorizado</b> y debes evitarlo.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Si tienes cualquier duda al respecto, no dudes en escribirnos al correo kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">¡Muchas gracias!<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">GOBIK<br></p>

                    
                    
                


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
    
    elif idioma == 3 and country == 'DE': #ALEMAN
        
        body = f'''
            <html>
                <body>
                    <a href="https://gobik.com/"><img src="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000LOGO_GOBIK_CUSTOM_WORKS.jpg?v=1707923377"/></a></br></br></br>

                    <p style="font-family: 'Arial Nova'; font-size: 18px">Hallo,<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Bei GOBIK haben wir gute Neuigkeiten für Sie: Ihr <b>Größen-Kit</b> wurde versendet.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Die GOBIK-Referenznummer lautet: <b>{doc}</b>. Bitte verwenden Sie diese für alle Ihre Mitteilungen und zur Identifikation für Rücksendungen.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Sie können <b>Ihre Sendung</b> mit diesem Link verfolgen: <a href="{trak}"><u>Link</u></a><br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px;">Denken Sie daran, Sie haben bis zum nächsten: <b style="color:red;">{tiempo}</b> um es zurückzusenden.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Bitte überprüfen Sie die Rücksendeanweisungen in diesem Link: <a href="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000GOBIK_DEVOLUCION_DE.pdf"><u>ANLEITUNG-GOBIK-RÜCKKEHR.pdf</u></a>.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Sie haben auch rund um die Uhr unser <a href="https://gobik.com/de/pages/virtual-sizing-kit"><u>Virtuelles Größen-Kit</u></a> zur Verfügung. Dort können 
                    Sie Ihre Größe mit einigen einfachen Maßen überprüfen.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px"><b>Die gesendeten Artikel sind wie folgt:</b><br></p>
                       
                    {html_content}
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">NOTIZ: Estandar (XS, S, M, L, XL und 2XL); Grandes (3XL und 4XL); Especiales (5XL und 6XL); Pequeñas (2XS)<br></p>              
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Wir möchten, dass Sie wissen, dass das Größen-Kit, das Sie erhalten haben, einen Wert von <b>{cantidad}</b> €. 
                    hat. Wir bieten Ihnen diesen Service völlig kostenlos an, um Ihnen die Auswahl der perfekten Größen ohne Sorgen zu ermöglichen. Wir glauben, es ist wichtig, dass Sie 
                    sich des Werts der gesendeten Produkte bewusst sind, was unser Engagement für Qualität und Ihre Zufriedenheit widerspiegelt.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Wenn Sie feststellen, dass eines der oben aufgeführten Artikel nicht mit den erhaltenen übereinstimmt, 
                    bitten wir Sie höflich, uns dies so schnell wie möglich unter kits@gobik.com mitzuteilen.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Denken Sie daran, dass die Kleidungsstücke in diesem Kit nur einem Zweck dienen: Ihnen bei der
                    korrekten Auswahl Ihrer GOBIK-Größe zu helfen. Jede andere Verwendung (Training, Wettkampf usw.) <b>ist nicht autorisiert</b> und sollte vermieden werden.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Wenn Sie Fragen dazu haben, zögern Sie bitte nicht, uns unter kits@gobik.com zu kontaktieren.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Vielen Dank!<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">GOBIK<br></p>

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
    
    elif idioma == 22: #FRANCES
        
        body = f'''
            <html>
                <body>
                    <a href="https://gobik.com/"><img src="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000LOGO_GOBIK_CUSTOM_WORKS.jpg?v=1707923377"/></a></br></br></br>

                    <p style="font-family: 'Arial Nova'; font-size: 18px">Bonjour,<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Chez GOBIK, nous avons une bonne nouvelle pour vous : votre <b>kit de tailles</b> a été envoyé.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Le numéro de référence GOBIK est : <b>{doc}</b>. Veuillez l'utiliser pour toutes vos communications et pour l'identification des retours.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Vous pouvez suivre votre <b>envoi en utilisant</b> ce lien: <a href="{trak}"><u>Link</u></a><br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px;">Rappelez-vous, vous avez jusqu'au: <b style="color:red;">{tiempo}</b> pour le retourner.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Veuillez consulter les instructions de retour sur ce lien: <a href="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000GOBIK_DEVOLUCION_FR.pdf"><u>INSTRUCTIONS-GOBIK-RETOUR.pdf</u></a>.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Vous avez également à votre disposition toujours notre <a href="https://gobik.com/fr/pages/virtual-sizing-kit"><u>Kit de Tailles Virtuel</u></a>. Là, vous pouvez vérifier votre taille à l'aide de quelques mesures simples.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px"><b>Les articles qui ont été envoyés sont les suivants:</b><br></p>
                       
                    {html_content}
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">NOTE: Estandar (XS, S, M, L, XL y 2XL); Grandes (3XL et 4XL); Especiales (5XL et 6XL); Pequeñas (2XS)<br></p>              
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Nous tenons à vous informer que le kit de tailles que vous avez reçu est évalué à <b>{cantidad}</b> €. Nous 
                    vous offrons ce service <b>entièrement gratuitement</b>, dans le but de vous permettre de choisir les tailles parfaites sans aucun souci. Nous jugeons important que vous 
                    soyez conscient de la valeur des produits envoyés, reflétant notre engagement envers la qualité et votre satisfaction.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Si vous remarquez que certains des articles listés ci-dessus ne correspondent pas à ceux reçus, 
                    nous vous prions de bien vouloir nous en informer dès que possible à kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Rappelez-vous que les vêtements inclus dans ce kit ont un seul but : vous aider à choisir votre 
                    taille GOBIK correctement ; tout autre usage (entraînement, compétition, etc.) <b>n'est pas autorisé</b> et doit être évité.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Si vous avez des questions à ce sujet, n'hésitez pas à nous envoyer un email à kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Merci beaucoup!<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">GOBIK<br></p>

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
    
    elif idioma == 13: #ITALIANO
        
        body = f'''
            <html>
                <body>
                    <a href="https://gobik.com/"><img src="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000LOGO_GOBIK_CUSTOM_WORKS.jpg?v=1707923377"/></a></br></br></br>

                    <p style="font-family: 'Arial Nova'; font-size: 18px">Ciao,<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Da GOBIK abbiamo buone notizie per te: il tuo <b>kit di taglie</b> è stato inviato.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Il numero di riferimento GOBIK è: <b>{doc}</b>. Si prega di utilizzarlo per tutte le tue comunicazioni e per l'identificazione per la restituzione.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Puoi <b>tracciare la tua spedizione</b> utilizzando questo link: <a href="{trak}"><u>Link</u></a><br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px;">Ricorda, hai tempo fino al prossimo: <b style="color:red;">{tiempo}</b> per restituirlo.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Si prega di rivedere le istruzioni per la restituzione in questo link: <a href="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000GOBIK_DEVOLUCION_IT.pdf"><u>ISTRUZIONI-GOBIK-RITORNO.pdf</u></a>.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Hai anche a disposizione 24/7, il nostro <a href="https://gobik.com/it/pages/virtual-sizing-kit"><u>Kit Virtuale per le Taglie</u></a>. Lì puoi controllare la tua taglia utilizzando alcune misurazioni semplici.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px"><b>Gli articoli che sono stati inviati sono i seguenti:</b><br></p>
                       
                    {html_content}
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">NOTA: Estandar (XS, S, M, L, XL e 2XL); Grandes (3XL e 4XL); Especiales (5XL e 6XL); Pequeñas (2XS)<br></p>              
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Vogliamo che tu sappia che il kit di taglie che hai ricevuto ha un valore di <b>{cantidad}</b> €. Offriamo 
                    questo servizio <b>completamente gratuito</b>, con l'obiettivo di permetterti di selezionare le taglie perfette senza preoccupazioni. Riteniamo sia importante che tu sia 
                    consapevole del valore dei prodotti inviati, riflettendo il nostro impegno per la qualità e la tua soddisfazione.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Se noti che alcuni degli articoli elencati sopra non corrispondono a quelli ricevuti,
                    ti chiediamo gentilmente di notificarcelo il prima possibile a kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Ricorda che i capi inclusi in questo kit hanno un unico scopo: aiutarti a scegliere 
                    correttamente la tua taglia GOBIK; qualsiasi altro uso (allenamento, competizione, ecc.) <b>non è autorizzato</b> e dovrebbe essere evitato.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Se hai domande a riguardo, non esitare a inviarci un'email a kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Grazie mille!<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">GOBIK<br></p>

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
        
    else:           #RESTO ENGLISH

        body = f'''
            <html>
                <body>
                    <a href="https://gobik.com/"><img src="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000LOGO_GOBIK_CUSTOM_WORKS.jpg?v=1707923377"/></a></br></br></br>

                    <p style="font-family: 'Arial Nova'; font-size: 18px">Hi,<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">At GOBIK, we have good news for you: your <b>size kit</b> has been Sent.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">The GOBIK reference number is: <b>{doc}</b>. Please use it for all your communications and for identification for returns.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">You can <b>track your shipment</b> using this link:  <a href="{trak}"><u>Link</u></a><br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px;">Remember, you have until the next: <b style="color:red;">{tiempo}</b> to return it.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Please review the return instructions in this link: <a href="https://cdn.shopify.com/s/files/1/0052/7876/1029/files/000GOBIK_DEVOLUCION_EN.pdf"><u>INSTRUCTIONS-GOBIK-RETURN.pdf</u></a>.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">You also have available 24/7 our <a href="https://gobik.com/en/pages/virtual-sizing-kit"><u>Virtual Size Kit.</u></a> There you can check your size using some simple measurements.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px"><b>The items that have been sent are as follows:</b><br></p>
                       
                    {html_content}
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">NOTE: Estandar (XS, S, M, L, XL and 2XL); Grandes (3XL and 4XL); Especiales (5XL and 6XL); Pequeñas (2XS)<br></p>              
                    <p style="font-family: 'Arial Nova'; font-size: 18px">We want you to know that the sizing kit you have received is valued at <b>{cantidad}</b> €. We offer this service 
                    to you <b>completely free of charge</b>, with the aim of enabling you to select the perfect sizes without any worries. We believe it's important for you to be aware of the value 
                    of the products sent, reflecting our commitment to quality and your satisfaction.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">If you notice that any of the items listed above do not match those received, we kindly ask you to notify us as 
                    soon as possible at kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Remember that the garments included in this kit serve a sole purpose: to help you choose your GOBIK size correctly; 
                    any other use (training, competition, etc.) <b>is not authorised</b> and should be avoided.<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">If you have any questions about this, please do not hesitate to email us at kits@gobik.com<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">Thank you very much!<br></p>
                    
                    <p style="font-family: 'Arial Nova'; font-size: 18px">GOBIK<br></p>

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
    
    # SE DEFINEN TODOS LOS DATOS DEL SERVIDOR DE CORREO
    smtp_server = 'smtp-es.securemail.pro'
    smtp_port = 465
    email_password = 'password-mail'
    msg.attach(MIMEText(body, 'html'))
    server = smtplib.SMTP_SSL(smtp_server, smtp_port)
    server.login(email_address, email_password)
    to_email = [recipient_email]
    if cc_email:
        to_email.append(cc_email)

    # CON TODA LA INFO REALIZA EL ENVIO
    server.sendmail(email_address, to_email, msg.as_string())
    server.quit()


# DEFINICION PRINCIPAL QUE LLAMA A LA DEFINICION DE ENVIO DE CORREO
def main():

    # CONEXION CON LA BBDD DE SAP
    conn = dbapi.connect(
        address='IP',
        port='PORT',
        user='USER',
        password='Password',
        ENCRYPT=True,
        sslValidateCertificate=False
    )

    # CONSULTA PRINCIPAL (EXTRAE INFORMACION DE LAS ENTREGAS DE KIT DE TALLAS EN EL DIA DE HOY) - PAIS(cliente) - IDIOMA(cliente) - U
    sql_queryA = pd.read_sql_query('SELECT T2."Email", T1."Country", T0."LangCode", T0."U_IFG_IntTCangopal_Tracking_Urls", T0."DocTotal", T0."DocNum", add_DAYS(T0."DocDate",15) AS "DaysRemained", T1."E_Mail" FROM "GB".ODLN T0  INNER JOIN "GB".OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN "GB".OSLP T2 ON T0."SlpCode" = T2."SlpCode" WHERE T0."U_IFG_IntTCangopal_Tracking_Urls" IS NOT NULL AND T0."U_IFG_TIPOPEDIDO" = \'KIT DE TALLAS\' AND T0."DocDate"=CURRENT_DATE', conn)

    # METEMOS DICHO QUERY EN UN DATAFRAME PARA EXTRAER LOS DATOS
    tabla_SQL_QueryA = pd.DataFrame(sql_queryA)
    
    # VARIABLE DE TIEMPO PARA REALIZAR ARCHIVO DE LOG's
    fecha_hora_actual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # RECORREMOS LINEA A LINEA DEL DATAFRAME Y EXTRAEMOS LA INFO QUE QUEREMOS Y ENVIAMOS UN CORREO POR CADA LINEA DEL QUERY
    for index, row in tabla_SQL_QueryA.iterrows():
        try:
            country_num = row['Country']
            doc_num = row['DocNum']
            doc_total = row['DocTotal']
            trak_url = row['U_IFG_IntTCangopal_Tracking_Urls']
            lang_code = row['LangCode']
            days_remained = row['DaysRemained']
            days_remained = datetime.strftime(days_remained, '%d / %m / %Y')
            e_mail_var = row['E_Mail']
            e_mail_COMERCIAL = row['Email']
            
            # CONSULTA QUE EXTRAE TODOS LOS ARTICULOS QUE TIENE CADA ENTREGA SEGUN Nº DE DOCUMENTO DE CADA LINEA DE LA PRIMERA CONSULTA
            sql_queryB = pd.read_sql_query(f'SELECT T2."ItemCode" AS "REFERENCE", T2."Dscription" AS "DESCRIPTION", T2."Quantity" AS "QUANTITY", CONCAT(TO_DECIMAL(T2."LineTotal",10,2),\' €\') AS "VALUE" FROM "GB".ODLN T0  INNER JOIN "GB".OCRD T1 ON T0."CardCode" = T1."CardCode" INNER JOIN "GB".DLN1 T2 ON T0."DocEntry" = T2."DocEntry" WHERE T0."DocNum"=\'{doc_num}\' ', conn)
            num_filas = sql_queryB.shape[0]
            
            # FILTRAMOS PARA QUE NO REALICE ENVIO CUANDO SOLO SEA UNA PRUEBA DE COLOR
            if not ((sql_queryB == '90-01-010-002-00').any().any() and num_filas == 1):
                # AHORA LLAMAMOS A LA DEFINICION PARA ENVIAR EL CORREO, METIENDO POR PARAMETROS LO QUE NECESITAMOS
                send_email_with_attachment(doc_num, doc_total, trak_url, lang_code, days_remained, country_num, e_mail_var, e_mail_COMERCIAL, sql_queryB)
                # SI LO HA ENVIADO CORRECTAMENTE AÑADE UNA LINEA A UN ARCHIVO LOG
                with open('C:\Tareas Programadas\RECORDATORIO EMAIL KIT DE TALLAS\event_history_Aviso1.log', 'a') as archivo:
                    archivo.write(f'{fecha_hora_actual} - INFO - Mail successfully sent to {e_mail_var} --- DocNum:{doc_num}\n')

        except Exception:
            e_mail_var = row['E_Mail']
            doc_num = row['DocNum']
            # SI FALLA EL ENVIO AÑADE UNA LINEA A UN ARCHIVO LOG CON LA TODA LA INFO
            with open('C:\Tareas Programadas\RECORDATORIO EMAIL KIT DE TALLAS\event_history_Aviso1.log', 'a') as archivo:
                archivo.write(f'{fecha_hora_actual} - ERROR - Error al enviar mensaje a {e_mail_var} --- DocNum:{doc_num}\n')
                archivo.write(traceback.format_exc())
                                
    conn.close()
 
if __name__ == '__main__':
    main()

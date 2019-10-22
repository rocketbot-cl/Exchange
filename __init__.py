# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"
    
    pip install <package> -t .

"""
try:

    import shelve
    import os
    from exchangelib import Account, Credentials, DELEGATE, Configuration, NTLM, BASIC, ServiceAccount
    from exchangelib.folders import Message, Mailbox
    from exchangelib import FileAttachment

    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter

    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
    from bs4 import BeautifulSoup

    """
        Obtengo el modulo que fue invocado
    """
    module = GetParams("module")
    global server
    global cred
    global address
    global config
    global a

    """
        Obtengo variables
    """
    if module == "exchange":
        global server
        global cred
        global address
        global config

        user = GetParams('user')
        password = GetParams('pass')
        server = GetParams('server')
        address = GetParams('address')
        print('USUARIO', user)

        cred = Credentials(username=user, password=password)
        config = Configuration(
            server=server,
            credentials=cred
        )

        print(config)

        """
            Obtengo la ruta con y sin extensión
        """

    if module == "send_mail":
        to = GetParams('to')
        subject = GetParams('subject')
        body = GetParams('body')
        attached_file = GetParams('attached_file')

        config = Configuration(
            server=server,
            credentials=cred
        )

        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        # print('TREE',a.root.tree())

        # If you want a local copy
        m = Message(
            account=a,
            folder=a.sent,
            subject=subject,
            body=body,
            to_recipients=to.split(",")
        )

        if attached_file is None or attached_file is "":
            pass
        else:
            with open(attached_file, 'rb') as f:
                content = f.read()  # Read the binary file contents

                attached_name = os.path.basename(attached_file)

            att = FileAttachment(name=attached_name, content=content)
            m.attach(att)
        m.save()

        m.send_and_save()

    if module == "get_mail":
        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        var = GetParams('var')
        tipo_filtro = GetParams('tipo_filtro')
        filtro = GetParams('filtro')
        id = []

        if filtro:
            if tipo_filtro == 'author':
                #id = [m.id for m in a.inbox.all() if not m.is_read and filtro in m.author.email_address]

                for m in a.inbox.all():
                    if filtro in m.author.email_address:
                        id.append(m.id)
                        #print('FOR',m.id)

            if tipo_filtro == 'subject':
                #id = [m.id for m in a.inbox.all() if tmp in m.subject]

                for m in a.inbox.all():
                    if filtro in m.subject:
                        id.append(m.id)
                        #print('FOR',m.id)
        else:
            id = [m.id for m in a.inbox.all() if not m.is_read]

        SetVar(var, id)

    if module == "get_new_mail":
        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        var = GetParams('var')
        tipo_filtro = GetParams('tipo_filtro')
        filtro = GetParams('filtro')
        id = []

        if filtro:
            if tipo_filtro == 'author':
                #id = [m.id for m in a.inbox.all() if not m.is_read and filtro in m.author.email_address]

                for m in a.inbox.all():
                    if not m.is_read and filtro in m.author.email_address:
                        id.append(m.id)
                        #print('FOR',m.id)

            if tipo_filtro == 'subject':
                #id = [m.id for m in a.inbox.all() if tmp in m.subject]

                for m in a.inbox.all():
                    if not m.is_read and filtro in m.subject:
                        id.append(m.id)
                        #print('FOR',m.id)

        else:
            id = [m.id for m in a.inbox.all() if not m.is_read]

        SetVar(var, id)

    if module == "read_mail":
        path_ = GetParams('path')
        if path_:
            path_ = os.path.normpath(path_)

            if not os.path.exists(path_):
                raise Exception('La carpeta no existe')

        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        id_mail = GetParams('id_mail')
        var = GetParams('var')

        mail = a.inbox.get(id=id_mail)

        datos_mail = mail.subject, mail.author.email_address, mail.body, mail.attachments
        cont = BeautifulSoup(mail.body, "html")
        final = {'subject': mail.subject, 'from': mail.author.email_address, 'body': cont.text.strip(),
                 'files': [m.name for m in mail.attachments]}
        print(final)

        SetVar(var, final)

        # DESCARGA ARCHIVO ADJUNTO

        if path_:

            for attachment in mail.attachments:
                fpath = os.path.join(path_, attachment.name)
                with open(fpath, 'wb') as f:
                    f.write(attachment.content)

    if module == "move_folder":
        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        folder_name = GetParams('folder_name')
        id_mail = GetParams('id_mail')

        to_folder = a.inbox / folder_name

        mail = a.inbox.get(id=id_mail)

        mail.move(to_folder)

except Exception as e:
    PrintException()
    print(e)
    raise e

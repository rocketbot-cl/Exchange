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
    from exchangelib import Account, DELEGATE, NTLM, BASIC, ServiceAccount, HTMLBody
    from exchangelib.folders import Message, Mailbox
    from exchangelib import FileAttachment

    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter

    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
    from bs4 import BeautifulSoup

    """
        Obtengo el modulo que fue invocado
    """
    module = GetParams("module")

    global exchange_module
    global a


    class ExchangeModule:
        def __init__(self, user_, pwd, server_, mail_):
            self.pwd = pwd
            self.mail = mail_
            self.server = server_
            self.user = user_
            self.credentials = None
            self.config = None

        def init(self):
            from exchangelib import Credentials, Configuration
            self.credentials = Credentials(username=self.user, password=self.pwd)
            self.config = Configuration(server=self.server, credentials=self.credentials)
            return self.config


    """
        Obtengo variables
    """
    if module == "exchange":
        user = GetParams('user')
        password = GetParams('pass')
        server = GetParams('server')
        address = GetParams('address')
        print('USUARIO', user)

        exchange_module = ExchangeModule(user, password, server, address)
        config = exchange_module.init()
        # print(config)

    if module == "send_mail":
        to = GetParams('to')
        subject = GetParams('subject')
        body = GetParams('body')
        is_html = GetParams('isHtml')
        cc = GetParams("cc")
        attached_file = GetParams('attached_file')
        attached_folder = GetParams("attached_folder")

        if exchange_module.config is not None:
            config = exchange_module.config
        else:
            raise Exception("Execute Email Configuration command")
        address = exchange_module.mail

        if is_html:
            body = HTMLBody(body)

        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        # print('TREE',a.root.tree())

        # If you want a local copy
        m = Message(
            account=a,
            folder=a.sent,
            subject=subject,
            body=body,
            to_recipients=to.split(","),
            cc_recipients=cc.split(",") if cc else None
        )
        att = []
        file_names = []
        if attached_file: file_names.append(attached_file)
        if attached_folder:
            for f in os.listdir(attached_folder):
                f = os.path.join(attached_folder, f)
                file_names.append(f)

        for file in file_names:
            with open(file, 'rb') as f:
                content = f.read()  # Read the binary file contents
                attached_name = os.path.basename(file)

            att.append(FileAttachment(name=attached_name, content=content))

        if len(att) > 0: m.attach(att)
        m.save()

        m.send_and_save()

    if module == "get_mail":

        if exchange_module.config is not None:
            config = exchange_module.config
        else:
            raise Exception("Execute Email Configuration command")
        address = exchange_module.mail
        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        var = GetParams('var')
        tipo_filtro = GetParams('tipo_filtro')
        filtro = GetParams('filtro')
        id_ = []

        if bool(filtro):
            print("Searching with filter")
            if tipo_filtro == 'author':
                # id = [m.id for m in a.inbox.all() if not m.is_read and filtro in m.author.email_address]

                for m in a.inbox.all():
                    if filtro in m.author.email_address:
                        id_.append(m.id)
                        # print('FOR',m.id)

            if tipo_filtro == 'subject':
                # id = [m.id for m in a.inbox.all() if tmp in m.subject]

                for m in a.inbox.all():
                    if filtro in m.subject:
                        id_.append(m.id)
                        # print('FOR',m.id)
        else:
            print("Searching without filter")
            id_ = [m.id for m in a.inbox.all()]

        SetVar(var, id_)

    if module == "get_new_mail":

        if exchange_module.config is not None:
            config = exchange_module.config
        else:
            raise Exception("Execute Email Configuration command")
        address = exchange_module.mail
        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        var = GetParams('var')
        tipo_filtro = GetParams('tipo_filtro')
        filtro = GetParams('filtro')
        id_ = []

        if filtro:
            if tipo_filtro == 'author':
                # id = [m.id for m in a.inbox.all() if not m.is_read and filtro in m.author.email_address]

                for m in a.inbox.all():
                    if not m.is_read and filtro in m.author.email_address:
                        id_.append(m.id)
                        # print('FOR',m.id)

            if tipo_filtro == 'subject':
                # id = [m.id for m in a.inbox.all() if tmp in m.subject]
                print("subject:" + filtro)
                mails_filters = a.inbox.filter(subject__contains=filtro)
                for m in mails_filters:
                    print(m.subject)
                    if not m.is_read:
                        id_.append(m.id)
                        # print('FOR',m.id)

        else:
            id_ = [m.id for m in a.inbox.all() if not m.is_read]

        SetVar(var, id_)

    if module == "read_mail":
        path_ = GetParams('path')
        if path_:
            path_ = os.path.normpath(path_)

            if not os.path.exists(path_):
                raise Exception('La carpeta no existe')

        if exchange_module.config is not None:
            config = exchange_module.config
        else:
            raise Exception("Execute Email Configuration command")
        address = exchange_module.mail

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
        mail.is_read = True
        mail.save()

        SetVar(var, final)

        # DESCARGA ARCHIVO ADJUNTO

        if path_:

            for attachment in mail.attachments:
                fpath = os.path.join(path_, attachment.name)
                with open(fpath, 'wb') as f:
                    f.write(attachment.content)

    if module == "move_folder":
        if exchange_module.config is not None:
            config = exchange_module.config
        else:
            raise Exception("Execute Email Configuration command")

        address = exchange_module.mail
        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        folder_name = GetParams('folder_name')
        id_mail = GetParams('id_mail')

        to_folder = a.inbox / folder_name

        mail = a.inbox.get(id=id_mail)

        mail.move(to_folder)

    if module == "reply_email":
        # Get Params from Rocketbot
        id_mail = GetParams('id_')
        body_ = GetParams('body')
        attached_file = GetParams('attached_file')
        attached_folder = GetParams('attached_folder')
        reply_all = GetParams("all")

        # Validation of configuration
        if exchange_module.config is not None:
            config = exchange_module.config
        else:
            raise Exception("Execute Email Configuration command")

        address = exchange_module.mail

        # Get Account
        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        # Get Mail
        mail = a.inbox.get(id=id_mail)

        if reply_all:
            to_recipients = mail.to_recipients + [mail.sender]
            cc_recipients = mail.cc_recipients
        else:
            to_recipients = [mail.sender]
            cc_recipients = []

        # Create Draft to reply
        m = mail.create_reply(
            subject=mail.subject,
            body=body_,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients
        )

        # Attachment
        att = []
        file_names = []
        if attached_file:
            file_names.append(attached_file)
        if attached_folder:
            for f in os.listdir(attached_folder):
                f = os.path.join(attached_folder, f)
                file_names.append(f)

        for file in file_names:
            with open(file, 'rb') as f:
                content = f.read()  # Read the binary file contents
                attached_name = os.path.basename(file)

            att.append(FileAttachment(name=attached_name, content=content))

        if len(att) > 0:
            m.attach(att)
        # m.save()
        m.send()

    if module == "forward":
        id_ = GetParams('id_')
        to_ = GetParams('email')
        attached_file = GetParams('attached_file')

        # Validation of configuration
        if exchange_module.config is not None:
            config = exchange_module.config
        else:
            raise Exception("Execute Email Configuration command")

        address = exchange_module.mail

        # Get Account
        a = Account(primary_smtp_address=address, config=config,
                    access_type=DELEGATE, autodiscover=False)

        # Get Mail
        mail = a.inbox.get(id=id_)

        # Create Draft to reply
        m = mail.create_forward(
            subject="Forward: " + mail.subject,
            body="",
            to_recipients=to_.split(";"),
            cc_recipients=mail.cc_recipients,
        )

        # Attachment
        att = []
        file_names = []
        if attached_file:
            file_names.append(attached_file)
        # if attached_folder:
        #     for f in os.listdir(attached_folder):
        #         f = os.path.join(attached_folder, f)
        #         file_names.append(f)

        for file in file_names:
            with open(file, 'rb') as f:
                content = f.read()  # Read the binary file contents
                attached_name = os.path.basename(file)

            att.append(FileAttachment(name=attached_name, content=content))

        if len(att) > 0:
            m.attach(att)
        # m.save()
        m.send()


except Exception as e:
    PrintException()
    raise e

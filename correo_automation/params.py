# -*- coding: utf-8 -*-
"""
Created on Wed Feb 11 08:02:42 2026

@author: USUARIO
"""

class EmailParams:

    #por defecto se consultan 10 correos máximo
    def __init__(
        self,
        email_objetivo = None,
        max_emails=10,
        iniciar_desde = 1,
        sender=None,
        from_date=None
    ):
        self.iniciar_desde = iniciar_desde
        self.email_objetivo = email_objetivo
        self.max_emails = max_emails
        self.sender = sender
        self.from_date = from_date

#ejemplo de usuaro personalizado:
# params = EmailParams(
#     max_emails=20,
#     sender="notificaciones@rama judicial.gov.co"
# )

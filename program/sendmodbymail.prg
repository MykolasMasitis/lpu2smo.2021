FUNCTION SendModByMail(lcmcod)
 
 m.mcod = lcmcod
 lcPeriod = STR(tYear,4) + PADL(tMonth,2,'0')
 m.DocName   = pout+'\'+lcPeriod+'\Pm' + m.mcod
 
 IF !fso.FolderExists(pOut)
  MESSAGEBOX('���������� '+pOut+' �����������!',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pOut+'\'+lcPeriod)
  MESSAGEBOX('���������� '+pOut+'\'+lcPeriod+' �����������!',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.DocName+'.doc')
  MESSAGEBOX('�� ���������� ���'+CHR(13)+CHR(10)+;
   '����������� � ������������ �� ������������!'+CHR(13)+CHR(10),0+16,m.Docname)
  RETURN 
 ENDIF 
 
 IF SEEK(m.mcod, 'emails')
  m.email = ALLTRIM(emails.e_mail)
 ELSE 
  m.email = ''
 ENDIF 
 
 IF EMPTY(m.email)
  MESSAGEBOX('��� ���������� ���'+CHR(13)+CHR(10)+;
   '����� ����������� ����� �� ���������!'+CHR(13)+CHR(10),0+16,m.Docname)
  RETURN 
 ENDIF 

 #Define MAPI_ORIG 0
 #Define MAPI_TO 1
 #Define MAPI_CC 2
 #Define MAPI_BCC 3

 #Define IMPORTANCE_LOW 0
 #Define IMPORTANCE_NORMAL 1
 #Define IMPORTANCE_HIGH 2

 #DEFINE CRCR CHR(13)

 LOCAL lcMessage, llSuccess
	
 m.lcSubject = "����������� � ������������"
 m.lcMessage = "������ ����!"+CHR(13)+CHR(10)
 m.lcMessage = m.lcMessage + "�� �������� ���������� ����������� � ������������"+CHR(13)+CHR(10)
 m.lcMessage = m.lcMessage + "�� "+NameOfMonth(tMonth)+ ' '+STR(tYear,4)+' ����'+CHR(13)+CHR(10)
 m.lcMessage = m.lcMessage + "�� ��������� ����������� ����������� "+m.qname

 m.lcAddress    = m.email
 m.lcAttachment = m.docname
 
 IF EMCreateMessage(m.lcSubject, m.lcMessage, IMPORTANCE_HIGH)
  IF EMAddRecipient(m.lcAddress, MAPI_TO)
   IF EMAddAttachment(m.lcAttachment)
    IF EMSend(.T.)
     m.llSuccess = .T.
    ENDIF
   ENDIF 
  ENDIF
 ENDIF
	
 IF m.llSuccess
  MESSAGEBOX("���� ��������� ���� ������� ����������!", 0+48, "�������� �������� ���������") 
 ELSE
  MESSAGEBOX("�� ������� ��������� ���������!", 64, "������ ��� ��������")
 ENDIF 

RETURN 
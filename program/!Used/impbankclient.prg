PROCEDURE ImpBankCLient

 IF MESSAGEBOX(CHR(13)+CHR(10)+;
  '�� ������ ��������� �������?'+CHR(13)+CHR(10),;
   4+32,'������ �� ����-������')==7
  RETURN 
 ENDIF 

 odir = SYS(5)+SYS(2003)
 SET DEFAULT TO (pBin)
 BCFile = GETFILE('Text:txt','','',0,'������� �� ����!')
 SET DEFAULT TO (odir)
 
 oBCFile = fso.GetFile(BCFile)
 IF oBCFile.size >= 20
  fhandl = oBCFile.OpenAsTextStream
  lcHead = fhandl.Read(20)
  fhandl.Close
 ENDIF 
 
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')

 IF lcHead == '1CClientBankExchange'
  MESSAGEBOX('������ ���������!',0+64,'')
 ELSE 
  MESSAGEBOX('����������� ������ �����!',0+64,'')
 ENDIF 

 fhandl = oBCFile.OpenAsTextStream

 nPayOrdersAll = 0
 nPayOrdersToLpu = 0

 DO WHILE !fhandl.AtEndOfStream
  cLine = fhandl.ReadLine
  IF cLine = '��������������=��������� ���������'
   nPayOrdersAll = nPayOrdersAll + 1
  ENDIF 
  IF cLine = '�����������������=' AND SEEK(SUBSTR(cLine, AT('=', cLine)+1, 7), 'sprlpu')
   nPayOrdersToLpu = nPayOrdersToLpu + 1
  ENDIF 
 ENDDO 
 fhandl.Close
 
 USE IN sprlpu
 
 MESSAGEBOX('�������� �������� '+STR(nPayOrdersAll,3)+ ' ��������, '+CHR(13)+CHR(10)+;
  '�� ��� � ����� ��� '+STR(nPayOrdersToLpu,3), 0+64, '')
  
RETURN 
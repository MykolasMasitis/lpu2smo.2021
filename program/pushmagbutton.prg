PROCEDURE PushMagButton
 IF MESSAGEBOX('�� ������ ��������� ��������������� ��� ��������?'+CHR(13)+CHR(10)+;
 	'1. ������������� ��������� ������� ���'+CHR(13)+CHR(10)+;
 	'2. �������������� ������� ���'+CHR(13)+CHR(10)+;
 	'3. �������������� ���',4+32,'')=7
  RETURN 
 ENDIF 
 
 WAIT "������������� ��������� ��..." WINDOW NOWAIT 
 DO CorStruct
 WAIT CLEAR 

 WAIT "�������������� ������� ���..." WINDOW NOWAIT 
 DO BasReind
 WAIT CLEAR 

 WAIT "�������������� ���..." WINDOW NOWAIT 
 DO ComReind
 WAIT CLEAR 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 
PROCEDURE PackBdSv
 IF MESSAGEBOX('��� ��������� ��������� �������'+CHR(13)+CHR(10)+;
  '���������� � �������� ������ ������� ������.'+CHR(13)+CHR(10)+'����������?',4+32, '')==7
  RETURN 
 ENDIF 
 
 IF OpenFile("&pBase\&gcPeriod\people", "people", "excl") == 0
  WAIT "�������� " + pBase + "\" + gcPeriod + "\" + " PEOPLE.DBF" WINDOW NOWAIT 
  SELECT people
  PACK 
  USE IN people
  WAIT CLEAR 
 ELSE 
  MESSAGEBOX('���������� ������� ���� PEOPLE!',0+16,'')
 ENDIF 

 IF OpenFile("&pBase\&gcPeriod\talon", "talon", "excl") == 0
  WAIT "�������� " + pBase + "\" + gcPeriod + "\" + " TALON.DBF" WINDOW NOWAIT 
  SELECT talon
  PACK 
  USE IN talon
  WAIT CLEAR 
 ELSE 
  MESSAGEBOX('���������� ������� ���� TALON!',0+16,'')
 ENDIF 

 IF OpenFile("&pBase\&gcPeriod\otdel", "otdel", "excl") == 0
  WAIT "�������� " + pBase + "\" + gcPeriod + "\" + " OTDEL.DBF" WINDOW NOWAIT 
  SELECT otdel
  PACK 
  USE IN otdel
  WAIT CLEAR 
 ELSE 
  MESSAGEBOX('���������� ������� ���� OTDEL!',0+16,'')
 ENDIF 

 IF OpenFile("&pBase\&gcPeriod\doctor", "doctor", "excl") == 0
  WAIT "�������� " + pBase + "\" + gcPeriod + "\" + " DOCTOR.DBF" WINDOW NOWAIT 
  SELECT doctor
  PACK 
  USE IN doctor
  WAIT CLEAR 
 ELSE 
  MESSAGEBOX('���������� ������� ���� DOCTOR!',0+16,'')
 ENDIF 
 
RETURN 
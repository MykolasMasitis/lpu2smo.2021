PROCEDURE ResetDolg
 IF MESSAGEBOX(CHR(13)+CHR(10)+'�� ������ �������� �������������?'+CHR(13)+CHR(10), '')==7
  RETURN 
 ENDIF 

 IF MESSAGEBOX(CHR(13)+CHR(10)+'�� ������� � ����� ���������?'+CHR(13)+CHR(10), '')==7
  RETURN 
 ENDIF 
 
 wasrec = RECNO()
 REPLACE ALL dolg_b WITH 0
 GO (wasrec)

RETURN 
PROCEDURE ResetAvans
 IF MESSAGEBOX(CHR(13)+CHR(10)+'�� ������ �������� ������?'+CHR(13)+CHR(10), '')==7
  RETURN 
 ENDIF 

 IF MESSAGEBOX(CHR(13)+CHR(10)+'�� ������� � ����� ���������?'+CHR(13)+CHR(10), '')==7
  RETURN 
 ENDIF 
 
 wasrec = RECNO()
 REPLACE ALL s_avans WITH 0, s_pr_avans WITH 0
 GO (wasrec)

RETURN 
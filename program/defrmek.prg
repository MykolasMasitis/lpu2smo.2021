FUNCTION DeFrMek

IF MESSAGEBOX('��� ��������� �������������'+CHR(13)+CHR(10)+;
 '���� ��� �� ��� ���.'+CHR(13)+CHR(10)+'����������?',4+32, '')==7
 RETURN 
ENDIF 

IF MESSAGEBOX(''+CHR(13)+CHR(10)+;
 '�� ��������� ������� � ����� ���������?'+CHR(13)+CHR(10)+;
 ''+CHR(13)+CHR(10),4+32, '')==7
 RETURN 
ENDIF 

IF OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'shar') > 0
 RETURN 
ENDIF 

SELECT AisOms
SCAN
 m.mcod = mcod

 WAIT m.mcod WINDOW NOWAIT 
 
 REPLACE IsPr WITH .f.

 WAIT CLEAR 
ENDSCAN 
USE IN AisOms

RETURN 
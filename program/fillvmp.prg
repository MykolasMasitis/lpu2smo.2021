PROCEDURE FillVmp
 IF MESSAGEBOX('��������� ���� VMP ����� TARIFN'+CHR(13)+CHR(10)+'��������� ���� VMP146 ����� USVMPXX?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(pcommon+'\usvmpxx.dbf')
  MESSAGEBOX('����������� ���� USVMPXX.DBF!',0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\nsi\tarifn.dbf')
  MESSAGEBOX('����������� ���� TARIFN.DBF!',0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\tarifn', 'tarif', 'shar')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\usvmpxx', 'usvmp', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('usvmp')
   USE IN usvmp
  ENDIF 
  RETURN 
 ENDIF 
 
 m.recsintar = RECCOUNT('tarif')
 m.recsinvmp = RECCOUNT('usvmp')
 
 IF m.recsintar!=m.recsinvmp
  IF MESSAGEBOX('���-�� ������� � ����� TARIFN ('+TRANSFORM(m.recsintar,'9999')+')'+CHR(13)+CHR(10)+;
   '�� ������������� '+CHR(13)+CHR(10)+;
   '���-�� ������� � ����� USVMPXX ('+TRANSFORM(m.recsintar,'9999')+')'+CHR(13)+CHR(10)+;
   '����������?',4+32,'')=7
   USE IN usvmp
   USE IN tarif
   RETURN 
  ENDIF 
 ENDIF 
 
 SELECT tarif
 SET RELATION TO cod INTO usvmp
 SCAN 
  IF EMPTY(usvmp.cod)
   LOOP 
  ENDIF 
  REPLACE vmp WITH usvmp.vmp146
 ENDSCAN 
 SET RELATION OFF INTO usvmp

 USE IN usvmp
 USE IN tarif
 MESSAGEBOX('��������� ���������!',0+64,'')
RETURN 

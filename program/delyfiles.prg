PROCEDURE DelYFiles
 IF MESSAGEBOX('����� ������� ����������?'+CHR(13)+CHR(10)+;
               '��� ��, ��� �� ������������� ������ �������?',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('�� ��������� ������� � ����� ���������?',4+48,'') != 6
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod') > 0
  RETURN
 ENDIF 
 
 SELECT AisOms
 
 SCAN
  m.mcod = mcod
  m.lpu_id = lpuid

  WAIT m.mcod WINDOW NOWAIT 

  mmy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
  bfile = 'D'+m.qcod+STR(m.lpu_id,4)+'.'+mmy
  IF fso.FileExists(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+bfile)
   fso.DeleteFile(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+bfile)
  ENDIF 


 ENDSCAN 
 WAIT CLEAR 
 USE 

RETURN 
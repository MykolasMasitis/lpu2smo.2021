PROCEDURE DelAllCtrl

 IF MESSAGEBOX('��� ��������� �������'+CHR(13)+CHR(10)+;
  '����� ������ (CtrlQQ.dbf)!'+CHR(13)+CHR(10)+'����������?',4+32, '')==7
  RETURN 
 ENDIF 

 IF MESSAGEBOX(''+CHR(13)+CHR(10)+;
  '�� ��������� ������� � ����� ���������?'+CHR(13)+CHR(10)+;
  ''+CHR(13)+CHR(10),4+32, '')==7
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'shar')>0
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\UsrLpu', "UsrLpu", "shar", "mcod") > 0
  USE IN aisoms
  RETURN
 ENDIF 

 SELECT AisOms
 
 SCAN 
  m.mcod = mcod
*  m.usr  = IIF(SEEK(m.mcod, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")
*  IF m.usr != m.gcUser AND m.gcUser!='OMS'
*   LOOP 
*  ENDIF 

  WAIT m.mcod WINDOW NOWAIT 

  lcDir = pBase + '\' + m.gcperiod + '\' + mcod
  IF !fso.FolderExists(lcDir)
   LOOP 
  ENDIF 
  IF !fso.FileExists(lcDir+'\Ctrl'+m.qcod+'.dbf')
   LOOP 
  ENDIF 

  fso.DeleteFile(lcDir+'\Ctrl'+m.qcod+'.dbf')

  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 USE IN AisOms
 USE IN UsrLpu
 
RETURN 
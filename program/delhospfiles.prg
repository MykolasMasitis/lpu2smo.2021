PROCEDURE DelHospFiles
 IF MESSAGEBOX('ÝÒÀ ÏÐÎÖÅÄÓÐÀ ÓÄÀËßÅÒ ÂÑÅ'+CHR(13)+CHR(10)+;
  'ÑÔÎÐÌÈÐÎÂÀÍÍÛÅ ÐÀÍÅÅ hosp-ÔÀÉËÛ!'+CHR(13)+CHR(10)+'ÏÐÎÄÎËÆÈÒÜ?',4+32, '')==7
  RETURN 
 ENDIF 

 IF MESSAGEBOX(''+CHR(13)+CHR(10)+;
  'ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?'+CHR(13)+CHR(10)+;
  ''+CHR(13)+CHR(10),4+32, '')==7
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+gcPeriod+'\AisOms', 'AisOms', 'shar')>0
  RETURN 
 ENDIF 

 SELECT AisOms
 
 SCAN 
  m.mcod  = mcod
  m.lpuid = lpuid

  WAIT m.mcod WINDOW NOWAIT 

  lcDir = pBase + '\' + m.gcperiod + '\' + mcod
  IF !fso.FolderExists(lcDir)
   LOOP 
  ENDIF 

  IF !fso.FileExists(lcDir+'\hosp.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(lcDir+'\hosp', 'hosp', 'shar')>0
   IF USED('hosp')
    USE IN hosp 
   ENDIF 
   SELECT aisoms 
   LOOP 
  ENDIF 
  m.n_recs = RECCOUNT('hosp')
  USE IN hosp
  
  IF m.n_recs>0
   SELECT aisoms
   LOOP 
  ENDIF 
  
  IF fso.FileExists(lcDir+'\hosp.dbf')
   fso.DeleteFile(lcDir+'\hosp.dbf')
  ENDIF 
  IF fso.FileExists(lcDir+'\hosp.cdx')
   fso.DeleteFile(lcDir+'\hosp.cdx')
  ENDIF 
  
  WAIT CLEAR 
 ENDSCAN 
 WAIT CLEAR 
 
 USE IN AisOms
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 
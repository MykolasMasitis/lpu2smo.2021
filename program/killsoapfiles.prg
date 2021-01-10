PROCEDURE KillSoapFiles

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÓÄÀËÈÒÜ SOAP-ÔÀÉËÛ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 m.lcperiod = m.gcperiod
 m.lpath = pbase+'\'+m.lcperiod
 IF !fso.FolderExists(m.lpath)
  LOOP 
 ENDIF 
 IF !fso.FileExists(m.lpath+'\aisoms.dbf')
  LOOP 
 ENDIF 
  
 WAIT m.lcperiod+'...' WINDOW NOWAIT 
 =KillOnePeriod(m.lpath)
 WAIT CLEAR 


RETURN 

FUNCTION KillOnePeriod
 PARAMETERS m.lpath
 PRIVATE m.llcpath
 m.llcpath = m.lpath
 IF OpenFile(m.llcpath+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 SELECT aisoms
 SCAN 
  m.mcod = mcod
  IF !fso.FolderExists(m.llcpath+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.llcpath+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  
  IF fso.FileExists(m.llcpath+'\'+m.mcod+'\request.http')
   fso.DeleteFile(m.llcpath+'\'+m.mcod+'\request.http')
  ENDIF 
  IF fso.FileExists(m.llcpath+'\'+m.mcod+'\request.xml')
   fso.DeleteFile(m.llcpath+'\'+m.mcod+'\request.xml')
  ENDIF 
  IF fso.FileExists(m.llcpath+'\'+m.mcod+'\answer.http')
   fso.DeleteFile(m.llcpath+'\'+m.mcod+'\answer.http')
  ENDIF 
  IF fso.FileExists(m.llcpath+'\'+m.mcod+'\answer.xml')
   fso.DeleteFile(m.llcpath+'\'+m.mcod+'\answer.xml')
  ENDIF 
  IF fso.FileExists(m.llcpath+'\'+m.mcod+'\answer.zip')
   fso.DeleteFile(m.llcpath+'\'+m.mcod+'\answer.zip')
  ENDIF 
  IF fso.FileExists(m.llcpath+'\'+m.mcod+'\data.xml')
   fso.DeleteFile(m.llcpath+'\'+m.mcod+'\data.xml')
  ENDIF 
  IF fso.FileExists(m.llcpath+'\'+m.mcod+'\soapans.dbf')
   fso.DeleteFile(m.llcpath+'\'+m.mcod+'\soapans.dbf')
  ENDIF 
  

 ENDSCAN 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 

RETURN 
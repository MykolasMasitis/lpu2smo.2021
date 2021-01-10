PROCEDURE DirSize
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÐÀÑÑ×ÈÒÀÒÜ ÐÀÇÌÅÐÛ ÄÈÐÅÊÒÎÐÈÉ?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod)
  MESSAGEBOX('ÄÈÐÅÊÒÎÐÈß '+m.pBase+'\'+m.gcPeriod+' ÍÅ ÍÀÉÄÅÍÀ!',0+64,'')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 fso = CREATEOBJECT('Scripting.FileSystemObject')

 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  
  x = fso.GetFolder(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
  m.d_size = ROUND(x.Size/(1024*1024),2)
  
  REPLACE dirsize WITH m.d_size
  
 ENDSCAN 
 
 USE IN aisoms
 RELEASE fso
 
 MESSAGEBOX('OK!',0+64,'')

RETURN 
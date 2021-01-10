PROCEDURE passEX
 IF MESSAGEBOX('ÏÅÐÅÄÀÒÜ ÔÀÉËÛ ÎØÈÁÎÊ?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pExpImp+'\'+m.gcPeriod)
  fso.CreateFolder(pExpImp+'\'+m.gcPeriod)
 ENDIF 
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FolderExists(pExpImp+'\'+m.gcPeriod+'\'+m.mcod)
   fso.CreateFolder(pExpImp+'\'+m.gcPeriod+'\'+m.mcod)
  ENDIF 
  fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf', pExpImp+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
  fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.cdx', pExpImp+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.cdx')
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\x'+m.mcod+'.dbf')
   fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\x'+m.mcod+'.dbf', pExpImp+'\'+m.gcPeriod+'\'+m.mcod+'\x'+m.mcod+'.dbf')
  ENDIF 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\x'+m.mcod+'.cdx')
   fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\x'+m.mcod+'.cdx', pExpImp+'\'+m.gcPeriod+'\'+m.mcod+'\x'+m.mcod+'.cdx')
  ENDIF 
  
 ENDSCAN 
 USE IN aisoms 
 
 MESSAGEBOX('OK!',0+64,'')
  
RETURN 
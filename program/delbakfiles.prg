PROCEDURE DelBakFiles
 IF MESSAGEBOX('ÁÓÄÓÒ ÓÄÀËÅÍÛ ÂÑÅ BAK-ÔÀÉËÛ?'+CHR(13)+CHR(10)+;
               'ÝÒÎ ÒÎ, ×ÒÎ ÂÛ ÄÅÉÑÒÂÈÒÅËÜÍÎ ÕÎÒÈÒÅ ÑÄÅËÀÒÜ?',4+48,'') != 6
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod') > 0
  RETURN
 ENDIF 
 
 SELECT AisOms
 
 m.nDeletedFiles = 0
 m.nDeletedSize  = 0
 
 SCAN
  m.mcod = mcod
  IF !fso.FolderExists(pBase+'\'+gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 

  WAIT m.mcod WINDOW NOWAIT 

  oMailDir        = fso.GetFolder(pBase+'\'+gcPeriod+'\'+m.mcod)
  oFilesInMailDir = oMailDir.Files
  nFilesInMailDir = oFilesInMailDir.Count
  
  IF nFilesInMailDir<=0
   RELEASE oMailDir, oFilesInMailDir, nFilesInMailDir
   LOOP 
  ENDIF 

  FOR EACH oFileInMailDir IN oFilesInMailDir
   m.bname = oFileInMailDir.Path
   m.ext = LOWER(RIGHT(ALLTRIM(m.bname),3))
   
   IF m.ext = 'bak'
    x = fso.GetFile(m.bname)
    m.f_size = x.Size
    fso.DeleteFile(m.bname)
    m.nDeletedFiles = m.nDeletedFiles + 1
    m.nDeletedSize  = m.nDeletedSize + m.f_size
    
    RELEASE x, m.f_size
   ENDIF 

  ENDFOR 

 ENDSCAN 
 
 MESSAGEBOX('ÓÄÀËÅÍÎ '+TRANSFORM(m.nDeletedFiles, '999999')+' ÔÀÉËÎÂ'+CHR(10)+CHR(13)+;
 	'ÎÑÂÎÁÎÆÄÅÍÎ '+TRANSFORM(ROUND(m.nDeletedSize/(1024*1024),0),'999999')+' ÌÁ',0+64,'')
 WAIT CLEAR 
 USE 

RETURN 
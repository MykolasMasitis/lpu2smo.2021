PROCEDURE Mee2Lpu

 m.pDir = pOut+'\'+m.gcPeriod
 IF !fso.FolderExists(m.pDir)
  MESSAGEBOX('ƒ»–≈ “Œ–»ﬂ '+m.pDir+' Õ≈ —”Ÿ≈—“¬”≈“!',0+64,'')
  RETURN 
 ENDIF 

 oMailDir = fso.GetFolder(m.pDir)
 MailDirName = oMailDir.Path
 oFilesInMailDir = oMailDir.Files
 nFilesInMailDir = oFilesInMailDir.Count
 
 IF nFilesInMailDir<=0
  MESSAGEBOX('¬ ƒ»–≈ “Œ–»» '+m.pDir+CHR(13)+CHR(10)+'Õ≈ Œ¡Õ¿–”∆≈ÕŒ Õ» ŒƒÕŒ√Œ ‘¿…À¿!',0+64,'')
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar', 'lpuid')>0
  RETURN .f. 
 ENDIF 

 FOR EACH oFileInMailDir IN oFilesInMailDir

  m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)

  m.BFullName = oFileInMailDir.Path
  m.bname     = oFileInMailDir.Name
  m.recieved  = oFileInMailDir.DateLastModified

  IF LEN(oFileInMailDir.Name)<8
   LOOP 
  ENDIF 
  IF LOWER(oFileInMailDir.Name) != 'me'
   LOOP 
  ENDIF 
  IF UPPER(SUBSTR(oFileInMailDir.Name,3,2)) != m.qcod
   LOOP 
  ENDIF 
  m.lpuid = SUBSTR(oFileInMailDir.Name,5,4)
  IF !SEEK(INT(VAL(m.lpuid)), 'aisoms')
   LOOP
  ENDIF 
  m.mcod = IIF(SEEK(m.lpuid, 'aisoms'), aisoms.mcod, '')
  IF EMPTY(m.mcod)
   LOOP 
  ENDIF 
  m.cmessage = IIF(SEEK(m.lpuid, 'aisoms'), ALLTRIM(aisoms.cmessage), '')
  
  m.mmy = PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 
  ZipFile = 'ot'+m.lpuid+m.qcod+'.'+m.mmy
  IF fso.FileExists(m.pDir+'\'+ZipFile)
   fso.DeleteFile(m.pDir+'\'+ZipFile)
  ENDIF 
 
  ZipOpen(m.pDir+'\'+ZipFile)
  ZipFile(m.BFullName)
  ZipClose()
 
  IF !fso.FileExists(m.pDir+'\'+ZipFile)
   LOOP 
  ENDIF 
  
  m.un_id    = SYS(3)
  m.bansfile = 'b_mee_'  + m.mcod
  m.tansfile = 't_flk_'  + m.mcod
  m.dzipfile  = 'd' + zipfile
  m.mmid     = m.un_id+'.'+m.usrmail+'@'+m.qmail
  m.csubj    = 'OMS#'+m.gcPeriod+'#'+m.lpuid+'##1'

  poi = fso.CreateTextFile(m.pDir + '\' + m.tansfile)

  poi.WriteLine('To: oms@spuemias.msk.oms')
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.csubj)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Resent-Message-Id: '+m.cmessage)
  poi.WriteLine('Attachment: '+m.dzipfile+' '+m.ZipFile)
 
  poi.Close

  fso.CopyFile(m.pDir+'\'+m.ZipFile, pAisOms+'\oms\output\'+m.dzipfile)
  fso.CopyFile(m.pDir+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

  fso.CopyFile(m.pDir+'\'+m.tansfile, m.pDir+'\'+m.bansfile)
  fso.DeleteFile(m.pDir+'\'+m.tansfile)

 ENDFOR
 USE IN aisoms 
 
 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!',0+64,'')

RETURN 
 

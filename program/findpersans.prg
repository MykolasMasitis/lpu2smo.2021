PROCEDURE FindPersAns
 IF MESSAGEBOX('ПОИСКАТЬ ОТВЕТЫ ПО ПЕРСОТЧЕТУ?',4+64,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\persmail.dbf')
  =CreatePersMail()
 ENDIF 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\persmail', 'persmail', 'excl')>0
  IF USED('persmail')
   USE IN persmail
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT persmail
 INDEX on sent_id TAG sent_id 
 SET ORDER TO sent_id
 
 oDir        = fso.GetFolder(pAisoms+'\OMS\INPUT')
 cDirName    = oDir.Path
 oFilesInDir = oDir.Files
 nFilesInDir = oFilesInDir.Count

 FOR EACH oFileInDir IN oFilesInDir
  m.BFullName = oFileInDir.Path
  m.bname     = oFileInDir.Name
  m.recieved  = oFileInDir.DateLastModified
  IF LOWER(LEFT(m.bname,1)) != 'b'
   LOOP 
  ENDIF 
   
  m.cmessage = '' 
  m.resmesid = ''
  m.csubject = ''
  
  m.bodypart   = ''
  m.bparts   = 0 && Сколько присоединенных файлов в одной ИП
  DIMENSION dbparts(10,2)
  dbparts = ''

  CFG = FOPEN(m.BFullName)
  =ReadCFGFile()
  =FCLOSE (CFG)
   
  IF EMPTY(m.resmesid)
   LOOP 
  ENDIF 

  IF SEEK(m.resmesid, 'persmail')
   m.mcod = persmail.mcod
   WAIT m.mcod+'...' WINDOW NOWAIT 

   UPDATE persmail SET rcvd=m.recieved, rcvd_id=m.csubject WHERE sent_id=m.resmesid

   FOR natt = 1 TO m.bparts
    *MESSAGEBOX(pAisOms+'\oms\input\'+dbparts(natt,1),0+64,m.mcod)
    IF fso.FileExists(pAisOms+'\oms\input\'+dbparts(natt,1))
     fso.CopyFile(pAisOms+'\oms\input\'+dbparts(natt,1), pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+dbparts(natt,1), .t.)
     fso.DeleteFile(pAisOms+'\oms\input\'+dbparts(natt,1))
    ENDIF 
   ENDFOR 

   fso.CopyFile(m.BFullName, pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.bname, .t.)
   fso.DeleteFile(m.BFullName)

  ENDIF 

 ENDFOR 

 USE IN persmail 
 
 MESSAGEBOX('OK!', 0+64, '')
 
RETURN 

FUNCTION CreatePersMail
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\persmail.dbf')
  RETURN 
 ENDIF 
 
 CREATE TABLE &pBase\&gcPeriod\PersMail (mcod c(7), lpuid n(4), sent t, sent_id c(75), ;
 	rcvd t, rcvd_id c(75), "flag" c(2))
 	
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT aisoms 
 SCAN 
  m.s_pred = s_pred
  IF m.s_pred<=0
   LOOP 
  ENDIF 
  
  m.mcod  = mcod 
  m.lpuid = lpuid
  
  INSERT INTO persmail FROM MEMVAR 
 ENDSCAN 
 USE IN aisoms 
 USE IN persmail
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\persmail', 'persmail', 'shar')>0
  IF USED('persmail')
   USE IN persmail
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT persmail
 SCAN 
  m.mcod  = mcod 
  m.lpuid = lpuid
  IF EMPTY(sent) OR EMPTY(sent_id)
   IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
    LOOP 
   ENDIF 
   oDir        = fso.GetFolder(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   cDirName    = oDir.Path
   oFilesInDir = oDir.Files
   nFilesInDir = oFilesInDir.Count
   
   m.snt_file   = ''
   m.d_snt_file = {}
   FOR EACH oFileInDir IN oFilesInDir
    m.BFullName = ALLTRIM(oFileInDir.Path)
    m.bname     = ALLTRIM(oFileInDir.Name)
    m.recieved  = oFileInDir.DateLastModified
    
    IF LOWER(m.bname) != 't_y_'
     LOOP 
    ENDIF 
    
    DO CASE 
     CASE EMPTY(m.snt_file)
      m.snt_file   = m.bname 
      m.d_snt_file = oFileInDir.DateLastModified
     OTHERWISE 
      IF oFileInDir.DateLastModified>m.d_snt_file
       m.snt_file   = m.bname 
       m.d_snt_file = oFileInDir.DateLastModified
      ENDIF 
    ENDCASE 

    IF LOWER(m.bname) = 't_y_'
     EXIT 
    ENDIF 

   ENDFOR 
   
   IF EMPTY(m.snt_file)
    LOOP 
   ENDIF 
   
   m.cmessage   = ''
   m.bodypart   = ''
   m.bparts   = 0 && Сколько присоединенных файлов в одной ИП
   DIMENSION dbparts(10,2)
   dbparts = ''
   
   CFG = FOPEN(m.BFullName)
   =ReadCFGFile()
   =FCLOSE (CFG)
   
   REPLACE sent WITH m.d_snt_file, sent_id WITH m.cmessage

  ENDIF 
 ENDSCAN 
 USE IN persmail

RETURN 

FUNCTION ReadCFGFile
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  DO CASE
   CASE UPPER(READCFG) = 'FROM'
    m.cfrom = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'DATE'
    m.cdate = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'CONTENT-TYPE'
    m.ctype = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'MESSAGE'
    m.cmessage = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'RESENT-MESSAGE-ID'
    m.resmesid = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'SUBJECT'
    m.csubject = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    m.csubject1 = LEFT(m.csubject, RAT('#',m.csubject,2))   && Делим subject для последующей вставки кода результата
    m.csubject2 = SUBSTR(m.csubject, RAT('#',m.csubject,1)) && Делим subject для последующей вставки кода результата
   CASE UPPER(READCFG) = 'ATTACHMENT'
    *m.attaches   = m.attaches + 1
    *m.attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    *dattaches(m.attaches,1) = ALLTRIM(SUBSTR(m.attachment, 1, AT(" ",m.attachment)-1)) && Название d-файла
    *dattaches(m.attaches,2) = ALLTRIM(SUBSTR(m.attachment, AT(" ",m.attachment)+1))    && Фактическое название файла
   CASE UPPER(READCFG) = 'BODYPART'
    m.bparts   = m.bparts + 1
    m.bodypart = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dbparts(m.bparts,1) = ALLTRIM(SUBSTR(m.bodypart, 1, AT(" ",m.bodypart)-1))
  ENDCASE
 ENDDO
RETURN 
PROCEDURE FindMeAns
 IF MESSAGEBOX('ПОИСКАТЬ ОТВЕТЫ ПО ME-ФАЙЛАМ?',4+64,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\me_mail.dbf')
  =Creatememail()
 ENDIF 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\me_mail', 'me_mail', 'excl')>0
  IF USED('me_mail')
   USE IN me_mail
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT me_mail
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
  
  IF (AT('#',m.csubject,3)-AT('#',m.csubject,2)-1)<>7
   LOOP 
  ENDIF 

  IF SEEK(m.resmesid, 'me_mail')
   m.mcod = me_mail.mcod
   WAIT m.mcod+'...' WINDOW NOWAIT 

   UPDATE me_mail SET rcvd=m.recieved, rcvd_id=m.csubject WHERE sent_id=m.resmesid

   FOR natt = 1 TO m.bparts
    *MESSAGEBOX(pAisOms+'\oms\input\'+dbparts(natt,1),0+64,m.mcod)
    IF fso.FileExists(pAisOms+'\oms\input\'+dbparts(natt,1))
     fso.CopyFile(pAisOms+'\oms\input\'+dbparts(natt,1), pOut+'\'+m.gcPeriod+'\'+dbparts(natt,1), .t.)
     fso.DeleteFile(pAisOms+'\oms\input\'+dbparts(natt,1))
    ENDIF 
   ENDFOR 

   fso.CopyFile(m.BFullName, pOut+'\'+m.gcPeriod+'\'+m.bname, .t.)
   fso.DeleteFile(m.BFullName)

  ENDIF 

 ENDFOR 

 USE IN me_mail 
 
 MESSAGEBOX('OK!', 0+64, '')
 
RETURN 

FUNCTION Creatememail
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\me_mail.dbf')
  RETURN 
 ENDIF 
 
 CREATE TABLE &pBase\&gcPeriod\me_mail (mcod c(7), lpuid n(4), sent t, sent_id c(75), ;
 	rcvd t, c_rcvd t, rcvd_id c(75), c_id c(75), "flag" c(2))
 USE 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\me_mail', 'me_mail', 'shar')>0
  IF USED('me_mail')
   USE IN me_mail
  ENDIF 
  RETURN 
 ENDIF 
 
 oDir        = fso.GetFolder(pOut+'\'+m.gcPeriod)
 cDirName    = oDir.Path
 oFilesInDir = oDir.Files
 nFilesInDir = oFilesInDir.Count
   
 m.snt_file   = ''
 m.d_snt_file = {}

 FOR EACH oFileInDir IN oFilesInDir
  m.BFullName = ALLTRIM(oFileInDir.Path)
  m.bname     = ALLTRIM(oFileInDir.Name)
  m.recieved  = oFileInDir.DateLastModified
    
  IF LOWER(m.bname) != 't_me_'
   LOOP 
  ENDIF 

  m.mcod       = SUBSTR(m.bname,6,7)
  m.snt_file   = m.bname 
  m.d_snt_file = oFileInDir.DateLastModified

  m.cmessage   = ''
  m.bodypart   = ''
  m.bparts   = 0 && Сколько присоединенных файлов в одной ИП
  DIMENSION dbparts(10,2)
  dbparts = ''
   
  CFG = FOPEN(m.BFullName)
  =ReadCFGFile()
  =FCLOSE (CFG)
 
  INSERT INTO me_mail (mcod, lpuid, sent, sent_id) VALUES (m.mcod,0,m.d_snt_file,m.cmessage)

 ENDFOR 

 USE IN me_mail

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
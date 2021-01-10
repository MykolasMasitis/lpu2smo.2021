PROCEDURE FindMeCtrls
 IF MESSAGEBOX('ПОИСКАТЬ ОТВЕТЫ ПО ME-ФАЙЛАМ?',4+64,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\me_mail.dbf')
  MESSAGEBOX('ОТСТУТСВУЕТ ФАЙЛ me_mail.dbf',0+64,'')
  RETURN 
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
  
  m.attaches = 0
  DIMENSION dattaches(10,2)
  dattaches = ''

  CFG = FOPEN(m.BFullName)
  =ReadCFGFile()
  =FCLOSE (CFG)
   
  IF EMPTY(m.resmesid)
   LOOP 
  ENDIF 
  
  IF (AT('#',m.csubject,3)-AT('#',m.csubject,2)-1)<>4
   LOOP 
  ENDIF 

  IF SEEK(m.resmesid, 'me_mail')
   m.mcod = me_mail.mcod
   WAIT m.mcod+'...' WINDOW NOWAIT 

   UPDATE me_mail SET c_rcvd=m.recieved, c_id=m.csubject WHERE sent_id=m.resmesid

   IF fso.FileExists(pAisOms+'\oms\input\'+dattaches(m.attaches,1))
    fso.CopyFile(pAisOms+'\oms\input\'+dattaches(m.attaches,1), pOut+'\'+m.gcPeriod+'\'+dattaches(m.attaches,2), .t.)
    fso.DeleteFile(pAisOms+'\oms\input\'+dattaches(m.attaches,1))
   ENDIF 

   fso.CopyFile(m.BFullName, pOut+'\'+m.gcPeriod+'\'+m.bname, .t.)
   fso.DeleteFile(m.BFullName)

  ENDIF 

 ENDFOR 

 USE IN me_mail 
 
 MESSAGEBOX('OK!', 0+64, '')
 
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
    m.attaches   = m.attaches + 1
    m.attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dattaches(m.attaches,1) = ALLTRIM(SUBSTR(m.attachment, 1, AT(" ",m.attachment)-1)) && Название d-файла
    dattaches(m.attaches,2) = ALLTRIM(SUBSTR(m.attachment, AT(" ",m.attachment)+1))    && Фактическое название файла
   *CASE UPPER(READCFG) = 'BODYPART'
   * m.bparts   = m.bparts + 1
   * m.bodypart = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   * dbparts(m.bparts,1) = ALLTRIM(SUBSTR(m.bodypart, 1, AT(" ",m.bodypart)-1))
  ENDCASE
 ENDDO
RETURN 
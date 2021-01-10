PROCEDURE CheckMail
PARAMETERS lcUser, loForm

m.IsSilent = .T.

IF !IsAisDir() && Проверка наличия директорий, OMS, INPUT, OUTPUT
 RETURN 
ENDIF 

ArcPumpDir = pAisOms+'\ARC.PUMP' && АИС-дубликаты счетов SOAP
IF !fso.FolderExists(ArcPumpDir)
 fso.CreateFolder(ArcPumpDir)
ENDIF 
ArcRMDir = pAisOms+'\ARC.RM'
IF !fso.FolderExists(ArcRMDir)
 fso.CreateFolder(ArcRMDir)
ENDIF 
ArcTextDir = pAisOms+'\ARC.TEXTONLY'
IF !fso.FolderExists(ArcTextDir)
 fso.CreateFolder(ArcTextDir)
ENDIF 

oMailDir        = fso.GetFolder(pAisOms+'\&lcUser\input')
MailDirName     = oMailDir.Path
oFilesInMailDir = oMailDir.Files
nFilesInMailDir = oFilesInMailDir.Count

IF !m.IsSilent
 MESSAGEBOX('ОБНАРУЖЕНО '+ALLTRIM(STR(nFilesInMailDir))+' ФАЙЛОВ!', 0+64, lcUser)
ENDIF 

IF nFilesInMailDir<=0
 RETURN 
ENDIF 

IF OpenTemplates() != 0
 =CloseTemplates() 
 RETURN 
ENDIF 

WAIT "ПРОСМОТР ПОЧТЫ..." WINDOW NOWAIT 
SELECT AisOms
prvorder = ORDER('aisoms')
SET ORDER TO 

m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
m.un_id = SYS(3)

OldEscStatus = SET("Escape")
SET ESCAPE OFF 
CLEAR TYPEAHEAD 

SET SAFETY OFF 

FOR EACH oFileInMailDir IN oFilesInMailDir

 SCATTER MEMVAR BLANK

 m.BFullName = oFileInMailDir.Path
 m.bname     = oFileInMailDir.Name
 m.recieved  = oFileInMailDir.DateLastModified
 m.lpuid     = 0
 m.processed = DATETIME()
 
 m.cfrom      = ''
 m.cdate      = ''
 m.ctype      = ''
 m.cmessage   = ''
 m.resmesid   = ''
 m.csubject   = ''
 m.csubject1  = ''
 m.csubject2  = ''
 m.attachment = ''
 m.bodypart   = ''

 m.attaches   = 0 && Сколько присоединенных файлов в одной ИП
 DIMENSION dattaches(10,2)
 dattaches = ''

 m.bparts   = 0 && Сколько присоединенных файлов в одной ИП
 DIMENSION dbparts(10,2)
 dbparts = ''

 DO CASE 
 CASE LOWER(oFileInMailDir.Name) = 'b'

 CFG = FOPEN(m.BFullName)
 =ReadCFGFile()
 =FCLOSE (CFG)
 
 IF !m.IsTestMode 
 IF RIGHT(UPPER(ALLTRIM(m.csubject)),2) = 'RM'
  FOR natt = 1 TO m.bparts
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1), ArcRMDir+'\'+dbparts(natt,1), .t.)
    fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
   ENDIF 
  ENDFOR 
  FOR natt = 1 TO m.attaches
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1), ArcRMDir+'\'+dattaches(natt,1), .t.)
    fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
   ENDIF 
  ENDFOR 

  fso.CopyFile(m.BFullName, ArcRMDir+'\'+m.bname, .t.)
  fso.DeleteFile(m.BFullName)

  LOOP
 ELSE 
  * LOOP && Убрать!!!
 ENDIF 
 ENDIF && IF !m.IsTestMode 
 
 IF !m.IsTestMode 
 IF m.ctype = 'text/plain' AND .f. 
  FOR natt = 1 TO m.bparts
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1), ArcTextDir+'\'+dbparts(natt,1), .t.)
    fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
   ENDIF 
  ENDFOR 
  FOR natt = 1 TO m.attaches
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1), ArcTextDir+'\'+dattaches(natt,1), .t.)
    fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
   ENDIF 
  ENDFOR 

  fso.CopyFile(m.BFullName, ArcTextDir+'\'+m.bname, .t.)
  fso.DeleteFile(m.BFullName)

  LOOP
 ELSE 
  *LOOP && Убрать!!!
 ENDIF 
 ENDIF && IF !m.IsTestMode 

 m.sent = dt2date(m.cdate)
   
 m.llIsSubject = .F.

 m.AisAddress = .F.
 m.adresat = PADR(LOWER(SUBSTR(m.cfrom,AT('@',m.cfrom)+1)),27)
 
 IF m.adresat = 'pump.msk.oms' && Временная заглушка!
  ArcDatePumpDir = ArcPumpDir+'\'+DTOC(TTOD(m.sent))
  IF !fso.FolderExists(ArcDatePumpDir)
   fso.CreateFolder(ArcDatePumpDir)
  ENDIF 

  IF !m.IsTestMode 
  FOR natt = 1 TO m.bparts
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1), ArcDatePumpDir+'\'+dbparts(natt,1), .t.)
    IF m.qcod<>'R2'
     fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
    ENDIF 
   ENDIF 
  ENDFOR 
  FOR natt = 1 TO m.attaches
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1), ArcDatePumpDir+'\'+dattaches(natt,1), .t.)
    IF m.qcod<>'R2'
     fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    ENDIF 
   ENDIF 
  ENDFOR 

  fso.CopyFile(m.BFullName, ArcDatePumpDir+'\'+m.bname, .t.)
  IF m.qcod<>'R2'
   fso.DeleteFile(m.BFullName)
  ENDIF 

  ENDIF && IF !m.IsTestMode 

  LOOP
 ELSE 
  *LOOP && Убрать!!!
 ENDIF 
 
 IF LEFT(UPPER(ALLTRIM(m.csubject)),3) != 'OMS'
  LOOP 
 ENDIF 

 IF SUBSTR(m.csubject, AT('#',m.csubject,1)+1, AT('#',m.csubject,2)-(AT('#',m.csubject,1)+1)) != m.gcPeriod
  LOOP 
 ENDIF 
 
 IF RIGHT(UPPER(ALLTRIM(m.csubject)),1) != '1'
  LOOP 
 ENDIF 

 m.lpuid = INT(VAL(SUBSTR(m.csubject, AT('#',m.csubject,2)+1, 4)))
 m.mcod  = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.mcod, "")
 
 WAIT m.mcod WINDOW NOWAIT 

 *IF !INLIST(m.adresat, 'spuemias.msk.oms', 'pump.msk.oms', 'pump.mskt.oms') && Сделано для отправки сообщения "Отправка по АИС запрещена!"
 IF !INLIST(m.adresat, 'spuemias.msk.oms', 'pump.msk.oms') && Сделано для отправки сообщения "Отправка по АИС запрещена!"
  m.AisAddress = .T.
  *m.lpuid   = IIF(SEEK(m.adresat, "sprabo"), sprabo.object_id, m.lpuid)
 ELSE
  *m.mcod  = SUBSTR(ALLTRIM(dattaches(1, 2)),2,7)
  *m.lpuid = IIF(SEEK(m.mcod, "sprlpu", 'mcod'), sprlpu.lpu_id, 0)
 ENDIF 
 
 IF m.lpuid == 0
*  MESSAGEBOX('АДРЕСАТ '+UPPER(ALLTRIM(m.adresat))+' НЕ НАЙДЕН В СПРАВОЧНИКЕ SPRABOXX.DBF!',0+48,lcUser)
  LOOP 
 ENDIF 

 IF EMPTY(m.mcod)
*  MESSAGEBOX('АДРЕСАТ '+UPPER(ALLTRIM(m.adresat))+' НЕ НАЙДЕН В СПРАВОЧНИКЕ SPRLPUXX.DBF!',0+48,STR(m.lpuid,4))
  LOOP 
 ENDIF 

 m.IsIPDouble = IIF(SEEK(m.mcod, 'AisOms', 'mcod') AND !EMPTY(AisOms.cmessage), .t., .f.)
* MESSAGEBOX('m.cmessage= ' + m.cmessage +CHR(13)+CHR(10)+;
 	'AisOms.cmessage= '+AisOms.cmessage)
 IF m.IsIPDouble
  IF ALLTRIM(AisOms.cmessage)==ALLTRIM(m.cmessage)
   LOOP 
  ENDIF 
 ENDIF 
 
 * Посылка ЕМИАС, но ранее была принята SOAP! 
 m.EMIASAFTSOAP = .F.
 IF SEEK(m.mcod, 'AisOms', 'mcod') AND !EMPTY(AisOms.cmessage)
  IF OCCURS('spuemias',ALLTRIM(AisOms.cmessage))<=0 AND OCCURS('spuemias',ALLTRIM(m.cmessage))>0 
   m.EMIASAFTSOAP = .T.
   *LOOP
  ENDIF 
 ENDIF 
 * Посылка ЕМИАС, но ранее была принята SOAP! 
 
 m.cokr    = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.cokr, "")
 m.moname  = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.name, "")
 m.usr     = IIF(SEEK(m.lpuid, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")

 *IF EMPTY(m.usr) AND m.gcUser!='OMS'
 * MESSAGEBOX('ЛПУ '+m.mcod+' НЕ "ПРИВЯЗАНО" К ПОЛЬЗВАТЕЛЮ В USRLPU.DBF!',0+48,lcUser)
 * LOOP 
 *ENDIF 
 
 *IF m.usr != m.gcUser AND m.gcUser!='OMS'
 * MESSAGEBOX('USR' ,0+48,lcUser)
 * MESSAGEBOX('USR' ,0+48,STR(m.lpuid,4))
 * LOOP 
 *ENDIF 

 m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)

 m.previous_id = m.un_id
 m.un_id     = SYS(3)
 DO WHILE m.un_id = m.previous_id
  m.un_id     = SYS(3)
 ENDDO 
 m.previous_id = m.un_id

 m.tansfile = 'tok_' + m.mcod
 m.bansfile = 'bok_' + m.mcod
 m.badfile  = 'bad_' + m.mcod
 iii = 1
 DO WHILE fso.FileExists(pAisOms+'\OMS\OUTPUT\'+m.bansfile)
  m.tansfile = 'tok_' + m.mcod + '_' + PADL(iii,2,'0')
  m.bansfile = 'bok_' + m.mcod + '_' + PADL(iii,2,'0')
  m.badfile  = 'bad_' + m.mcod + '_' + PADL(iii,2,'0')
  iii = iii + 1
 ENDDO 

 m.messageid = ALLTRIM(m.un_id+'.OMS@'+m.qmail)

 && Получено из АИС ОМС
 IF m.AisAddress = .T.
  IF !m.IsTestMode 
   m.csubject = m.csubject1 + '06' +m.csubject2
   poi = fso.CreateTextFile(pAisOms+'\&lcUser\output\'+m.tansfile)
   poi.WriteLine('To: '+m.cfrom)
   poi.WriteLine('Message-Id: ' + m.messageid)
   poi.WriteLine('Content-Type: multipart/mixed')
   poi.WriteLine('Resent-Message-Id: ' + m.cmessage)
   poi.WriteLine('Subject: '+m.csubject)
   poi.WriteLine('Comment: отправка по АИС ОМС запрещена!')
   poi.Close
   
   fso.CopyFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile, pAisOms+'\'+lcUser+'\OUTPUT\'+bansfile)
   fso.DeleteFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile)
  ENDIF 
  
  LOOP 
 ENDIF 
 && Получено из АИС ОМС

 && Получено из ЕМИАС после принятой посылки SOAP
 IF m.EMIASAFTSOAP = .T.
  IF !m.IsTestMode 
   m.csubject = m.csubject1 + '06' +m.csubject2
   poi = fso.CreateTextFile(pAisOms+'\&lcUser\output\'+m.tansfile)
   poi.WriteLine('To: '+m.cfrom)
   poi.WriteLine('Message-Id: ' + m.messageid)
   poi.WriteLine('Content-Type: multipart/mixed')
   poi.WriteLine('Resent-Message-Id: ' + m.cmessage)
   poi.WriteLine('Subject: '+m.csubject)
   poi.WriteLine('Comment: посылка ранее принята по SOAP!')
   poi.Close
   
   fso.CopyFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile, pAisOms+'\'+lcUser+'\OUTPUT\'+bansfile)
   fso.DeleteFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile)
  ENDIF 
  
  LOOP 
 ENDIF 
 && Получено из ЕМИАС после принятой посылки SOAP

 && Присоединено ли что-нибудь к файлу? Если нет, то - в спам!
 *IF m.attaches == 0 AND m.bparts == 0 && Если к файлу ничего не присоединено!
 IF m.attaches == 0  && Если к файлу ничего не присоединено!
  TextToWrite="MyComment: к файлу ничего не присоединено"
  *fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  *fso.DeleteFile(m.BFullName)
  *=WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  =WriteInBFile(m.BFullName, TextToWrite)
  LOOP 
 ENDIF 
 && Присоединено ли что-нибудь к файлу? Если нет, то - в спам!

 && Проверка комплектности посылки
 IsComplect = .T.
 IF m.attaches>0
  FOR natt = 1 TO m.attaches
   IF !fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    IsComplect = .F.
*    MESSAGEBOX('ПРИСОЕДИНЕННЫЙ К ФАЙЛУ '+m.bname+CHR(13)+CHR(10)+;
     ' ATTACHMENT '+dattaches(natt,1)+ ' ОТСУТСТВУЕТ!', 0+48, lcUser)
    LOOP 
   ENDIF 
  ENDFOR 
 ENDIF 

 IF IsComplect = .F.
  FOR natt = 1 TO m.attaches
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    *fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1), SpamDir+'\'+dattaches(natt,1), .t.)
    *fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
   ENDIF 
  ENDFOR 
  TextToWrite="MyComment: отсутствует или недоступен присоединенный файл"
  *fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  *fso.DeleteFile(m.BFullName)
  *=WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  =WriteInBFile(m.BFullName, TextToWrite)
  LOOP 
 ENDIF  

 && Проверка комплектности посылки

 && Есть ли хотя бы один zip-архив?
 llIsOneZip = .T.
 FOR nattach = 1 TO m.attaches

  m.dname   = ALLTRIM(dattaches(nattach, 1)) && так называется файл в посылке, например dOA24R2U.0WK
  m.attname = ALLTRIM(dattaches(nattach, 2)) && так должен называться - b0105012.108
  ffile = fso.GetFile(MailDirName + '\' + m.dname)
  IF ffile.size >= 2
   fhandl = ffile.OpenAsTextStream
   lcHead = fhandl.Read(2)
   fhandl.Close
  ELSE 
   lcHead = ''
  ENDIF 

  IF lcHead == 'PK' && Это zip-файл!
   ZipName = pAisOms+'\'+lcUser+'\input\'+m.dname
   IF !UnzipOpen(ZipName)
    llIsOneZip = .F.
    UnzipClose()
    EXIT 
   ENDIF 
   *rItem   = 'R' + m.qcod + '.' + m.mmy
   *sItem   = 'S' + m.qcod + '.' + m.mmy
   *IF UnzipGotoFileByName(rItem) AND UnzipGotoFileByName(sItem)
    *llIsOneZip = .t.
    *UnzipClose()
    *EXIT 
   *ENDIF 
   UnzipClose()
  ENDIF 

 ENDFOR 
 && Есть ли хотя бы один zip-архив?
 
 IF llIsOneZip == .F.
  TextToWrite="MyComment: среди присоединенных файлов нет ни одного zip-архива"
  *fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  *fso.DeleteFile(m.BFullName)
  *=WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  =WriteInBFile(m.BFullName, TextToWrite)

  IF m.attaches>0
   FOR nattach = 1 TO m.attaches
    m.dname   = ALLTRIM(dattaches(nattach, 1))
    IF !EMPTY(m.dname)
     *fso.CopyFile(MailDirName + '\' + m.dname, SpamDir+'\'+m.dname)
     *fso.DeleteFile(MailDirName + '\' + m.dname)
    ENDIF 
   ENDFOR 
  ENDIF 
  *IF m.bparts > 0
  * FOR npart = 1 TO m.bparts
  *  m.bpname   = ALLTRIM(dbparts(npart, 1))
  *  IF !EMPTY(m.bpname)
  *   *fso.CopyFile(MailDirName + '\' + m.bpname, SpamDir+'\'+m.bpname, .t.)
  *   *fso.DeleteFile(MailDirName + '\' + m.bpname)
  *  ENDIF 
  * ENDFOR 
  *ENDIF 
  LOOP 
 ENDIF 
 && Если нет ни одного zip-архива

 poi = fso.CreateTextFile(pAisOms+'\&lcUser\output\'+m.tansfile)
 poi.WriteLine('To: '+m.cfrom)
 poi.WriteLine('Message-Id: ' + m.messageid)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: ' + m.cmessage)

 && Проверяем комплектность посылки - наличие 5 файлов!
 UnzipOpen(ZipName)

 hItem       = 'H'  + STR(m.lpuid,4) + '.' + m.mmy
 dItem       = 'D'  + STR(m.lpuid,4) + '.' + m.mmy
 nvItem      = 'NV' + STR(m.lpuid,4) + '.' + m.mmy
 nsItem      = 'NS' + STR(m.lpuid,4) + '.' + m.mmy
 rItem       = 'R' + m.qcod + '.' + m.mmy
 sItem       = 'S' + m.qcod + '.' + m.mmy
 hoItem      = 'HO' + m.qcod + '.' + m.mmy
 dsItem      = 'D79S' + m.qcod + '.' + m.mmy
 sprItem     = 'SPR' + STR(m.lpuid,4) + '.' + m.mmy

 * Файлы онкологии
 onkItem  = 'ONK_SL' + m.qcod + '.' + m.mmy && "старый" файл

 SLItem    = 'ONK_SL' + m.qcod + '.' + m.mmy
 USLItem   = 'ONK_USL' + m.qcod + '.' + m.mmy
 CONSItem  = 'ONK_CONS' + m.qcod + '.' + m.mmy
 LSItem    = 'ONK_LS' + m.qcod + '.' + m.mmy
 NAPRItem  = 'ONK_NAPR_V_OUT' + m.qcod + '.' + m.mmy
 DIAGItem  = 'ONK_DIAG' + m.qcod + '.' + m.mmy
 PROTItem  = 'ONK_PROT' + m.qcod + '.' + m.mmy
 * Файлы онкологии

 IF !IsIPComplete()
  LOOP 
 ENDIF 
 
 UnzipClose()
 && Проверяем комплектность посылки - наличие 5 файлов!

 && Если посылка повторная
* IsThisIPDouble = IIF(SEEK(m.mcod, 'AisOms', 'mcod'), .t., .f.)
 IsThisIPDouble = IIF(SEEK(m.mcod, 'AisOms', 'mcod') AND !EMPTY(AisOms.Sent), .t., .f.)
* IsThisIPDouble = IIF(SEEK(m.mcod, 'AisOms', 'mcod') AND !EMPTY(AisOms.Processed), .t., .f.)
 IF IsThisIPDouble 
  frst_time = AisOms.Sent
*  frst_time = AisOms.Processed
  IF frst_time > m.sent && Принятая ранее посылка отправлена позже обнаруженной!
*  IF frst_time > m.processed && Принятая ранее посылка отправлена позже обнаруженной! этот код бессмыслен и недостижим! processed=datetime()
   m.csubject = m.csubject1 + '99' +m.csubject2
   m.cerrmessage = [Уже загружена более поздняя посылка]
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pAisOms+'\&lcUser\OUTPUT\'+m.badfile)
   fso.DeleteFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile)
   *DoubleDir = pDouble + '\' + m.mcod
   *IF !fso.FolderExists(DoubleDir)
   * fso.CreateFolder(DoubleDir)
   *ENDIF 

   *fso.CopyFile(m.BFullName, DoubleDir+'\'+m.bname)
   *fso.DeleteFile(m.BFullName)
   *IF m.attaches>0
   * FOR nattach = 1 TO m.attaches
   *  m.dname   = ALLTRIM(dattaches(nattach, 1))
   *  IF !EMPTY(m.dname)
   *   fso.CopyFile(MailDirName + '\' + m.dname, pDouble+'\'+m.mcod+'\'+m.dname)
   *   fso.DeleteFile(MailDirName + '\' + m.dname)
   *  ENDIF 
   * ENDFOR 
   *ENDIF 
   *IF m.bparts > 0
   * FOR npart = 1 TO m.bparts
   *  m.bpname   = ALLTRIM(dbparts(npart, 1))
   *  IF !EMPTY(m.bpname)
   *   fso.CopyFile(MailDirName + '\' + m.bpname, pDouble+'\'+m.mcod+'\'+m.dname)
   *   fso.DeleteFile(MailDirName + '\' + m.bpname)
   *  ENDIF 
   * ENDFOR 
   *ENDIF 

   *IF NOT SEEK(m.cmessage, "daisoms")
   * INSERT INTO daisoms FROM MEMVAR 
   *ENDIF

   LOOP 

  ELSE                  && Обнаруженная посылка более свежая, чем принятая ранее!

   m.prv_bfile = ALLTRIM(AisOms.bname)
   *DoubleDir   = pDouble + '\' + m.mcod
   *IF !fso.FolderExists(DoubleDir)
   * fso.CreateFolder(DoubleDir)
   *ENDIF 
   
   CFG = FOPEN(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.prv_bfile)
   DO WHILE NOT FEOF(CFG)
    READCFG = FGETS (CFG)
    IF UPPER(READCFG) = 'ATTACHMENT'
     m.dbl_attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
     m.dbl_dname      = ALLTRIM(SUBSTR(m.dbl_attachment, 1, AT(" ",m.dbl_attachment)-1)) && Название d-файла
     m.dbl_attname    = ALLTRIM(SUBSTR(m.dbl_attachment, AT(" ",m.dbl_attachment)+1))    && Фактическое название файла
     IF !EMPTY(m.dbl_attname)
      fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dbl_attname, m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dbl_dname)
      fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dbl_attname)
     ENDIF 
    ENDIF 
   ENDDO
   = FCLOSE (CFG)
   
   *MESSAGEBOX(m.dbl_dname,0+64,m.mcod)
   *fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.prv_bfile, DoubleDir+'\'+m.prv_bfile, .t.)
   
   *fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\*.*') && вот это зачем?!

   ** Удаляем файлы-признаки отправки ФЛК и МЭК
   m.b_flk = pbase+'\'+m.gcperiod+'\'+m.mcod+'\b_flk_'+mcod
   IF fso.FileExists(m.b_flk)
    fso.DeleteFile(m.b_flk)
   ENDIF 
   m.b_mek = pbase+'\'+m.gcperiod+'\'+m.mcod+'\b_mek_'+mcod
   IF fso.FileExists(m.b_mek)
    fso.DeleteFile(m.b_mek)
   ENDIF 

   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\expselected.dbf')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\expselected.dbf')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\expselected.cdx')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\expselected.cdx')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.zip')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.zip')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.xml')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.xml')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.http')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\answer.http')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\polltag.xml')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\polltag.xml')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\polltag.http')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\polltag.http')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\soapans.dbf')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\soapans.dbf')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\request.http')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\request.http')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\request.xml')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\request.xml')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\data.xml')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\data.xml')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ctrl'+m.qcod+'.dbf')
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ctrl'+m.qcod+'.dbf')
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\t_y_'+m.mcod)
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\t_y_'+m.mcod)
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+STR(m.lpuid,4)+m.qcod+'.'+m.mmy)
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+STR(m.lpuid,4)+m.qcod+'.'+m.mmy)
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ot'+STR(m.lpuid,4)+m.qcod+'.'+m.mmy)
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\ot'+STR(m.lpuid,4)+m.qcod+'.'+m.mmy)
   ENDIF 
   IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\d'+m.qcod+STR(m.lpuid,4)+'.'+m.mmy)
    fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\d'+m.qcod+STR(m.lpuid,4)+'.'+m.mmy)
   ENDIF 
   ** Удаляем файлы-признаки отправки ФЛК и МЭК
   
   ** Удаляем все файлы: протокол, акт, реестр актов, табличную форму актов
   m.l_path = pbase+'\'+m.gcperiod+'\'+m.mcod
   m.mmy    = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
   
   DIMENSION dim_files(5)
   dim_files(1) = "Pr"+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
   dim_files(2) = "Mk" + STR(m.lpuid,4) + m.qcod + m.mmy
   dim_files(3) = "Mt" + STR(m.lpuid,4) + m.qcod + m.mmy
   dim_files(4) = "Mc" + STR(m.lpuid,4) + m.qcod + m.mmy
   dim_files(5) = 'pdf'+m.qcod+m.mmy
   
   FOR i=1 TO ALEN(dim_files,1)
    IF fso.FileExists(m.l_path+'\'+ALLTRIM(dim_files(i))+'.xls')
     fso.DeleteFile(m.l_path+'\'+ALLTRIM(dim_files(i))+'.xls')
    ENDIF 
    IF fso.FileExists(m.l_path+'\'+ALLTRIM(dim_files(i))+'.pdf')
     fso.DeleteFile(m.l_path+'\'+ALLTRIM(dim_files(i))+'.pdf')
    ENDIF 
   ENDFOR 
   
   RELEASE dim_files, l_path
   ** Удаляем все файлы: протокол, акт, реестр актов, табличную форму актов

   m.t_BName      = AisOms.BName
   m.t_Sent       = AisOms.Sent
   m.t_Recieved   = AisOms.Recieved
   m.t_Processed  = AisOms.Processed
   m.t_CMessage   = AisOms.CMessage
   m.t_Paz        = AisOms.Paz
   m.t_nsch       = AisOms.nsch
   m.t_s_pred     = AisOms.s_pred
   m.t_dname      = AisOms.dname
*   DELETE IN AisOms && !!!
   
   loForm.get_recs.value  = loForm.get_recs.value - 1
   loForm.get_paz.value   = loForm.get_paz.value - m.t_paz
   loForm.get_nsch.value  = loForm.get_nsch.value - m.t_nsch
   loForm.get_sum.value   = loForm.get_sum.value - m.t_s_pred

   *IF NOT SEEK(m.t_CMessage, "daisoms")
   * INSERT INTO daisoms (LpuId,Mcod,BName,Sent,Recieved,Processed,CFrom,CMessage,;
   *  Paz,s_pred,dname ) ;
   *  VALUES ;
   *  (m.lpuid,m.mcod,m.t_BName,m.t_Sent,m.t_Recieved,m.t_Processed,m.cfrom,m.t_CMessage,;
   *   m.t_Paz, m.t_s_pred,m.t_dname)
   *ENDIF

   RELEASE m.t_BName,m.t_Sent,m.t_Recieved,m.t_Processed,m.t_CMessage,;
    m.t_Paz,m.t_nsch, m.t_s_pred,m.t_dname

  ENDIF 
 ENDIF 
 && Если посылка повторная
 
 * С этого места начинается реальная обработка посылки!
 
 m.t_0 = SECONDS()

 && Если это нормальная и новая посылка!
 InDirPeriod = pBase + '\' + m.gcPeriod
 IF !fso.FolderExists(InDirPeriod)
  fso.CreateFolder(InDirPeriod)
 ENDIF 
 InDir = pBase + '\' + m.gcPeriod + '\' + m.mcod
 IF !fso.FolderExists(InDir)
  fso.CreateFolder(InDir)
 ENDIF 

 *IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)

 * FOR tParam = 1 TO 999
 *  m.fname000 = STRTRAN(m.fname, m.mmy, PADL(tParam,3,'0'))
 *  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname000)
 *   EXIT 
 *  ENDIF 
 * ENDFOR 
   
 * fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname, ;
 * 	m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname000)
 * fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.fname)

 *ENDIF 

 fso.CopyFile(m.BFullName, InDir + '\' + m.bname)
 fso.DeleteFile(m.BFullName)
 IF m.attaches>0
  FOR nattach = 1 TO m.attaches
   m.ddname   = ALLTRIM(dattaches(nattach, 1))
   m.aattname = ALLTRIM(dattaches(nattach, 2))
   IF !EMPTY(m.dname)
    fso.CopyFile(MailDirName + '\' + m.ddname, InDir+'\'+m.aattname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.ddname)
   ENDIF 
  ENDFOR 
 ENDIF 
* IF m.bparts > 0
*  FOR npart = 1 TO m.bparts
*   m.bpname   = ALLTRIM(dbparts(npart, 1))
*   IF !EMPTY(m.bpname)
*    fso.CopyFile(MailDirName + '\' + m.bpname, InDir+'\'+m.bpname, .t.)
*    fso.DeleteFile(MailDirName + '\' + m.bpname)
*   ENDIF 
*  ENDFOR 
* ENDIF 

 ZipName = InDir + '\' + m.attname
 ZipDir  = InDir + '\'

 UnzipOpen(ZipName)
 UnZipSetFolder(m.pbase+'\'+m.gcperiod+'\'+m.mcod)
 UnZip()

 *UnzipGotoFileByName(rItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(sItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(dItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(nvItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(sprItem)
 *UnzipFile(ZipDir)
 *UnzipGotoFileByName(HOItem)
 *UnzipFile(ZipDir)
 *IF UnzipGotoFileByName(OnkItem)
 * UnzipFile(ZipDir)
 *ENDIF 
 *IF UnzipGotoFileByName(onkDiagItem)
 * UnzipFile(ZipDir)
 *ENDIF 
 *IF UnzipGotoFileByName(onkProtItem)
 * UnzipFile(ZipDir)
 *ENDIF 

 UnzipClose()

 m.lcCurDir = pBase + '\' + m.gcPeriod + '\' + m.mcod+'\'
 SET DEFAULT TO (lcCurDir)

 IF OpenFile("&dItem",  "dfile",  "SHARED")>0
  IF USED('dfile')
   USE IN dfile
  ENDIF 
  TextToWrite="MyComment: в посылке не dbf файлы!"
  =WriteInBFile(m.BFullName, TextToWrite)
  fso.CopyFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile, pAisOms+'\'+lcUser+'\OUTPUT\'+bansfile)
  fso.DeleteFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile)
  LOOP 
 ENDIF 
 IF OpenFile("&nvItem", "nvfile", "SHARED")>0
  USE IN dfile
  IF USED('nvfile')
   USE IN nvfile
  ENDIF 
  TextToWrite="MyComment: в посылке не dbf файлы!"
  =WriteInBFile(m.BFullName, TextToWrite)
  LOOP 
 ENDIF 
 IF OpenFile("&rItem",  "rfile",  "SHARED")>0
  USE IN dfile
  USE IN nvfile
  IF USED('rfile')
   USE IN rfile
  ENDIF 
  TextToWrite="MyComment: в посылке не dbf файлы!"
  =WriteInBFile(m.BFullName, TextToWrite)
  LOOP 
 ENDIF 
 IF OpenFile("&sItem",  "sfile",  "SHARED")>0
  USE IN dfile
  USE IN nvfile
  USE IN rfile
  IF USED('sfile')
   USE IN sfile
  ENDIF 
  TextToWrite="MyComment: в посылке не dbf файлы!"
  =WriteInBFile(m.BFullName, TextToWrite)
  LOOP 
 ENDIF 
 IF OpenFile("&sprItem", "sprfile", "SHARED")>0
  USE IN dfile
  USE IN nvfile
  USE IN rfile
  USE IN sfile
  IF USED('sprfile')
   USE IN sprfile
  ENDIF 
  TextToWrite="MyComment: в посылке не dbf файлы!"
  =WriteInBFile(m.BFullName, TextToWrite)
  LOOP 
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+HOitem)
  IF OpenFile("&HOitem", "hofile", "SHARED")>0
   USE IN dfile
   USE IN nvfile
   USE IN rfile
   USE IN sfile
   USE IN sprfile
   IF USED('hofile')
    USE IN hofile
   ENDIF 
   TextToWrite="MyComment: в посылке не dbf файлы!"
   =WriteInBFile(m.BFullName, TextToWrite)
   LOOP 
  ENDIF 
 ENDIF 
 *IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+OnkItem)
 * =OpenFile("&OnkItem", "onkfile", "SHARED")
 *ENDIF 
 
 IF !CheckFilesStucture()
  LOOP 
 ENDIF 

 m.csubject = m.csubject1 + '01' +m.csubject2
 poi.WriteLine('Subject: '+m.csubject)
 poi.WriteLine('BodyPart: OK' )
 poi.Close

 =CloseItems()

 fso.CopyFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile, pAisOms+'\'+lcUser+'\OUTPUT\'+bansfile)
 fso.DeleteFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile)
 
 lcDir  = m.pBase + '\' + m.gcPeriod + '\' + m.mcod
 People = lcDir + '\people'
 Talon  = lcDir + '\talon'
 Otdel  = lcDir + '\otdel'
 Doctor = lcDir + '\doctor'
 Error  = lcDir + '\e' + m.mcod
 mError = lcDir + '\m' + m.mcod

 IF fso.FileExists(lcDir + '\people.dbf')
  fso.DeleteFile(lcDir + '\people.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\people.cdx')
  fso.DeleteFile(lcDir + '\people.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\talon.dbf')
  fso.DeleteFile(lcDir + '\talon.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\talon.cdx')
  fso.DeleteFile(lcDir + '\talon.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\otdel.dbf')
  fso.DeleteFile(lcDir + '\otdel.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\otdel.cdx')
  fso.DeleteFile(lcDir + '\otdel.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\doctor.dbf')
  fso.DeleteFile(lcDir + '\doctor.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\doctor.cdx')
  fso.DeleteFile(lcDir + '\doctor.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\e'+m.mcod+'.dbf')
  fso.DeleteFile(lcDir + '\e'+m.mcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\e'+m.mcod+'.dbf')
  fso.DeleteFile(lcDir + '\e'+m.mcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\m'+m.mcod+'.dbf')
  fso.DeleteFile(lcDir + '\m'+m.mcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\m'+m.mcod+'.dbf')
  fso.DeleteFile(lcDir + '\m'+m.mcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_sl'+m.qcod+'.dbf')
  fso.DeleteFile(lcDir + '\onk_sl'+m.qcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_sl'+m.qcod+'.cdx')
  fso.DeleteFile(lcDir + '\onk_sl'+m.qcod+'.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_ls'+m.qcod+'.dbf')
  fso.DeleteFile(lcDir + '\onk_ls'+m.qcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_ls'+m.qcod+'.cdx')
  fso.DeleteFile(lcDir + '\onk_ls'+m.qcod+'.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_diag'+m.qcod+'.dbf')
  fso.DeleteFile(lcDir + '\onk_diag'+m.qcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_diag'+m.qcod+'.cdx')
  fso.DeleteFile(lcDir + '\onk_diag'+m.qcod+'.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_cons'+m.qcod+'.dbf')
  fso.DeleteFile(lcDir + '\onk_cons'+m.qcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_cons'+m.qcod+'.cdx')
  fso.DeleteFile(lcDir + '\onk_cons'+m.qcod+'.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_usl'+m.qcod+'.dbf')
  fso.DeleteFile(lcDir + '\onk_usl'+m.qcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_usl'+m.qcod+'.cdx')
  fso.DeleteFile(lcDir + '\onk_usl'+m.qcod+'.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_prot'+m.qcod+'.dbf')
  fso.DeleteFile(lcDir + '\onk_prot'+m.qcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_prot'+m.qcod+'.cdx')
  fso.DeleteFile(lcDir + '\onk_prot'+m.qcod+'.cdx')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_napr_v_out'+m.qcod+'.dbf')
  fso.DeleteFile(lcDir + '\onk_napr_v_out'+m.qcod+'.dbf')
 ENDIF 
 IF fso.FileExists(lcDir + '\onk_napr_v_out'+m.qcod+'.cdx')
  fso.DeleteFile(lcDir + '\onk_napr_v_out'+m.qcod+'.cdx')
 ENDIF 

 =CreateFilesStructure()
 =OpenLocalFiles()
 
 m.s_pred  = 0
 m.nsch    = 0
 m.krank   = 0
 m.paz_dst = 0
 m.paz_st  = 0
 m.paz_vmp = 0
 m.s_lek   = 0
 
 =MakePeople()
 m.t_1 = SECONDS()
 =MakeTalon()
 m.t_2 = SECONDS()
 =MakeOtdel() 
 m.t_3 = SECONDS()
 =MakeDoctor()
 m.t_4 = SECONDS()
 =MakeHO()
 m.t_5 = SECONDS()

 =OpenFile(lcDir + '\talon', 'talon', 'shar', 'recid_lpu')

 =MakeOnkFile(SLItem)
 IF fso.FileExists(STRTRAN(SLItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(SLItem, m.mmy, 'dbf'), 'onk_sl', 'excl')=0
   SELECT onk_sl
   SET SAFETY OFF
   ALTER table onk_sl ADD COLUMN rid i
   ALTER table onk_sl ADD COLUMN sqlid i
   ALTER table onk_sl ADD COLUMN sqldt t
   IF USED('talon')
    SET RELATION TO recid_s INTO talon 
    REPLACE ALL rid WITH talon.recid 
    SET RELATION OFF INTO talon 
   ENDIF 
   INDEX on rid TAG rid 
   INDEX on recid_s TAG recid_s
   INDEX on recid TAG recid
   SET ORDER TO recid 
   SET SAFETY ON 
   *USE IN onk_sl
  ELSE 
   IF USED('onk_sl')
    USE IN onk_sl
   ENDIF 
  ENDIF 
 ENDIF 
 =MakeOnkFile(USLItem)
 IF fso.FileExists(STRTRAN(USLItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(USLItem, m.mmy, 'dbf'), 'onk_usl', 'excl')=0
   SELECT onk_usl
   SET SAFETY OFF
   ALTER table onk_usl ADD COLUMN rid i
   ALTER table onk_usl ADD COLUMN sqlid i
   ALTER table onk_usl ADD COLUMN sqldt t
   IF USED('onk_sl')
    SET RELATION TO recid_sl INTO onk_sl
    REPLACE ALL rid WITH onk_sl.rid 
    SET RELATION OFF INTO onk_sl
   ENDIF 
   INDEX on recid TAG recid
   INDEX on recid_sl TAG recid_s
   SET ORDER TO recid 
   SET SAFETY ON 
   *USE IN onk_usl
  ELSE 
   IF USED('onk_usl')
    USE IN onk_usl
   ENDIF 
  ENDIF 
 ENDIF 

 =MakeOnkFile(LSItem)
 IF fso.FileExists(STRTRAN(LSItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(LSItem, m.mmy, 'dbf'), 'onk_ls', 'excl')=0
   SELECT onk_ls
   SET SAFETY OFF
   ALTER table onk_ls ADD COLUMN rid i
   ALTER table onk_ls ADD COLUMN sqlid i
   ALTER table onk_ls ADD COLUMN sqldt t
   IF USED('onk_usl')
    SET RELATION TO recid_usl INTO onk_usl
    REPLACE ALL rid WITH onk_usl.rid 
    SET RELATION OFF INTO onk_usl
   ENDIF 
   ALTER TABLE onk_ls ADD COLUMN s_all n(11,2)
   ALTER TABLE onk_ls ADD COLUMN oms l
   INDEX on recid_usl TAG recid_s

   CREATE CURSOR sss (recid i, s_all n(11,2))
   SELECT sss
   INDEX on recid TAG recid
   SET ORDER TO recid
   
   SELECT onk_ls
   SCAN 
    m.tip_opl = IIF(FIELD('tip_opl')=UPPER('tip_opl'), tip_opl, 1)
    IF m.tip_opl!=1
     LOOP 
    ENDIF 
    m.cod = cod
    IF !INLIST(m.cod, 97158, 81094)
     LOOP 
    ENDIF 
    m.date_inj = date_inj
    IF !BETWEEN(m.date_inj, {01.03.2019}, m.tdat2)
     LOOP 
    ENDIF 
    
    m.recid = rid   
    m.ds_c = IIF(USED('talon') AND SEEK(m.recid, 'talon', 'recid'), talon.ds, '')
    IF EMPTY(m.ds_c)
     LOOP 
    ENDIF 
    IF !SEEK(m.ds_c, 'mkb_c')
     LOOP 
    ENDIF 

    m.r_up   = ALLTRIM(r_up) && розничая упаковка
    IF EMPTY(m.r_up)
     LOOP 
    ENDIF 

    m.dd_sid = sid
    m.dt_d   = dt_d && курсовая (дневная) доза в единицах назначения!

    m.s_all = FLS(m.dd_sid, m.dt_d, m.r_up)
    IF m.s_all<=0
     LOOP 
    ENDIF 
    
    IF m.s_all>0
     IF !SEEK(m.recid, 'sss')
      INSERT INTO sss FROM MEMVAR 
     ELSE 
      m.o_s_all = sss.s_all
      m.n_s_all = m.o_s_all + m.s_all
      UPDATE sss SET s_all = m.n_s_all WHERE recid=m.recid
     ENDIF 
    ENDIF 

    REPLACE s_all WITH m.s_all

   ENDSCAN 
   
   SELECT Talon
   IF USED('sss')
    SET RELATION TO recid INTO sss
    REPLACE ALL s_lek WITH sss.s_all
    SET RELATION OFF INTO sss
    USE IN sss 
   ENDIF 
   SUM s_lek TO m.s_lek

   SET SAFETY ON 
   USE IN onk_ls
  ELSE 
   IF USED('onk_ls')
    USE IN onk_ls
   ENDIF 
  ENDIF 
 ENDIF 


 =MakeOnkFile(CONSItem)
 IF fso.FileExists(STRTRAN(CONSItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(CONSItem, m.mmy, 'dbf'), 'onk_cons', 'excl')=0
   SELECT onk_cons
   SET SAFETY OFF
   ALTER table onk_cons ADD COLUMN rid i
   ALTER table onk_cons ADD COLUMN sqlid i
   ALTER table onk_cons ADD COLUMN sqldt t
   IF USED('talon')
    SET RELATION TO recid_s INTO talon
    REPLACE ALL rid WITH talon.recid
    SET RELATION OFF INTO talon
   ENDIF 
   INDEX on recid_s TAG recid
   SET SAFETY ON 
   USE IN onk_cons
  ELSE 
   IF USED('onk_cons')
    USE IN onk_cons
   ENDIF 
  ENDIF 
 ENDIF 

 =MakeOnkFile(NAPRItem)
 IF fso.FileExists(STRTRAN(NAPRItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(NAPRItem, m.mmy, 'dbf'), 'onk_napr', 'excl')=0
   SELECT onk_napr
   SET SAFETY OFF
   ALTER table onk_napr ADD COLUMN rid i
   ALTER table onk_napr ADD COLUMN sqlid i
   ALTER table onk_napr ADD COLUMN sqldt t
   IF USED('talon')
    SET RELATION TO recid_s INTO talon
    REPLACE ALL rid WITH talon.recid
    SET RELATION OFF INTO talon
   ENDIF 
   INDEX on recid_s TAG recid
   SET SAFETY ON
   USE IN onk_napr
  ELSE 
   IF USED('onk_napr')
    USE IN onk_napr
   ENDIF 
  ENDIF 
 ENDIF 
 
 =MakeOnkFile(DIAGItem)
 IF fso.FileExists(STRTRAN(DIAGItem, m.mmy, 'dbf'))
  IF OpenFile(STRTRAN(DIAGItem, m.mmy, 'dbf'), 'onk_diag', 'excl')=0
   SELECT onk_diag
   SET SAFETY OFF
   ALTER table onk_diag ADD COLUMN rid i
   ALTER table onk_diag ADD COLUMN sqlid i
   ALTER table onk_diag ADD COLUMN sqldt t
   IF USED('onk_sl')
    SET RELATION TO recid_sl INTO onk_sl
    REPLACE ALL rid WITH onk_sl.rid
    SET RELATION OFF INTO onk_sl
   ENDIF 
   INDEX on recid_sl TAG recid
   SET SAFETY ON 
   USE IN onk_diag
  ELSE 
   IF USED('onk_diag')
    USE IN onk_diag
   ENDIF 
  ENDIF 
 ENDIF 

 =MakeOnkFile(PROTItem)

 USE IN talon 
 IF USED('onk_sl')
  USE IN onk_sl
 ENDIF 
 IF USED('onk_usl')
  USE IN onk_usl
 ENDIF 
 
 IF fso.FileExists(STRTRAN(SLItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(SLItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(USLItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(USLItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(LSItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(LSItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(DIAGItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(DIAGItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(CONSItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(CONSItem, m.mmy, 'bak'))
 ENDIF 
 IF fso.FileExists(STRTRAN(NAPRItem, m.mmy, 'bak'))
  fso.DeleteFile(STRTRAN(NAPRItem, m.mmy, 'bak'))
 ENDIF 

 m.t_6 = SECONDS()

 loForm.get_recs.value = loForm.get_recs.value + 1
 loForm.get_paz.value  = loForm.get_paz.value + m.paz
 loForm.get_nsch.value = loForm.get_nsch.value + m.nsch
 loForm.get_sum.value  = loForm.get_sum.value + m.s_pred
 
 UPDATE aisoms SET bname=m.bname, dname=m.dname, sent=m.sent, recieved=m.recieved, ;
 	processed=m.processed, cfrom=m.cfrom, cmessage=m.cmessage, paz=m.paz, nsch=m.nsch, ;
 	s_pred=m.s_pred, s_lek=m.s_lek, krank=m.krank, paz_dst=m.paz_dst, paz_st=m.paz_st, paz_vmp=m.paz_vmp,; 
 	erz_id='', erz_status=0, sum_flk=0, ls_flk=0, ispr=.f., t_1=m.t_6-m.t_0,;
 	polltag='', polltagdt={}, soapstatus='' WHERE mcod=m.mcod 
 	
 
 
 *IF m.qcod!='I3' 
 m.t_a = SECONDS()
 WAIT "Отправка запроса "+m.mcod+"..." WINDOW NOWAIT 
 loForm.erzsend(m.mcod)
 WAIT CLEAR 
 m.t_b = SECONDS()
 
 UPDATE aisoms SET t_2=m.t_b-m.t_a WHERE mcod=m.mcod
 *ENDIF 

 SELECT AisOms
 
 WAIT CLEAR 
 loForm.Refresh 

 CASE LOWER(oFileInMailDir.Name) = 'r' && Если это r-файл

 CFG = FOPEN(m.BFullName)
 =ReadCFGFile()
 = FCLOSE (CFG)

 WAIT m.cfrom WINDOW NOWAIT 
   
 fso.CopyFile(m.BFullName, DaemonDir + '\' + m.bname, .t.)
 fso.DeleteFile(m.BFullName)

 IF m.attaches > 0
  FOR nattach = 1 TO m.attaches
   m.ddname   = ALLTRIM(dattaches(nattach, 1))
   m.aattname = ALLTRIM(dattaches(nattach, 2))
   IF !EMPTY(m.dname) AND fso.FileExists(MailDirName + '\' + m.ddname)
    fso.CopyFile(MailDirName + '\' + m.ddname, DaemonDir+'\'+m.aattname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.ddname)
   ENDIF 
  ENDFOR 
 ENDIF 

 *IF m.bparts > 0
 * FOR npart = 1 TO m.bparts
 *  m.bpname   = ALLTRIM(dbparts(npart, 1))
 *  IF !EMPTY(m.bpname) AND fso.FileExists(MailDirName + '\' + m.bpname)
 *   fso.CopyFile(MailDirName + '\' + m.bpname, DaemonDir+'\'+m.bpname, .t.)
 *   fso.DeleteFile(MailDirName + '\' + m.bpname)
 *  ENDIF 
 * ENDFOR 
 *ENDIF 

 SELECT AisOms
 
 WAIT CLEAR 
 loForm.Refresh 

 ENDCASE 

 IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 

*MESSAGEBOX('MakePeople: '+trans(m.t_1 - m.t_0, '99999.999')+CHR(13)+CHR(10)+;
 	'MakeTalon: '+trans(m.t_2 - m.t_1, '99999.999')+CHR(13)+CHR(10)+;
 	'MakeOtdel: '+trans(m.t_3 - m.t_2, '99999.999')+CHR(13)+CHR(10)+;
 	'MakeDoctor: '+trans(m.t_4 - m.t_3, '99999.999')+CHR(13)+CHR(10)+;
 	'MakeHO: '+trans(m.t_5 - m.t_4, '99999.999')+CHR(13)+CHR(10)+;
 	'MakeOnk: '+trans(m.t_6 - m.t_5, '99999.999')+CHR(13)+CHR(10);
 	,0+64,m.mcod)

NEXT && Цикл по файлам

SET ESCAPE &OldEscStatus

=CloseTemplates() 

SET ORDER TO (prvorder)
loForm.Refresh
loForm.LockScreen=.f.

WAIT CLEAR 

nFilesInMailDir = oFilesInMailDir.Count
IF USED('lpuias')
 USE IN lpuais
ENDIF 

SET SAFETY ON 

IF !m.IsSilent
 MESSAGEBOX('ОСТАЛОСЬ '+ALLTRIM(STR(nFilesInMailDir))+' НЕОБРАБОТАННЫХ ИП!', 0+64, lcUser)
ENDIF 

RETURN 

FUNCTION CopyToTrash(lcPath, nTip)
 *fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pAisOms+'\&lcUser\OUTPUT\'+m.badfile)
 *fso.DeleteFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile)
 *TrashDir = pTrash + '\' + m.mcod
 *IF !fso.FolderExists(TrashDir)
 * fso.CreateFolder(TrashDir)
 *ENDIF 
 *fso.CopyFile(lcPath + '\' + m.bname, TrashDir+'\'+m.bname)
 *fso.DeleteFile(lcPath + '\' + m.bname)

 *FOR nattach = 1 TO m.attaches
 * IF !EMPTY(ALLTRIM(dattaches(nattach, 1)))
 *  fso.CopyFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)), TrashDir + '\' + ALLTRIM(dattaches(nattach, 2)))
 *  fso.DeleteFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)))
 * ENDIF 
 *ENDFOR 

 *IF NOT SEEK(m.cmessage, "taisoms", "cmessage")
 * INSERT INTO taisoms FROM MEMVAR 
 *ENDIF
RETURN 

FUNCTION ClDir
 IF fso.FileExists(dItem)
  DELETE FILE &dItem
 ENDIF 
 IF fso.FileExists(hItem)
  DELETE FILE &hItem
 ENDIF 
 IF fso.FileExists(nvItem)
  DELETE FILE &nvItem
 ENDIF 
 IF fso.FileExists(nsItem)
  DELETE FILE &nsItem
 ENDIF 
 IF fso.FileExists(rItem)
  DELETE FILE &ritem
 ENDIF 
 IF fso.FileExists(sItem)
  DELETE FILE &sItem
 ENDIF 
 IF fso.FileExists(dsItem)
  DELETE FILE &dsItem
 ENDIF 
 IF fso.FileExists(sprItem)
  DELETE FILE &sprItem
 ENDIF 
 IF fso.FileExists(hoItem)
  DELETE FILE &hoItem
 ENDIF 
 IF fso.FileExists(SLItem)
  DELETE FILE &SLItem
 ENDIF 
 IF fso.FileExists(USLItem)
  DELETE FILE &USLItem
 ENDIF 
 IF fso.FileExists(CONSItem)
  DELETE FILE &CONSItem
 ENDIF 
 IF fso.FileExists(LSItem)
  DELETE FILE &LSItem
 ENDIF 
 IF fso.FileExists(NAPRItem)
  DELETE FILE &NAPRItem
 ENDIF 
 IF fso.FileExists(DIAGItem)
  DELETE FILE &DIAGItem
 ENDIF 
 IF fso.FileExists(PROTItem)
  DELETE FILE &PROTItem
 ENDIF 
RETURN 

FUNCTION OpenTemplates
 tn_result = 0
 tn_result = tn_result + OpenFile("&ptempl\dxxxx.mmy", "d_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\hxxxx.mmy", "h_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\nvxxxx.mmy", "nv_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\nsxxxx.mmy", "ns_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\rqq.mmy", "r_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\sqq.mmy", "s_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\sqqv01.mmy", "sv01_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\d79sqq.mmy", "d79s_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\sprxxxx.mmy", "spr_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\HOqq.mmy", "spr_ho", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\ONK_SLqq.mmy", "spr_onk", "SHARED")
 
RETURN tn_result

FUNCTION CloseTemplates
 IF USED('d_et')
  USE IN d_et
 ENDIF 
 IF USED('h_et')
  USE IN h_et
 ENDIF 
 IF USED('nv_et')
  USE IN nv_et
 ENDIF 
 IF USED('ns_et')
  USE IN ns_et
 ENDIF 
 IF USED('r_et')
  USE IN r_et
 ENDIF 
 IF USED('s_et')
  USE IN s_et
 ENDIF 
 IF USED('sv01_et')
  USE IN sv01_et
 ENDIF 
 IF USED('d79s_et')
  USE IN d79s_et
 ENDIF 
 IF USED('spr_et')
  USE IN spr_et
 ENDIF 
 IF USED('spr_ho')
  USE IN spr_ho
 ENDIF 
 IF USED('spr_onk')
  USE IN spr_onk
 ENDIF 
RETURN 

FUNCTION CloseItems
 IF USED('dfile')
  USE IN dfile
 ENDIF 
 IF USED('hfile')
  USE IN hfile
 ENDIF 
 IF USED('nvfile')
  USE IN nvfile
 ENDIF 
 IF USED('nsfile')
  USE IN nsfile
 ENDIF 
 IF USED('rfile')
  USE IN rfile
 ENDIF 
 IF USED('sfile')
  USE IN sfile
 ENDIF 
 IF USED('dsfile')
  USE IN dsfile
 ENDIF 
 IF USED('sprfile')
  USE IN sprfile
 ENDIF 
 IF USED('hofile')
  USE IN hofile
 ENDIF 
 IF USED('onkfile')
  USE IN onkfile
 ENDIF 
RETURN 

FUNCTION CompFields(NameOfFile)
 FOR nFld = 1 TO fld_1
  IF (tabl_1(nFld,1) == tabl_2(nFld,1)) AND ;
     (tabl_1(nFld,2) == tabl_2(nFld,2)) AND ;
     (tabl_1(nFld,3) == tabl_2(nFld,3))
  ELSE 
   =CloseItems()
*   =ClDir()
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Wrong structure of ] + NameOfFile
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   RETURN 0 
  ENDIF 
 ENDFOR 
RETURN 1

FUNCTION DiffFields(NameOfFile)
 =CloseItems()
* =ClDir()
 m.csubject = m.csubject1 + '08' +m.csubject2
 m.cerrmessage = [Wrong number of fields in ] + NameOfFile
 IF m.llIsSubject = .F.
  m.llIsSubject = .T.
  poi.WriteLine('Subject: '+m.csubject)
 ENDIF 
 poi.WriteLine('BodyPart: ' + m.cerrmessage)
 poi.Close
RETURN 

FUNCTION IsAisDir()
 IF !fso.FolderExists(pAisOms)
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms, 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\INPUT')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser\INPUT', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\OUTPUT')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser\OUTPUT', 0+16, '')
  RETURN .F. 
 ENDIF

RETURN .T. 

FUNCTION WriteInBFile(BFullName, TextToWrite)
 CFG = FOPEN(BFullName,12)
 IsMyCommentExists = .F.
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  IF UPPER(READCFG) = 'MYCOMMENT'
   IsMyCommentExists = .T.
   LOOP 
  ENDIF 
 ENDDO
 IF !IsMyCommentExists
  nFileSize = FSEEK(CFG,0,2)
  =FWRITE(CFG, TextToWrite)
 ENDIF 
 = FCLOSE (CFG)
RETURN 

FUNCTION MakeDoctor
 tnvFile = lcDir+'\'+nvItem
 oSettings.CodePage('&tnvFile', 866, .t.)
* tnsFile = lcDir+'\'+nsItem
* oSettings.CodePage('&tnsFile', 866, .t.)
* USE (tnsFile) IN 0 ALIAS lcDoctor  EXCLUSIVE 
 USE (tnvFile) IN 0 ALIAS lcDoctor2 EXCLUSIVE 
 SELECT lcDoctor2
* INDEX on pcod TAG pcod 
* SET ORDER TO pcod 
* SELECT lcDoctor
* SET RELATION TO pcod INTO lcDoctor2
 SCAN 
  SCATTER MEMVAR
*  m.prvs_1 = lcDoctor2.prvs_1
*  m.prvs_2 = lcDoctor2.prvs_2
*  m.prvs_3 = lcDoctor2.prvs_3
*  m.prvs_4 = lcDoctor2.prvs_4
*  m.prvs_5 = lcDoctor2.prvs_5
*  m.prvs_6 = lcDoctor2.prvs_6
*  m.d_ser_1 = lcDoctor2.d_ser_1
*  m.d_ser_2 = lcDoctor2.d_ser_2
*  m.d_ser_3 = lcDoctor2.d_ser_3
*  m.d_ser_4 = lcDoctor2.d_ser_4
*  m.d_ser_5 = lcDoctor2.d_ser_5
*  m.d_ser_6 = lcDoctor2.d_ser_6
*  m.ps_1 = lcDoctor2.ps_1
*  m.ps_2 = lcDoctor2.ps_2
*  m.ps_3 = lcDoctor2.ps_3
*  m.ps_4 = lcDoctor2.ps_4
*  m.ps_5 = lcDoctor2.ps_5
*  m.ps_6 = lcDoctor2.ps_6

  m.dr = CTOD(SUBSTR(m.dr,7,2)+'.'+SUBSTR(m.dr,5,2)+'.'+SUBSTR(m.dr,1,4))

  INSERT INTO Doctor FROM MEMVAR 
*  m.un_key = m.mcod + ' ' + m.pcod
*  IF !SEEK(m.un_key, 'doctor_sv', 'unkey')
*   INSERT INTO Doctor_sv FROM MEMVAR 
*  ENDIF 
 ENDSCAN 
* SET RELATION OFF INTO lcDoctor2
 USE 
* SELECT lcDoctor2
* SET ORDER TO 
* DELETE TAG ALL 
* USE 
 USE IN Doctor
* fso.DeleteFile(lcDir+'\'+nvItem)
RETURN 

FUNCTION MakeOtdel
 tFile = lcDir+'\'+dItem
 oSettings.CodePage('&tFile', 866, .t.)
 USE (tFile) IN 0 ALIAS lcOtdel  EXCLUSIVE 
 SELECT lcOtdel
 SCAN 
  SCATTER FIELDS EXCEPT mcod MEMVAR
  INSERT INTO Otdel FROM MEMVAR 
*  m.un_key = m.mcod+' '+m.iotd
*  IF !SEEK(m.un_key, 'otdel_sv', 'unkey')
*   INSERT INTO Otdel_sv FROM MEMVAR 
*  ENDIF 
 ENDSCAN 
 USE 
 USE IN Otdel
 fso.DeleteFile(lcDir+'\'+dItem)
RETURN 

FUNCTION MakePeople
 tFile = lcDir+'\'+rItem
 oSettings.CodePage('&tFile', 866, .t.)
 USE &lcDir\&rItem IN 0 ALIAS lcRFile  EXCLUSIVE 
 SELECT lcRFile
 m.paz = 0 
 SCAN 
  SCATTER FIELDS EXCEPT tip_p MEMVAR 
  m.recid_lpu = m.recid
  m.period    = m.gcPeriod
  m.tipp      = tip_p

  m.prmcod  = IIF(SEEK(m.prik, 'sprlpu'), sprlpu.mcod, '')
  m.prmcods = IIF(SEEK(m.priks, 'sprlpu'), sprlpu.mcod, '')
  
  IF m.SaveInitPr = 1 && Сверка с номерником включена, режим по умолчанию

   m.qq = ''
   m.sv = ''
   DO CASE 
    CASE m.tipp='В'
     m.polis = ALLTRIM(sn_pol)
     IF LEN(m.polis)=9
      m.lpuid   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_tera, 0)
      m.prmcod  = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
      m.lpuids  = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_stom, 0)
      m.prmcods = IIF(SEEK(m.lpuids, 'sprlpu'), sprlpu.mcod, '')
     ENDIF 

    CASE INLIST(m.tipp,'П','Э','К')
     *m.polis   = enp
     m.polis   = LEFT(sn_pol,16)
     m.lpuid   = IIF(SEEK(m.polis, 'enp'), enp.lpu_tera, 0)
     m.prmcod  = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
     m.lpuids = IIF(SEEK(m.polis, 'enp'), enp.lpu_stom, 0)
     m.prmcods = IIF(SEEK(m.lpuids, 'sprlpu'), sprlpu.mcod, '')

    CASE m.tipp='С'
     m.polis = ALLTRIM(sn_pol)
     m.lpuid   = IIF(SEEK(m.polis, 'kms'), kms.lpu_tera, 0)
     m.prmcod  = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.mcod, '')
     m.lpuids = IIF(SEEK(m.polis, 'kms'), kms.lpu_stom, 0)
     m.prmcods = IIF(SEEK(m.lpuids, 'sprlpu'), sprlpu.mcod, '')

    OTHERWISE 
     && оставляем так, как подало МО!
   ENDCASE 
  
  ENDIF 
  
  m.prmcod  = IIF(m.d_type='9', '', m.prmcod)
  m.prmcods = IIF(m.d_type='9', '', m.prmcods)


  RELEASE m.recid, m.d_beg, m.d_end, m.tip_p, m.s_all
  INSERT INTO People FROM MEMVAR
  m.paz = m.paz + 1
 ENDSCAN 
 USE 
 fso.DeleteFile(lcDir+'\'+rItem)
RETURN 

FUNCTION MakeHO
 tFile = lcDir+'\'+hoItem
 IF !fso.FileExists(tFile)
  RETURN 
 ENDIF 
 oSettings.CodePage('&tFile', 866, .t.)
 USE &lcDir\&hoItem IN 0 ALIAS lcHOFile  EXCLUSIVE 
 SELECT lcHOFile
 IF fso.FileExists(lcDir+'\ho'+m.qcod+'.dbf')
  fso.DeleteFile(lcDir+'\ho'+m.qcod+'.dbf')
 ENDIF 
 COPY TO &lcDir\ho&qcod
 USE 
 USE &lcDir\ho&qcod IN 0 ALIAS lcHOFile  EXCLUSIVE 
 ALTER TABLE lcHOFile ALTER COLUMN c_i c(30)
 SELECT lcHOFile
 INDEX on sn_pol+c_i+PADL(cod,6,'0') TAG unik
 INDEX on sn_pol+c_i TAG snp_ci
 USE 
 IF fso.FileExists(lcDir+'\'+hoItem)
  fso.DeleteFile(lcDir+'\'+hoItem)
 ENDIF 
RETURN 

FUNCTION MakeOnkFile(para1)
 PRIVATE tFile
 m.tFile = ALLTRIM(para1)
 IF !fso.FileExists(m.tFile)
  RETURN 
 ENDIF 

 m.ntFile = STRTRAN(m.tFile, m.mmy, 'dbf')
 fso.CopyFile(m.tFile, m.ntFile)

 IF !fso.FileExists(m.ntFile)
  RETURN 
 ENDIF 
 
 fso.DeleteFile(m.tFile)

 oSettings.CodePage('&ntFile', 866, .t.)
RETURN 

FUNCTION MakeOnk
 tFile = lcDir+'\'+onkItem
 IF !fso.FileExists(tFile)
  RETURN 
 ENDIF 
 oSettings.CodePage('&tFile', 866, .t.)
 USE &lcDir\&OnkItem IN 0 ALIAS lcOnkFile  EXCLUSIVE 
 SELECT lcOnkFile
 INDEX ON sn_pol TAG sn_pol
 INDEX ON c_i TAG c_i
 COPY TO &lcDir\onk_sl&qcod WITH cdx 
 DELETE TAG ALL 
 USE 
 IF fso.FileExists(lcDir+'\'+OnkItem)
  fso.DeleteFile(lcDir+'\'+OnkItem)
 ENDIF 
RETURN 

FUNCTION MakeTalon
 tFile = lcDir+'\'+sItem
 oSettings.CodePage('&tFile', 866, .t.)
 USE pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN' IN 0 ALIAS Tarif SHARED ORDER cod 
 USE &lcDir\&sItem  IN 0 ALIAS lcSFile EXCLUSIVE 

 SELECT lcSFile
 m.nsch = RECCOUNT('lcSFile')
 SCAN
  SCATTER MEMVAR 
  m.recid_lpu = m.recid
  RELEASE m.recid

  IF IsUsl(m.cod) OR (IsKdP(m.cod) AND !IsEko(m.cod))
   IF !SEEK(m.sn_pol, 'pazamb')
    INSERT INTO pazamb (sn_pol) VALUES (m.sn_pol)
    m.krank = m.krank + 1
   ENDIF 
  ENDIF 
   
  IF IsMes(m.cod)
   IF !SEEK(m.c_i, 'pazst')
    INSERT INTO pazst (c_i) VALUES (m.c_i)
    m.paz_st = m.paz_st + 1
   ENDIF 
  ENDIF 

  IF IsKdS(m.cod) OR IsEko(m.cod)
   IF !SEEK(m.sn_pol, 'pazdst')
    INSERT INTO pazdst (sn_pol) VALUES (m.sn_pol)
    m.paz_dst = m.paz_dst + 1
   ENDIF 
  ENDIF 

  IF IsVMP(m.cod)
   IF !SEEK(m.sn_pol, 'pazvmp')
    INSERT INTO pazvmp (sn_pol) VALUES (m.sn_pol)
    m.paz_vmp = m.paz_vmp + 1
   ENDIF 
  ENDIF 

  m.otd    = m.iotd
  *m.s_all  = fsumm(m.cod, m.tip, m.k_u, m.IsVed)
  m.s_all  = fsumm(m.cod, m.tip, IIF(BETWEEN(m.cod,97107,97158), m.kd_fact, m.k_u), m.IsVed)
  m.period = m.gcPeriod
  *m.profil = IIF(SEEK(m.cod, 'profus'), ALLTRIM(profus.profil), '')
  m.profil = SUBSTR(m.iotd,4,3)
  m.n_kd   = IIF(SEEK(m.cod,'tarif'), tarif.n_kd, 0)

  m.s_pred = m.s_pred + s_all
  
  IF OCCURS(' ',ALLTRIM(m.pcod)) > 0 && Составной код врача
   m.pcod  = SUBSTR(ALLTRIM(m.pcod),1,AT(' ',ALLTRIM(m.pcod))-1)
  ELSE 
   m.pcod  = ALLTRIM(LEFT(ALLTRIM(m.pcod),10))
  ENDIF 
  
  INSERT INTO Talon FROM MEMVAR 
 ENDSCAN 
 USE 

 fso.DeleteFile(lcDir+'\'+sItem)
 USE IN Tarif

 SELECT sn_pol, 1 AS tip_p, MIN(d_u) as min_p, MAX(d_u) as max_p, SUM(s_all) as s_all FROM talon WHERE EMPTY(tip) ;
   GROUP BY sn_pol INTO CURSOR intp
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 SELECT sn_pol, 2 AS tip_p, MIN(d_u-k_u) as min_s, MAX(d_u) as max_s, SUM(s_all) as s_all FROM talon GROUP BY sn_pol ;
  WHERE !EMPTY(tip) INTO CURSOR ints
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 * сюда вставить создание файла hosp
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\hosp.dbf')
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\hosp.dbf')
  ENDIF 

  SELECT c_i, SPACE(25) as sn_pol, MAX(d_u)-SUM(k_u)+1 as d_pos, MAX(d_u) as d_vip, coun(*) as cnt, ;
  	SUM(k_u) as k_u FROM talon WHERE IsMes(cod) OR IsVMP(cod) ;
  	GROUP BY c_i ORDER BY c_i ASC INTO CURSOR cur_h READWRITE 
  SELECT cur_h
  INDEX on c_i TAG c_i 
  INDEX on sn_pol TAG sn_pol
  INDEX on d_pos TAG d_pos
  SET ORDER TO c_i
  SET ORDER TO c_i IN talon 
  SET RELATION TO c_i INTO talon 
  REPLACE ALL sn_pol WITH talon.sn_pol
  SET RELATION OFF INTO talon 

  IF tMonth=1
  ELSE
   m.p_period = STR(tYear,4)+PADL(tMonth-1,2,'0') 
   IF fso.FileExists(pBase+'\'+m.p_period+'\'+m.mcod+'\hosp.dbf')
    APPEND FROM &pBase\&p_period\&mcod\hosp
   ENDIF 
  ENDIF 
  
  SET ORDER TO d_pos
  COPY TO &pBase\&gcPeriod\&mcod\hosp CDX 
  USE 
 * сюда вставить создание файла hosp

 USE IN talon
 SELECT people
 SET RELATION TO sn_pol INTO intp
 SET RELATION TO sn_pol INTO ints ADDITIVE 
 m.t_sum = 0
 SCAN 
  m.d_beg = MIN(IIF(!EMPTY(intp.min_p), intp.min_p, m.tdat2), IIF(!EMPTY(ints.min_s), ints.min_s, m.tdat2))
  m.d_end = MAX(intp.max_p, ints.max_s)
  DO CASE 
   CASE intp.tip_p == 1 AND ints.tip_p != 2
    m.tip_p = 1
   CASE intp.tip_p != 1 AND ints.tip_p == 2
    m.tip_p = 2
   CASE intp.tip_p == 1 AND ints.tip_p == 2
    m.tip_p = 3
   OTHERWISE 
    m.tip_p = 0
  ENDCASE 

  m.s_all = IIF(!EMPTY(intp.s_all), intp.s_all, 0) + IIF(!EMPTY(ints.s_all), ints.s_all, 0)
  m.t_sum = m.t_sum + m.s_all
  REPLACE people.d_beg WITH m.d_beg, people.d_end WITH m.d_end, tip_p WITH m.tip_p,;
  	people.s_all WITH m.s_all

  *REPLACE people.s_all WITH m.s_all
  *m.s_all = IIF(!EMPTY(intp.s_all), intp.s_all, 0) + IIF(!EMPTY(ints.s_all), ints.s_all, 0)

 ENDSCAN 
 SET RELATION OFF INTO ints
 SET RELATION OFF INTO intp
 USE IN people
 USE IN ints
 USE IN intp 
 
 IF m.s_pred != m.t_sum
  *MESSAGEBOX('m.s_pred: '+TRANSFORM(m.s_pred, '9999999.99')+CHR(13)+CHR(10)+;
  	'm.t_sum: '+TRANSFORM(m.t_sum, '9999999.99')+CHR(13)+CHR(10), 0+64, m.mcod)
 ENDIF 
RETURN

FUNCTION CreateFilesStructure
* CREATE TABLE (People) ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1,;
   mcod c(7), prmcod c(7), period c(6), d_beg d, d_end d, s_all n(11,2), ;
   tip_p n(1), sn_pol c(25), tipp c(1), enp c(16), qq c(2), ;
   fam c(25), im c(20), ot c(20), w n(1), dr d, ;
   ul n(5), dom c(7), kor c(5), str c(5), kv c(5), d_type c(1), ;
   sv c(3), recid_lpu c(7), IsPr L)
 CREATE TABLE (People) ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1,;
   mcod c(7), prmcod c(7), prmcods c(7), period c(6), d_beg d, d_end d, s_all n(11,2), ;
   tip_p n(1), sn_pol c(25), tipp c(1), enp c(16), qq c(2), ;
   fam c(25), im c(20), ot c(20), w n(1), dr d, d_type c(1), ;
   sv c(3), recid_lpu c(7), IsPr L, fil_id n(6))
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON sn_pol TAG sn_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX on dr TAG dr
 INDEX on s_all TAG s_all
 USE 
* CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6),  ;
	 pcod c(10), otd c(8), cod n(6), tip c(1), d_u d, ;
	 k_u n(3), d_type c(1), s_all n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3),;
	 codnom c(14), kur n(5,3), ds_2 c(6), ds_3 c(6), det n(1), k2 n(5,3), tipgr c(1), ;
	 vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17),;
	 ord n(1), date_ord d, lpu_ord n(6), recid_lpu c(7), fil_id n(6), ;
	 IsPr L, vz l, mp c(1), n_kd n(3), f_type c(2)) && Новая структура с 01082018
 *CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6),  ;
	 pcod c(10), otd c(8), cod n(6), tip c(1), d_u d, ;
	 k_u n(4), kd_fact n(3), d_type c(1), s_all n(11,2), s_lek n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3),;
	 codnom c(14), kur n(5,3), ds_2 c(6), ds_3 c(6), det n(1), k2 n(5,3), tipgr c(1), ;
	 vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17),;
	 ord n(1), date_ord d, lpu_ord n(6), recid_lpu c(7), fil_id n(6), ;
	 ds_onk n(1), p_cel c(3), dn n(1), reab n(1), tal_d d, napr_v_in n(1), ;
	 c_zab n(1), napr_usl c(15), vid_vme c(15),IsPr L, vz l, mp c(1), n_kd n(3), f_type c(2), ;
	 mm c(1), typ c(1)) && Новая структура с 01082018
 CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6), pcod c(10), otd c(8), ;
	 cod n(6), tip c(1), d_u d, k_u n(4), kd_fact n(4), n_kd n(3), d_type c(1), s_all n(11,2), ;
	 s_lek n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3), kur n(5,3), ds_2 c(6), ds_3 c(6), ;
	 det n(1), k2 n(5,3), vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17), ord n(1), date_ord d, ;
	 lpu_ord n(6), recid_lpu c(7), fil_id n(6), ds_onk n(1), p_cel c(3), dn n(1), reab n(1), tal_d d, napr_v_in n(1), ;
	 c_zab n(1), mp c(1), typ c(1), dop_r n(2), vz n(1), IsPr L,;
	 sqlid i, sqldt t, prcell c(3), nsif n(1)) && Убрал поля codnom, napr_usl, vid_vme, tipgr, mm, vz, f_type 01.06.2019

 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON c_i TAG c_i
 INDEX ON sn_pol TAG sn_pol
 INDEX ON otd TAG otd
 INDEX on pcod TAG pcod
 INDEX ON ds TAG ds
 INDEX ON d_u TAG d_u
 INDEX ON cod TAG cod
 INDEX ON profil TAG profil
 INDEX ON sn_pol+STR(cod,6)+DTOS(d_u) TAG ExpTag
 INDEX ON sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u) TAG unik 
 INDEX ON tip TAG tip
 INDEX ON s_all TAG s_all
 USE 
 
 CREATE TABLE (Otdel) ;
	(recid c(6), mcod c(7), iotd c(8), name c(100), pr_name c(100), cnt_bed n(5), fil_id n(6))
 INDEX ON iotd TAG iotd
 USE 

 CREATE TABLE (Doctor) ;
   (pcod c(10),sn_pol c(25),fam c(25),im c(20),ot c(20),dr d, w n(1),;
    prvs n(4), d_ser d, d_ser2 d, d_prik d, iotd c(8),;
	lgot_r c(1),c_ogrn c(15),lpu_id n(6), fil_id n(6))
 INDEX ON pcod TAG pcod
 USE 

 CREATE TABLE (Error) (f c(1), c_err c(3), et n(1), detail c(1), rid i, tip n(1), dt t, usr c(6), ;
 	"comment" c(250), sqlid i, sqldt t)
 INDEX FOR UPPER(f)='R' ON rid TAG rrid
 INDEX FOR UPPER(f)='S' ON rid TAG rid
 USE 

 CREATE TABLE (mError) ;
  (rid i autoinc, RecId i, cod n(6), k_u n(4), tip c(1), et c(1), ee c(1), usr c(6), d_exp d,;
   e_cod n(6), e_ku n(3), e_tip c(1), err_mee c(3), osn230 c(5), e_period c(6),  ;
   koeff n(4,2), straf n(4,2), docexp c(7), s_all n(11,2), s_1 n(11,2), s_2 n(11,2), impdata d,;
   subet n(1), reason c(1), n_akt c(15), t_akt c(2), d_edit d, d_akt d)
 INDEX ON rid TAG rid 
 INDEX ON RecId TAG recid
 *INDEX ON PADL(recid,6,'0')+et TAG id_et
 *INDEX ON PADL(recid,6,'0')+et+LEFT(err_mee,2) TAG unik
 INDEX ON PADL(recid,6,'0')+et+docexp+reason TAG id_et
 INDEX ON PADL(recid,6,'0')+et+docexp+reason+LEFT(err_mee,2) TAG unik
  INDEX ON PADL(recid,6,'0')+et TAG uniket
 USE 

 CREATE CURSOR pazamb (sn_pol c(25), d_beg d, d_end d, tip_p n(1), s_all n(11,2))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
  
 CREATE CURSOR pazdst (sn_pol c(25), d_beg d, d_end d, tip_p n(1), s_all n(11,2))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazst (c_i c(30), d_beg d, d_end d, tip_p n(1), s_all n(11,2))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR pazvmp (sn_pol c(25), d_beg d, d_end d, tip_p n(1), s_all n(11,2))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol
RETURN 

FUNCTION OpenLocalFiles
 USE (People) IN 0 ALIAS People SHARED
 USE (Talon)  IN 0 ALIAS Talon  SHARED   
 USE (Otdel)  IN 0 ALIAS Otdel  SHARED 
 USE (Doctor) IN 0 ALIAS Doctor SHARED 
RETURN 

FUNCTION CheckFilesStucture
 fld_1 = AFIELDS(tabl_1, 'dfile') && Проверка d-файла 
 fld_2 = AFIELDS(tabl_2, 'd_et')  && 1 столбец - название, 2 - тип,  3 - размерность, 4 - нулей после запятой
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(dItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(dItem)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 

* fld_1 = AFIELDS(tabl_1, 'hfile') && Проверка h-файла 
* fld_2 = AFIELDS(tabl_2, 'h_et')  && 1 столбец - название, 2 - тип,  3 - размерность, 4 - нулей после запятой
* IF fld_1 == fld_2 && Кол-во полей совпадает!
*  FieldsIdent = CompFields(hItem) && 0 - есть отличия, 1 - полное совпадение
*  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
*   =ClDir()
*   RETURN .F.
*  ENDIF 
* ELSE 
*  =DiffFields(hItem)
*  =CopyToTrash(m.InDir,2)
*  =ClDir()
*  RETURN .F.
* ENDIF 

 fld_1 = AFIELDS(tabl_1, 'nvfile') && проверка nv-файла
 fld_2 = AFIELDS(tabl_2, 'nv_et')
 IF 3=2
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(nvItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(nvItem)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 
 ENDIF 

* fld_1 = AFIELDS(tabl_1, 'nsfile') && проверка ns-файла
* fld_2 = AFIELDS(tabl_2, 'ns_et')
* IF fld_1 == fld_2 && Кол-во полей совпадает!
*  FieldsIdent = CompFields(nsItem) && 0 - есть отличия, 1 - полное совпадение
*  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
*   =ClDir()
*   RETURN .F.
*  ENDIF 
* ELSE 
*  =DiffFields(nsItem)
*  =CopyToTrash(m.InDir,2)
*  =ClDir()
*  RETURN .F.
* ENDIF 

 fld_1 = AFIELDS(tabl_1, 'rfile') && проверка r-файла
 fld_2 = AFIELDS(tabl_2, 'r_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(rItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(rItem)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 
IF 3=2
 m.SCompResult = .T.
 fld_1 = AFIELDS(tabl_1, 'sfile') &&& проверка s-файла
 fld_2 = AFIELDS(tabl_2, 's_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(sItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   m.SCompResult = .F.
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  m.SCompResult = .F.
  =DiffFields(sItem)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 
ENDIF 
 IF 1=2
 IF m.SCompResult = .F.
 fld_1 = AFIELDS(tabl_1, 'sfile') &&& проверка s-файла
 fld_2 = AFIELDS(tabl_2, 'sv01_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(sItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(sItem)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 
 ENDIF 
 ENDIF 
 
 fld_1 = AFIELDS(tabl_1, 'sprfile') && проверка spr-файла
 fld_2 = AFIELDS(tabl_2, 'spr_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(sprItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(sprItem)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 

* IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+HOitem)
 IF USED('hofile') AND 1=2
 fld_1 = AFIELDS(tabl_1, 'hofile') && проверка ho-файла
 fld_2 = AFIELDS(tabl_2, 'spr_ho')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(hoItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(hoItem)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 
 ENDIF 

 IF USED('onkfile') AND 1=2
 fld_1 = AFIELDS(tabl_1, 'onkfile') && проверка onk_sl-файла
 fld_2 = AFIELDS(tabl_2, 'spr_onk')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields(onkItem) && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(onkItem)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 
 ENDIF 

RETURN .T. 

FUNCTION IsIPComplete
 DO CASE 
  CASE !UnzipGotoFileByName(dItem)
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Отсутствует ] + dItem + [ файл]
   UnzipClose()
   m.cmnt = m.cerrmessage
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   =CopyToTrash(m.MailDirName,1)
   RETURN .F. 

*  CASE !UnzipGotoFileByName(hItem)
*   m.csubject = m.csubject1 + '08' +m.csubject2
*   m.cerrmessage = [Отсутствует ] + hItem + [ файл]
*   UnzipClose()
*   m.cmnt = m.cerrmessage
*   IF m.llIsSubject = .F.
*    m.llIsSubject = .T.
*    poi.WriteLine('Subject: '+m.csubject)
*   ENDIF 
*   poi.WriteLine('BodyPart: ' + m.cerrmessage)
*   poi.Close
*   =CopyToTrash(m.MailDirName,1)
*   RETURN .F. 

  CASE !UnzipGotoFileByName(nvItem)
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Отсутствует ] + nvItem + [ файл]
   UnzipClose()
   m.cmnt = m.cerrmessage
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   =CopyToTrash(m.MailDirName,1)
   RETURN .F. 

*  CASE !UnzipGotoFileByName(nsItem)
*   m.csubject = m.csubject1 + '08' +m.csubject2
*   m.cerrmessage = [Отсутствует ] + nsItem + [ файл]
*   UnzipClose()
*   m.cmnt = m.cerrmessage
*   IF m.llIsSubject = .F.
*    m.llIsSubject = .T.
*    poi.WriteLine('Subject: '+m.csubject)
*   ENDIF 
*   poi.WriteLine('BodyPart: ' + m.cerrmessage)
*   poi.Close
*   =CopyToTrash(m.MailDirName,1)
*   RETURN .F. 

  CASE !UnzipGotoFileByName(rItem)
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Отсутствует ] + rItem + [ файл]
   UnzipClose()
   m.cmnt = m.cerrmessage
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   =CopyToTrash(m.MailDirName,1)
   RETURN .F. 

  CASE !UnzipGotoFileByName(sItem)
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Отсутствует ] + sItem + [ файл]
   UnzipClose()
   m.cmnt = m.cerrmessage
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   =CopyToTrash(m.MailDirName,1)
   RETURN .F. 

*  CASE !UnzipGotoFileByName(dsItem)
*   m.csubject = m.csubject1 + '08' +m.csubject2
*   m.cerrmessage = [Отсутствует ] + dsItem + [ файл]
*   UnzipClose()
*   m.cmnt = m.cerrmessage
*   IF m.llIsSubject = .F.
*    m.llIsSubject = .T.
*    poi.WriteLine('Subject: '+m.csubject)
*   ENDIF 
*   poi.WriteLine('BodyPart: ' + m.cerrmessage)
*   poi.Close
*   =CopyToTrash(m.MailDirName,1)
*   RETURN .F. 

  CASE !UnzipGotoFileByName(sprItem)
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Отсутствует ] + sprItem + [ файл]
   UnzipClose()
   m.cmnt = m.cerrmessage
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   =CopyToTrash(m.MailDirName,1)
   RETURN .F. 
 ENDCASE
RETURN .T.

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
   CASE UPPER(READCFG) = 'BODYPART'
    m.bparts   = m.bparts + 1
    m.bodypart = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dbparts(m.bparts,1) = ALLTRIM(SUBSTR(m.bodypart, 1, AT(" ",m.bodypart)-1))
  ENDCASE
 ENDDO
RETURN 
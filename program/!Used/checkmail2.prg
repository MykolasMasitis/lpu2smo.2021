PROCEDURE CheckMail2
PARAMETERS lcUser

IF !IsAisDir() && �������� ������� ����������, OMS, INPUT, OUTPUT
 RETURN 
ENDIF 

oMailDir        = fso.GetFolder(pAisOms+'\&lcUser\input')
MailDirName     = oMailDir.Path
oFilesInMailDir = oMailDir.Files
nFilesInMailDir = oFilesInMailDir.Count

MESSAGEBOX('���������� '+ALLTRIM(STR(nFilesInMailDir))+' ������!', 0+64, lcUser)

IF nFilesInMailDir<=0
 RETURN 
ENDIF 

IF OpenTemplates() != 0
 =CloseTemplates() 
 RETURN 
ENDIF 

WAIT "�������� �����..." WINDOW NOWAIT 
SELECT AisOms
prvorder = ORDER('aisoms')
SET ORDER TO 

m.un_id = SYS(3)

OldEscStatus = SET("Escape")
SET ESCAPE OFF 
CLEAR TYPEAHEAD 

FOR EACH oFileInMailDir IN oFilesInMailDir

 SCATTER MEMVAR BLANK
 m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)

 m.BFullName = oFileInMailDir.Path
 m.bname     = oFileInMailDir.Name
 m.recieved  = oFileInMailDir.DateLastModified
 m.lpuid     = 0
 m.processed = DATETIME()
 
 m.cfrom      = ''
 m.cdate      = ''
 m.cmessage   = ''
 m.resmesid   = ''
 m.csubject   = ''
 m.csubject1  = ''
 m.csubject2  = ''
 m.attachment = ''
 m.bodypart   = ''

 m.attaches   = 0 && ������� �������������� ������ � ����� ��
 DIMENSION dattaches(10,2)
 dattaches = ''

 m.bparts   = 0 && ������� �������������� ������ � ����� ��
 DIMENSION dbparts(10,2)
 dbparts = ''
 MESSAGEBOX(oFileInMailDir.Name,0+64,'')
 DO CASE 
 CASE LOWER(oFileInMailDir.Name) = 'b' AND SEEK(SUBSTR(oFileInMailDir.Name,2,7), 'sprlpu', 'mcod') AND ;
 	RIGHT(oFileInMailDir.Name,3)=m.mmy

 m.sent = {} && ����� ���������� ���� �������� �����!
 m.mcod = SUBSTR(oFileInMailDir.Name, 2, 7)
 WAIT m.mcod WINDOW NOWAIT 
   
 m.llIsSubject = .F.

 m.AisAddress = .F.
 m.adresat = 'pump.msk.oms' &&  'spuemias.msk.oms'
 m.lpuid = IIF(SEEK(m.mcod, "sprlpu", 'mcod'), sprlpu.lpu_id, 0)
 
 IF m.lpuid==0
  LOOP 
 ENDIF 
 
 IF EMPTY(m.mcod)
  LOOP 
 ENDIF 

 m.cokr    = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.cokr, "")
 m.moname  = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.name, "")
 m.usr     = IIF(SEEK(m.lpuid, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")

 IF EMPTY(m.usr) AND m.gcUser!='OMS'
  MESSAGEBOX('��� '+m.mcod+' �� "���������" � ����������� � USRLPU.DBF!',0+48,lcUser)
  LOOP 
 ENDIF 
 
 IF m.usr != m.gcUser AND m.gcUser!='OMS'
  MESSAGEBOX('USR' ,0+48,lcUser)
  MESSAGEBOX('USR' ,0+48,STR(m.lpuid,4))
  LOOP 
 ENDIF 

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

 && ���� �� ���� �� ���� zip-�����?
 llIsOneZip = .F.
  ffile = fso.GetFile(MailDirName + '\' + m.bname)
  IF ffile.size >= 2
   fhandl = ffile.OpenAsTextStream
   lcHead = fhandl.Read(2)
   fhandl.Close
  ELSE 
   lcHead = ''
  ENDIF 

  IF lcHead == 'PK' && ��� zip-����!
   ZipName = pAisOms+'\'+lcUser+'\input\'+m.bname
   UnzipOpen(ZipName)
   rItem   = 'R' + m.qcod + '.' + m.mmy
   sItem   = 'S' + m.qcod + '.' + m.mmy
   IF UnzipGotoFileByName(rItem) AND UnzipGotoFileByName(sItem)
    llIsOneZip = .t.
    UnzipClose()
*    EXIT 
   ENDIF 
   UnzipClose()
  ENDIF 
 && ���� �� ���� �� ���� zip-�����?
 
 && ���� ��� �� ������ zip-������
 IF llIsOneZip == .F.
  LOOP 
 ENDIF 
 && ���� ��� �� ������ zip-������

 poi = fso.CreateTextFile(pAisOms+'\&lcUser\output\'+m.tansfile)
 poi.WriteLine('To: '+m.cfrom)
 poi.WriteLine('Message-Id: ' + m.messageid)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: ' + m.cmessage)

 && ��������� ������������� ������� - ������� 5 ������!
 UnzipOpen(ZipName)

 hItem    = 'H'  + STR(m.lpuid,4) + '.' + m.mmy
 dItem    = 'D'  + STR(m.lpuid,4) + '.' + m.mmy
 nvItem   = 'NV' + STR(m.lpuid,4) + '.' + m.mmy
 nsItem   = 'NS' + STR(m.lpuid,4) + '.' + m.mmy
 rItem    = 'R' + m.qcod + '.' + m.mmy
 sItem    = 'S' + m.qcod + '.' + m.mmy
 hoItem   = 'HO' + m.qcod + '.' + m.mmy
 onkItem  = 'ONK_SL' + m.qcod + '.' + m.mmy
 dsItem   = 'D79S' + m.qcod + '.' + m.mmy
 sprItem  = 'SPR' + STR(m.lpuid,4) + '.' + m.mmy

 IF !IsIPComplete()
  LOOP 
 ENDIF 
 
 UnzipClose()
 && ��������� ������������� ������� - ������� 5 ������!

 fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\*.*')

 && ���� ��� ���������� � ����� �������!
 InDirPeriod = pBase + '\' + m.gcPeriod
 IF !fso.FolderExists(InDirPeriod)
  fso.CreateFolder(InDirPeriod)
 ENDIF 
 InDir = pBase + '\' + m.gcPeriod + '\' + m.mcod
 IF !fso.FolderExists(InDir)
  fso.CreateFolder(InDir)
 ENDIF 

 fso.CopyFile(m.BFullName, InDir + '\' + m.bname)
 fso.DeleteFile(m.BFullName)

 ZipName = InDir + '\' + m.bname
 ZipDir  = InDir + '\'

 UnzipOpen(ZipName)

 UnzipGotoFileByName(rItem)
 UnzipFile(ZipDir)
 UnzipGotoFileByName(sItem)
 UnzipFile(ZipDir)
 UnzipGotoFileByName(dItem)
 UnzipFile(ZipDir)
 UnzipGotoFileByName(nvItem)
 UnzipFile(ZipDir)
 UnzipGotoFileByName(sprItem)
 UnzipFile(ZipDir)
 UnzipGotoFileByName(HOItem)
 UnzipFile(ZipDir)
 IF UnzipGotoFileByName(OnkItem)
  UnzipFile(ZipDir)
 ENDIF 

 UnzipClose()

 m.lcCurDir = pBase + '\' + m.gcPeriod + '\' + m.mcod+'\'
 SET DEFAULT TO (lcCurDir)

 =OpenFile("&dItem",  "dfile",  "SHARED")
* =OpenFile("&hItem",  "hfile",  "SHARED")
 =OpenFile("&nvItem", "nvfile", "SHARED")
* =OpenFile("&nsItem", "nsfile", "SHARED")
 =OpenFile("&rItem",  "rfile",  "SHARED")
 =OpenFile("&sItem",  "sfile",  "SHARED")
* =OpenFile("&dsItem", "dsfile", "SHARED")
 =OpenFile("&sprItem", "sprfile", "SHARED")
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+HOitem)
  =OpenFile("&HOitem", "hofile", "SHARED")
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+OnkItem)
  =OpenFile("&OnkItem", "onkfile", "SHARED")
 ENDIF 
 
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

 =CreateFilesStructure()
 =OpenLocalFiles()
 
 m.s_pred  = 0
 m.nsch    = 0
 m.krank   = 0
 m.paz_dst = 0
 m.paz_st  = 0
 m.paz_vmp = 0

 =MakePeople()
 =MakeTalon()
 =MakeOtdel() 
 =MakeDoctor()
 =MakeHO()

* IF NOT SEEK(m.cmessage, "aisoms", "cmessage")

  MailView.get_recs.value = MailView.get_recs.value + 1
  MailView.get_paz.value  = MailView.get_paz.value + m.paz
  MailView.get_nsch.value = MailView.get_nsch.value + m.nsch
  MailView.get_sum.value  = MailView.get_sum.value + m.s_pred
 
  UPDATE aisoms SET bname=m.bname, dname=m.dname, sent=m.sent, recieved=m.recieved, ;
   processed=m.processed, cfrom=m.cfrom, cmessage=m.cmessage, paz=m.paz, nsch=m.nsch, ;
   s_pred=m.s_pred, krank=m.krank, paz_dst=m.paz_dst, paz_st=m.paz_st, paz_vmp=m.paz_vmp,; 
   erz_id='', erz_status=0, sum_flk=0, ispr=.f. ; 
  WHERE mcod=m.mcod

*  INSERT INTO aisoms FROM MEMVAR && !!!

* ENDIF
 
 SELECT AisOms
 
 WAIT CLEAR 
 MailView.Refresh 

 CASE LOWER(oFileInMailDir.Name) = 'r' && ���� ��� r-����

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

 IF m.bparts > 0
  FOR npart = 1 TO m.bparts
   m.bpname   = ALLTRIM(dbparts(npart, 1))
   IF !EMPTY(m.bpname) AND fso.FileExists(MailDirName + '\' + m.bpname)
    fso.CopyFile(MailDirName + '\' + m.bpname, DaemonDir+'\'+m.bpname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.bpname)
   ENDIF 
  ENDFOR 
 ENDIF 

 SELECT AisOms
 
 WAIT CLEAR 
 MailView.Refresh 

 ENDCASE 

 IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 

NEXT && ���� �� ������

SET ESCAPE &OldEscStatus

=CloseTemplates() 

SET ORDER TO (prvorder)
MailView.Refresh
MailView.LockScreen=.f.

WAIT CLEAR 

nFilesInMailDir = oFilesInMailDir.Count
IF USED('lpuias')
 USE IN lpuais
ENDIF 

MESSAGEBOX('�������� '+ALLTRIM(STR(nFilesInMailDir))+' �������������� ��!', 0+64, lcUser)

RETURN 

FUNCTION CopyToTrash(lcPath, nTip)
 fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pAisOms+'\&lcUser\OUTPUT\'+m.badfile)
 fso.DeleteFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile)
 TrashDir = pTrash + '\' + m.mcod
 IF !fso.FolderExists(TrashDir)
  fso.CreateFolder(TrashDir)
 ENDIF 
 fso.CopyFile(lcPath + '\' + m.bname, TrashDir+'\'+m.bname)
 fso.DeleteFile(lcPath + '\' + m.bname)

 FOR nattach = 1 TO m.attaches
  IF !EMPTY(ALLTRIM(dattaches(nattach, 1)))
   fso.CopyFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)), TrashDir + '\' + ALLTRIM(dattaches(nattach, 2)))
   fso.DeleteFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)))
  ENDIF 
 ENDFOR 

 IF NOT SEEK(m.cmessage, "taisoms", "cmessage")
  INSERT INTO taisoms FROM MEMVAR 
 ENDIF
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
 IF fso.FileExists(OnkItem)
  DELETE FILE &OnkItem
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
 IF USED('people_sv')
  USE IN people_sv
 ENDIF 
 IF USED('talon_sv')
  USE IN talon_sv
 ENDIF 
 IF USED('otdel_sv')
  USE IN otdel_sv
 ENDIF 
 IF USED('doctor_sv')
  USE IN doctor_sv
 ENDIF 
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
  MESSAGEBOX('����������� ���������� '+pAisOms, 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser')
  MESSAGEBOX('����������� ���������� '+pAisOms+'\&lcUser', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\INPUT')
  MESSAGEBOX('����������� ���������� '+pAisOms+'\&lcUser\INPUT', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\OUTPUT')
  MESSAGEBOX('����������� ���������� '+pAisOms+'\&lcUser\OUTPUT', 0+16, '')
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
  m.qq = ''
  m.sv = ''
  m.recid_lpu = m.recid
  m.period = m.gcPeriod
  m.tipp = tip_p
  m.prmcod = IIF(SEEK(m.prik, 'sprlpu'), sprlpu.mcod, '')
  m.prmcod = IIF(m.d_type='9', '', m.prmcod)

  DO CASE 
   CASE m.tipp='�'
    m.polis = ALLTRIM(sn_pol)
    IF LEN(m.polis)=9
     m.lpuids  = IIF(SEEK(m.polis, 'outs', 'vsn'), outs.lpu_stom, 0)
     m.prmcods = IIF(SEEK(m.lpuids, 'pilots'), pilots.mcod, '')
    ENDIF 

   CASE INLIST(m.tipp,'�','�')
    m.polis = enp
    m.lpuids = IIF(SEEK(m.polis, 'outs', 'enp'), outs.lpu_stom, 0)
    m.prmcods = IIF(SEEK(m.lpuids, 'pilots'), pilots.mcod, '')

   CASE m.tipp='�'
    m.polis = ALLTRIM(sn_pol)
    m.lpuids = IIF(SEEK(m.polis, 'outs', 'kms'), outs.lpu_stom, 0)
    m.prmcods = IIF(SEEK(m.lpuids, 'pilots'), pilots.mcod, '')

   OTHERWISE 
  ENDCASE 

*  m.prmcods = IIF(SEEK(m.priks, 'sprlpu'), sprlpu.mcod, '')
*  m.prmcods = IIF(m.d_type='9', '', m.prmcods)
  RELEASE m.recid, m.d_beg, m.d_end, m.tip_p, m.s_all
  INSERT INTO People FROM MEMVAR
*  INSERT INTO People_sv FROM MEMVAR 
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
 COPY TO &lcDir\ho&qcod
 USE 
 USE &lcDir\ho&qcod IN 0 ALIAS lcHOFile  EXCLUSIVE 
 ALTER TABLE lcHOFile ALTER COLUMN c_i c(30)
 SELECT lcHOFile
 INDEX on sn_pol+c_i+PADL(cod,6,'0') TAG unik
 USE 
 fso.DeleteFile(lcDir+'\'+hoItem)
RETURN 

FUNCTION MakeOnk
 tFile = lcDir+'\'+onkItem
 IF !fso.FileExists(tFile)
  RETURN 
 ENDIF 
 oSettings.CodePage('&tFile', 866, .t.)
 USE &lcDir\&OnkItem IN 0 ALIAS lcOnkFile  EXCLUSIVE 
 SELECT lcOnkFile
 COPY TO &lcDir\onk_sl&qcod
 USE 
 IF fso.FileExists(lcDir+'\'+OnkItem)
  fso.DeleteFile(lcDir+'\'+OnkItem)
 ENDIF 
RETURN 

FUNCTION MakeTalon
 tFile = lcDir+'\'+sItem
 oSettings.CodePage('&tFile', 866, .t.)
* tFile = lcDir+'\'+dsItem
* oSettings.CodePage('&tFile', 866, .t.)
 USE pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN' IN 0 ALIAS Tarif SHARED ORDER cod 
 USE &lcDir\&sItem  IN 0 ALIAS lcSFile EXCLUSIVE 
* USE &lcDir\&dsItem IN 0 ALIAS lcDSFile EXCLUSIVE 
* SELECT lcDSFile
* INDEX ON recid TAG recid 
* SET ORDER TO recid
* SELECT lcSFile
* SET RELATION TO recid INTO lcDSFile

 SELECT lcSFile
 m.nsch = RECCOUNT('lcSFile')
 SCAN
  SCATTER MEMVAR 
*  SCATTER FIELDS lcDSFile.novor, lcDSFile.ds_s, lcDSFile.ds_p, lcDSFile.profil, lcDSFile.rslt,;
   lcDSFile.prvs, lcDSFile.ord, lcDSFile.ishod, lcDSFile.fil_id MEMVAR 
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

  m.otd   = m.iotd
  m.s_all = fsumm(m.cod, m.tip, m.k_u, m.IsVed)
  m.period = m.gcPeriod
  m.profil = IIF(SEEK(m.cod, 'profus'), ALLTRIM(profus.profil), '')
  m.n_kd = IIF(SEEK(m.cod,'tarif'), tarif.n_kd, 0)

  m.s_pred = m.s_pred + s_all
  
  IF OCCURS(' ',ALLTRIM(m.pcod)) > 0 && ��������� ��� �����
   m.pcod  = SUBSTR(ALLTRIM(m.pcod),1,AT(' ',ALLTRIM(m.pcod))-1)
*   m.docvs = SUBSTR(ALLTRIM(m.pcod),AT(' ',ALLTRIM(m.pcod))+1)
  ELSE 
   m.pcod  = ALLTRIM(LEFT(ALLTRIM(m.pcod),10))
*   m.docvs = ''
  ENDIF 
  
  INSERT INTO Talon FROM MEMVAR 
*  INSERT INTO Talon_sv FROM MEMVAR 
 ENDSCAN 
* SET RELATION OFF INTO lcDSFile
 USE 
* SELECT lcDSFile
* SET ORDER TO 
* DELETE TAG ALL 
* USE 
 fso.DeleteFile(lcDir+'\'+sItem)
* fso.DeleteFile(lcDir+'\'+dsItem)
 USE IN Tarif

 SELECT sn_pol, 1 AS tip_p, MIN(d_u) as min_p, MAX(d_u) as max_p, SUM(s_all) as s_all FROM talon WHERE EMPTY(tip) ;
   GROUP BY sn_pol INTO CURSOR intp
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 SELECT sn_pol, 2 AS tip_p, MIN(d_u-k_u) as min_s, MAX(d_u) as max_s, SUM(s_all) as s_all FROM talon GROUP BY sn_pol ;
  WHERE !EMPTY(tip) INTO CURSOR ints
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
 USE IN talon
 SELECT people
 SET RELATION TO sn_pol INTO intp
 SET RELATION TO sn_pol INTO ints ADDITIVE 
* SET ORDER TO unkey IN people_sv
* SET RELATION TO mcod+sn_pol INTO people_sv ADDITIVE 
* SET ORDER TO sn_pol IN people_sv
* SET RELATION TO sn_pol INTO people_sv ADDITIVE 
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
  REPLACE people.d_beg WITH m.d_beg, people.d_end WITH m.d_end, tip_p WITH m.tip_p
  m.s_all = IIF(!EMPTY(intp.s_all), intp.s_all, 0) + IIF(!EMPTY(ints.s_all), ints.s_all, 0)
  REPLACE people.s_all WITH m.s_all

*  REPLACE people_sv.d_beg WITH m.d_beg, people_sv.d_end WITH m.d_end, people_sv.tip_p WITH m.tip_p
  m.s_all = IIF(!EMPTY(intp.s_all), intp.s_all, 0) + IIF(!EMPTY(ints.s_all), ints.s_all, 0)
*  REPLACE people_sv.s_all WITH m.s_all

 ENDSCAN 
* SET RELATION OFF INTO people_sv
 SET RELATION OFF INTO ints
 SET RELATION OFF INTO intp
 USE IN people
 USE IN ints
 USE IN intp 
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
   sv c(3), recid_lpu c(7), IsPr L)
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON sn_pol FOR IsPplValid() TAG sn_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX on dr TAG dr
 INDEX on s_all TAG s_all
 USE 
 IF m.qobjid != 3415
 CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6),  ;
	 pcod c(10), otd c(8), cod n(6), tip c(1), d_u d, ;
	 k_u n(3), d_type c(1), s_all n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3),;
	 codnom c(14), kur n(5,3), ds_2 c(6), ds_3 c(6), det n(1), k2 n(5,3), tipgr c(1), ;
	 vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17),;
	 ord n(1), date_ord d, lpu_ord n(6), recid_lpu c(7), fil_id n(6), IsPr L, vz l, mp c(1), n_kd n(3), f_type c(2), mm c(1))
 ELSE 
 CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6),  ;
	 pcod c(10), otd c(8), cod n(6), tip c(1), d_u d, ;
	 k_u n(3), d_type c(1), s_all n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3),;
	 codnom c(14), kur n(5,3), ds_2 c(6), ds_3 c(6), det n(1), k2 n(5,3), tipgr c(1), ;
	 vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17),;
	 ord n(1), date_ord d, lpu_ord n(6), recid_lpu c(7), fil_id n(6), IsPr L, vz l, mp c(1), n_kd n(3), mm c(1))
 ENDIF 

 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON c_i TAG c_i
 INDEX ON sn_pol FOR IsTlnValid() TAG sn_pol
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

 CREATE TABLE (Error) (f c(1), c_err c(3), rid i)
 INDEX FOR UPPER(f)='R' ON rid TAG rrid
 INDEX FOR UPPER(f)='S' ON rid TAG rid
 USE 

 CREATE TABLE (mError) ;
  (rid i autoinc, RecId i, cod n(6), k_u n(3), tip c(1), et c(1), ee c(1), usr c(6), d_exp d,;
   e_cod n(6), e_ku n(3), e_tip c(1), err_mee c(3), osn230 c(5), e_period c(6),  ;
   koeff n(4,2), straf n(4,2), docexp c(7), s_all n(11,2), s_1 n(11,2), s_2 n(11,2), impdata d,;
   subet n(1), reason c(1), n_akt n(6), t_akt c(2), d_edit d)
 INDEX ON rid TAG rid 
 INDEX ON RecId TAG recid
 *INDEX ON PADL(recid,6,'0')+et TAG id_et
 *INDEX ON PADL(recid,6,'0')+et+LEFT(err_mee,2) TAG unik
 INDEX ON PADL(recid,6,'0')+et+docexp+reason TAG id_et
 INDEX ON PADL(recid,6,'0')+et+docexp+reason+LEFT(err_mee,2) TAG unik
 USE 

 CREATE CURSOR pazamb (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol
  
 CREATE CURSOR pazdst (sn_pol c(25))
 INDEX on sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR pazst (c_i c(30))
 INDEX ON c_i TAG c_i
 SET ORDER TO c_i

 CREATE CURSOR pazvmp (sn_pol c(25))
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
 fld_1 = AFIELDS(tabl_1, 'dfile') && �������� d-����� 
 fld_2 = AFIELDS(tabl_2, 'd_et')  && 1 ������� - ��������, 2 - ���,  3 - �����������, 4 - ����� ����� �������
 IF fld_1 == fld_2 && ���-�� ����� ���������!
  FieldsIdent = CompFields(dItem) && 0 - ���� �������, 1 - ������ ����������
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

* fld_1 = AFIELDS(tabl_1, 'hfile') && �������� h-����� 
* fld_2 = AFIELDS(tabl_2, 'h_et')  && 1 ������� - ��������, 2 - ���,  3 - �����������, 4 - ����� ����� �������
* IF fld_1 == fld_2 && ���-�� ����� ���������!
*  FieldsIdent = CompFields(hItem) && 0 - ���� �������, 1 - ������ ����������
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

 fld_1 = AFIELDS(tabl_1, 'nvfile') && �������� nv-�����
 fld_2 = AFIELDS(tabl_2, 'nv_et')
 IF 3=2
 IF fld_1 == fld_2 && ���-�� ����� ���������!
  FieldsIdent = CompFields(nvItem) && 0 - ���� �������, 1 - ������ ����������
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

* fld_1 = AFIELDS(tabl_1, 'nsfile') && �������� ns-�����
* fld_2 = AFIELDS(tabl_2, 'ns_et')
* IF fld_1 == fld_2 && ���-�� ����� ���������!
*  FieldsIdent = CompFields(nsItem) && 0 - ���� �������, 1 - ������ ����������
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

 fld_1 = AFIELDS(tabl_1, 'rfile') && �������� r-�����
 fld_2 = AFIELDS(tabl_2, 'r_et')
 IF fld_1 == fld_2 && ���-�� ����� ���������!
  FieldsIdent = CompFields(rItem) && 0 - ���� �������, 1 - ������ ����������
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

* fld_1 = AFIELDS(tabl_1, 'sfile') &&& �������� s-�����
* fld_2 = AFIELDS(tabl_2, 's_et')
* IF fld_1 == fld_2 && ���-�� ����� ���������!
*  FieldsIdent = CompFields(sItem) && 0 - ���� �������, 1 - ������ ����������
*  IF FieldsIdent==0
*   =CopyToTrash(m.InDir,2)
*   =ClDir()
*   RETURN .F.
*  ENDIF 
* ELSE 
*  =DiffFields(sItem)
*  =CopyToTrash(m.InDir,2)
*  =ClDir()
*  RETURN .F.
* ENDIF 

 m.SCompResult = .T.
 fld_1 = AFIELDS(tabl_1, 'sfile') &&& �������� s-�����
 fld_2 = AFIELDS(tabl_2, 's_et')
 IF fld_1 == fld_2 && ���-�� ����� ���������!
  FieldsIdent = CompFields(sItem) && 0 - ���� �������, 1 - ������ ����������
  IF FieldsIdent==0
   m.SCompResult = .F.
*   =CopyToTrash(m.InDir,2)
*   =ClDir()
*   RETURN .F.
  ENDIF 
 ELSE 
  m.SCompResult = .F.
*  =DiffFields(sItem)
*  =CopyToTrash(m.InDir,2)
*  =ClDir()
*  RETURN .F.
 ENDIF 

 IF m.SCompResult = .F.
 fld_1 = AFIELDS(tabl_1, 'sfile') &&& �������� s-�����
 fld_2 = AFIELDS(tabl_2, 'sv01_et')
 IF fld_1 == fld_2 && ���-�� ����� ���������!
  FieldsIdent = CompFields(sItem) && 0 - ���� �������, 1 - ������ ����������
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

 fld_1 = AFIELDS(tabl_1, 'sprfile') && �������� spr-�����
 fld_2 = AFIELDS(tabl_2, 'spr_et')
 IF fld_1 == fld_2 && ���-�� ����� ���������!
  FieldsIdent = CompFields(sprItem) && 0 - ���� �������, 1 - ������ ����������
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
 fld_1 = AFIELDS(tabl_1, 'hofile') && �������� ho-�����
 fld_2 = AFIELDS(tabl_2, 'spr_ho')
 IF fld_1 == fld_2 && ���-�� ����� ���������!
  FieldsIdent = CompFields(hoItem) && 0 - ���� �������, 1 - ������ ����������
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

RETURN .T. 

FUNCTION IsIPComplete
 DO CASE 
  CASE !UnzipGotoFileByName(dItem)
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [����������� ] + dItem + [ ����]
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
*   m.cerrmessage = [����������� ] + hItem + [ ����]
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
   m.cerrmessage = [����������� ] + nvItem + [ ����]
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
*   m.cerrmessage = [����������� ] + nsItem + [ ����]
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
   m.cerrmessage = [����������� ] + rItem + [ ����]
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
   m.cerrmessage = [����������� ] + sItem + [ ����]
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
*   m.cerrmessage = [����������� ] + dsItem + [ ����]
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
   m.cerrmessage = [����������� ] + sprItem + [ ����]
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
   CASE UPPER(READCFG) = 'MESSAGE'
    m.cmessage = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'RESENT-MESSAGE-ID'
    m.resmesid = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'SUBJECT'
    m.csubject = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    m.csubject1 = LEFT(m.csubject, RAT('#',m.csubject,2))   && ����� subject ��� ����������� ������� ���� ����������
    m.csubject2 = SUBSTR(m.csubject, RAT('#',m.csubject,1)) && ����� subject ��� ����������� ������� ���� ����������
   CASE UPPER(READCFG) = 'ATTACHMENT'
    m.attaches   = m.attaches + 1
    m.attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dattaches(m.attaches,1) = ALLTRIM(SUBSTR(m.attachment, 1, AT(" ",m.attachment)-1)) && �������� d-�����
    dattaches(m.attaches,2) = ALLTRIM(SUBSTR(m.attachment, AT(" ",m.attachment)+1))    && ����������� �������� �����
*   CASE UPPER(READCFG) = 'BODYPART'
*    m.bparts   = m.bparts + 1
*    m.bodypart = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
*    dbparts(m.bparts,1) = ALLTRIM(SUBSTR(m.bodypart, 1, AT(" ",m.bodypart)-1))
  ENDCASE
 ENDDO
RETURN 
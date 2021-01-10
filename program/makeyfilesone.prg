FUNCTION MakeYFilesOne
 PARAMETERS para1
 PRIVATE m.lcpath
  m.lcpath = para1
  hfile   = 'h'+STR(m.lpu_id,4)+'.'+mmy
  hsfile  = 'hs'+STR(m.lpu_id,4)+'.'+mmy
  dfile   = 'd'+STR(m.lpu_id,4)+'.'+mmy
  nvfile  = 'nv'+STR(m.lpu_id,4)+'.'+mmy
  nsfile  = 'ns'+STR(m.lpu_id,4)+'.'+mmy
  sprfile = 'spr' + STR(m.lpu_id,4) + '.' + mmy
  ryfile  = 'R'+m.qcod+'Y.'+mmy
  syfile  = 'S'+m.qcod+'Y.'+mmy
  eyfile  = 'CTRL'+m.qcod+'Y.'+mmy
  hofile  = 'ho'+m.qcod+'.'+mmy

  IF !fso.FileExists(lcpath + '\people.dbf') OR ;
     !fso.FileExists(lcpath + '\talon.dbf') OR ;
     !fso.FileExists(lcpath + '\e'+m.mcod+'.dbf')
   RETURN 
  ENDIF 
  
  fso.CopyFile(pTempl+'\rqqy.mmy', lcpath+'\'+ryfile, .t.)
  fso.CopyFile(pTempl+'\sqqy.mmy', lcpath+'\'+syfile, .t.)
  fso.CopyFile(pTempl+'\ctrlqqy.mmy', lcpath+'\'+eyfile, .t.)
  
  CREATE CURSOR curdbls (sn_pol c(25), k_u n(2))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  
  =OpenFile(lcpath+'\'+ryfile, "ryfile", "excl")
  SELECT ryfile 
  INDEX ON sn_pol TAG sn_pol
  SET ORDER TO sn_pol IN ryfile
  =OpenFile(lcpath+'\'+syfile, "syfile", "share")
  
  =OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\people', 'people', 'share', 'recid')
  =OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\talon', 'talon', 'share', 'recid')
  =OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'rerror', 'share')
  
  =OpenFile(lcpath+'\'+eyfile, "ctrl", "share")

  SELECT rerror
  SET RELATION TO RId INTO People
  SET RELATION TO RId INTO Talon ADDITIVE 

  SCAN FOR !DELETED()
   m.er_f = f
   IF m.er_f == 'R'
    m.recid = People.recid_lpu
   ELSE 
    m.recid = Talon.recid_lpu
   ENDIF 
   m.er_c = c_err
  
   INSERT INTO Ctrl FROM MEMVAR 
  
  ENDSCAN 
  SET ORDER TO rrid
  SET RELATION OFF INTO Talon
  SET RELATION OFF INTO People 
  
  USE IN ctrl
  
  SELECT people 
  SET ORDER TO 
  SELECT talon 
  SET ORDER TO 
 
  =OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'serror', 'share', 'rid', 'again')

  CREATE CURSOR deftln (sn_pol c(25))
  INDEX on sn_pol TAG sn_pol
  SET ORDER TO sn_pol
  
  m.IsStomat   = IIF(SUBSTR(m.mcod,3,2)='07', .T., .F.)
  m.IsIskl     = IIF(INLIST(m.lpu_id, 1912, 1940, 2049), .T., .F.)

  SELECT people
  SET RELATION TO RecId INTO rerror
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   m.recno = RECNO()
   *m.recid = PADL(m.recno,6,'0')
   m.recid = recid_lpu
   
   m.tip_p     = m.tipp
   
   *IF rerror.c_err='PNA' AND m.qcod='S7' && 'PNA'
   * m.er_c      = ''
   * m.refreason = ''
   * m.osn230    = ''
   *ELSE 
    m.er_c      = rerror.c_err
    m.refreason = IIF(SEEK(LEFT(m.er_c,2), 'sookod'), sookod.refreason, '')
    m.osn230    = IIF(SEEK(LEFT(m.er_c,2), 'sookod'), sookod.osn230, '')
   *ENDIF 

   m.et        = 'A'
   m.et230     = 1 && Если МЭК!
   m.oplata    = IIF(EMPTY(m.er_c), 1, 2) && 1-полная оплата, 2-полный отказ, 3-частичное снятие

   m.prmcod    = prmcod
   m.prik      = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
   m.lpu_prik  = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)

   m.prmcods   = prmcods
   m.priks     = IIF(SEEK(m.prmcods, 'sprlpu'), sprlpu.lpu_id, 0)

   IF !SEEK(m.sn_pol, 'ryfile')
    INSERT INTO ryfile FROM MEMVAR 
   ELSE 
    IF !SEEK(m.sn_pol, 'curdbls')
     INSERT INTO curdbls (sn_pol, k_u) VALUES (m.sn_pol, 1)
    ELSE 
     m.o_k_u = curdbls.k_u
     m.n_k_u = m.o_k_u + 1
     UPDATE curdbls SET k_u=m.n_k_u
    ENDIF 
    m.nIsDoubles = m.nIsDoubles + 1
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO rerror
  
  m.sum_st2 = 0
  
  SELECT people
  SET ORDER TO sn_pol
  SELECT talon
  SET RELATION TO recid INTO serror
  SET RELATION TO sn_pol INTO people ADDITIVE 
  SCAN  
   SCATTER MEMVAR 
   m.iotd = m.otd
   m.recno = RECNO()
   *m.recid = PADL(m.recno,6,'0')
   m.recid = recid_lpu

   m.er_c      = serror.c_err
   m.refreason = IIF(SEEK(LEFT(m.er_c,2), 'sookod'), sookod.refreason, '')
   m.osn230    = IIF(SEEK(LEFT(m.er_c,2), 'sookod'), sookod.osn230, '')
   m.et        = 'A'
   m.et230     = 1 && Если МЭК!

   m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod) OR IsEko(m.cod)), .T., .F.)
   
   IF !EMPTY(m.er_c)
    IF !SEEK(m.sn_pol, 'deftln')
     INSERT INTO deftln FROM MEMVAR 
    ENDIF 
   ENDIF 

   m.sum_st2 = m.sum_st2 + IIF(EMPTY(m.er_c), m.s_all, 0)
  
   m.prmcod = people.prmcod
   m.prik   = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
   m.prmcods= people.prmcods
   m.priks  = IIF(SEEK(m.prmcods, 'sprlpu'), sprlpu.lpu_id, 0)
   
   m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)
   m.lpu_ord = IIF(!EMPTY(FIELD('lpu_ord')), lpu_ord, 0)
   m.lpu_ord = ALLTRIM(STR(m.lpu_ord))
   m.paztip = TipOfPaz(m.mcod, m.prmcod) && 0 (не прикреплен),1 (прикреплен по месту обращения),2 (к пилоту),3 (не к пилоту)
   
   m.UslIskl      = IIF(FLOOR(m.cod/1000)=146, .T., .F.)
   m.IsStomatUsl  = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
   m.IsStomatUsl2 = IIF(INLIST(m.cod,1101,1102,101171,101172), .T., .F.)
   

   m.Typ   = Typ
   m.Mp    = Mp
   m.vz    = vz
   m.dop_r = dop_r
   
   m.f_type = ''

   * fp – при оплате из средств  подушевого  финансирования данной МО пациентов, прикрепленных к данной МО с ПФ;
   * up – при оплате из средств  подушевого  финансирования данной МО пациентов, прикрепленных к другим МО с ПФ;	
   * vz– при возмещении средств за МП, оказанную прикрепленным к другим МО (взаимозачеты)
   * fh—оплата по тарифу раздела «Дополнительные услуги» МО с ПФ
   * ft – при оплате по тарифу
   
   * Заполнение обязательно за исключением случаев стационарной медицинской помощи (МС, ВМП, раздел 99 КСГ).
   * Пояснение: Для МО, не включенных в подушевое финансирование, медицинские услуги реестра (за искл. раздела 99) 
   	* отмечаются  кодом «ft»

   DO CASE 
    CASE m.Typ = '0' && неприкрепленные
     DO CASE 
      CASE INLIST(m.Mp,'p','s')
       m.f_type = 'fp'
      CASE INLIST(m.Mp,'4','8')
       m.f_type = 'fh'
      OTHERWISE 
     ENDCASE 

    CASE m.Typ = '1' && свои
     DO CASE 
      CASE INLIST(m.Mp,'p','s')
       m.f_type = 'fp'
      CASE INLIST(m.Mp,'4','8')
       m.f_type = 'fh'
      OTHERWISE 
     ENDCASE 

    CASE m.Typ = '2' && чужие 
     DO CASE 
      CASE INLIST(m.Mp,'p','s') AND vz=0
       m.f_type = 'up'
      CASE INLIST(m.Mp,'p','s') AND vz>0
       m.f_type = 'vz'
      CASE INLIST(m.Mp,'4','8')
       m.f_type = 'fh'
      OTHERWISE 
     ENDCASE 
    OTHERWISE 
   ENDCASE 
   
   IF !m.IsPilot AND !m.IsPilotS
    IF IsKDS(m.cod)
     m.f_type=' '
    ELSE 
     m.f_type='ft'
    ENDIF 
   ENDIF 
   
   IF INLIST(m.prcell, '306','606')
    m.f_type = 'ft'
   ENDIF 

   INSERT INTO syfile FROM MEMVAR 
  
  ENDSCAN 
  SET RELATION OFF INTO serror
  SET RELATION OFF INTO people
  USE 

  SELECT people
  SET ORDER TO recid
  SET RELATION TO sn_pol INTO deftln
  SELECT rerror
  SET RELATION TO rid INTO people ADDITIVE 
  SCAN 
   IF EMPTY(deftln.sn_pol)
    DELETE 
   ENDIF 
  ENDSCAN 
  SET RELATION OFF INTO rerror
  USE 

  USE IN people 
  USE IN serror
  USE IN syfile
  SELECT ryfile
  SET ORDER TO 
  DELETE TAG ALL 
  SET RELATION TO sn_pol INTO deftln
  SCAN
   IF !EMPTY(er_c)
    IF EMPTY(deftln.sn_pol)
     REPLACE er_c WITH '', refreason WITH '', oplata WITH 1
    ENDIF 
   ENDIF  
  ENDSCAN 
  SET RELATION OFF INTO deftln
  USE 
  USE IN deftln
  
  IF m.sum_st1 != m.sum_st2
   *MESSAGEBOX(''+CHR(13)+CHR(10)+;
    'ВНИМАНИЕ!'+CHR(13)+CHR(10)+;
    'СУММА, РАСЧИТАННАЯ КАК РАЗНИЦА МЕЖДУ ПРЕДСТАВЛЕННОЙ И ФЛК, СОСТАВЛЯЕТ'+CHR(13)+CHR(10)+;
    TRANSFORM(m.sum_st1,'999 999 999.99')+CHR(13)+CHR(10)+;
    'СУММА ПО ДАННЫМ ПЕРСОНИФИЦИРОВАННОЙ ОТЧЕТНОСТИ СОСТАВЛЯЕТ'+CHR(13)+CHR(10)+;
    TRANSFORM(m.sum_st2,'999 999 999.99')+CHR(13)+CHR(10), 0+48, m.mcod)
  ENDIF 
  
  IF m.nIsDoubles>0
   m.stroka = ''
   SELECT curdbls
   SCAN 
    m.sn_pol = sn_pol
    m.stroka = m.stroka + m.sn_pol + CHR(13)+CHR(10)
   ENDSCAN 
  ENDIF 
 
  IF m.nIsDoubles>0
   MESSAGEBOX('ОБНАРУЖЕНЫ ДУБЛИ В РЕГИСТРЕ!'+CHR(13)+CHR(10)+;
   m.stroka+'ПЕРСОТЧЕТ НЕВЕРЕН!',0+64,m.mcod)
  ENDIF 
  
  USE IN curdbls

  ZipPath = lcPath
  ZipName = 'D'+m.qcod+STR(m.lpu_id,4)+'.zip'
  MmyName = 'D'+m.qcod+STR(m.lpu_id,4)+'.'+mmy

  IF fso.FileExists(lcpath+'\'+ZipName)
   fso.DeleteFile(lcpath+'\'+ZipName)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+MmyName)
   fso.DeleteFile(lcpath+'\'+MmyName)
  ENDIF 

  SET DEFAULT TO (lcpath)
  
  bfile='b'+m.mcod+'.'+mmy
  IF fso.FileExists(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+bfile)
   UnZipOpen(pbase+'\'+gcPeriod+'\'+m.mcod+'\'+bfile)

   UnzipGotoFileByName(dfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(hfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(hsfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(nvfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(nsfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(sprfile)
   UnzipFile(lcPath+'\')

   UnzipGotoFileByName(hofile)
   UnzipFile(lcPath+'\')

   SLItem    = 'ONK_SL' + m.qcod + '.' + m.mmy
   USLItem   = 'ONK_USL' + m.qcod + '.' + m.mmy
   CONSItem  = 'ONK_CONS' + m.qcod + '.' + m.mmy
   LSItem    = 'ONK_LS' + m.qcod + '.' + m.mmy
   NAPRItem  = 'ONK_NAPR_V_OUT' + m.qcod + '.' + m.mmy
   DIAGItem  = 'ONK_DIAG' + m.qcod + '.' + m.mmy
   PROTItem  = 'ONK_PROT' + m.qcod + '.' + m.mmy
   
   CVItem  = 'CV_LS' + m.qcod + '.' + m.mmy

   IF UnzipGotoFileByName(SLItem)
    UnzipFile(lcPath+'\')
   ENDIF 
   IF UnzipGotoFileByName(USLItem)
    UnzipFile(lcPath+'\')
   ENDIF 
   IF UnzipGotoFileByName(CONSItem)
    UnzipFile(lcPath+'\')
   ENDIF 
   IF UnzipGotoFileByName(LSItem)
    UnzipFile(lcPath+'\')
   ENDIF 
   IF UnzipGotoFileByName(NAPRItem)
    UnzipFile(lcPath+'\')
   ENDIF 
   IF UnzipGotoFileByName(DIAGItem)
    UnzipFile(lcPath+'\')
   ENDIF 
   IF UnzipGotoFileByName(PROTItem)
    UnzipFile(lcPath+'\')
   ENDIF 
   IF UnzipGotoFileByName(CVItem)
    UnzipFile(lcPath+'\')
   ENDIF 

   UnZipClose()

  ENDIF 
  
  ZipOpen(MmyName, lcPath+'\')
  IF fso.FileExists(lcpath+'\'+ryfile)
   ZipFile(ryfile, .T.)
   fso.DeleteFile(lcpath+'\'+ryfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+syfile)
   ZipFile(syfile, .T.)
   fso.DeleteFile(lcpath+'\'+syfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+eyfile)
   ZipFile(eyfile, .T.)
   fso.DeleteFile(lcpath+'\'+eyfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+dfile)
   ZipFile(dfile, .T.)
   fso.DeleteFile(lcpath+'\'+dfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+nvfile)
   ZipFile(nvfile, .T.)
   fso.DeleteFile(lcpath+'\'+nvfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+nsfile)
   ZipFile(nsfile, .T.)
   fso.DeleteFile(lcpath+'\'+nsfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+sprfile)
   ZipFile(sprfile, .T.)
   fso.DeleteFile(lcpath+'\'+sprfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+hfile)
   ZipFile(hfile, .T.)
   fso.DeleteFile(lcpath+'\'+hfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+hsfile)
   ZipFile(hsfile, .T.)
   fso.DeleteFile(lcpath+'\'+hsfile)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+hofile)
   IF OpenFile(lcpath+'\'+hofile, 'hoo', 'excl')>0
    IF USED('')
     USE IN hoo
    ENDIF 
   ELSE 
    ALTER TABLE hoo ALTER COLUMN c_i c(30)
    USE IN hoo 
   ENDIF 
   ZipFile(hofile, .T.)
   fso.DeleteFile(lcpath+'\'+hofile)
  ENDIF 
  
  ** Онкология
  SLItem    = 'ONK_SL' + m.qcod + '.' + m.mmy
  USLItem   = 'ONK_USL' + m.qcod + '.' + m.mmy
  CONSItem  = 'ONK_CONS' + m.qcod + '.' + m.mmy
  LSItem    = 'ONK_LS' + m.qcod + '.' + m.mmy
  NAPRItem  = 'ONK_NAPR_V_OUT' + m.qcod + '.' + m.mmy
  DIAGItem  = 'ONK_DIAG' + m.qcod + '.' + m.mmy
  PROTItem  = 'ONK_PROT' + m.qcod + '.' + m.mmy
   
  CVItem  = 'CV_LS' + m.qcod + '.' + m.mmy

  IF fso.FileExists(lcpath+'\'+SLItem)
   ZipFile(SLItem, .T.)
   fso.DeleteFile(lcpath+'\'+SLItem)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+USLItem)
   ZipFile(USLItem, .T.)
   fso.DeleteFile(lcpath+'\'+USLItem)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+CONSItem)
   ZipFile(CONSItem, .T.)
   fso.DeleteFile(lcpath+'\'+CONSItem)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+LSItem)
   ZipFile(LSItem, .T.)
   fso.DeleteFile(lcpath+'\'+LSItem)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+NAPRItem)
   ZipFile(NAPRItem, .T.)
   fso.DeleteFile(lcpath+'\'+NAPRItem)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+DIAGItem)
   ZipFile(DIAGItem, .T.)
   fso.DeleteFile(lcpath+'\'+DIAGItem)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+PROTItem)
   ZipFile(PROTItem, .T.)
   fso.DeleteFile(lcpath+'\'+PROTItem)
  ENDIF 
  IF fso.FileExists(lcpath+'\'+CVItem)
   ZipFile(CVItem, .T.)
   fso.DeleteFile(lcpath+'\'+CVItem)
  ENDIF 
  ** Онкология

  ZipClose()
RETURN 
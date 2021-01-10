PROCEDURE MakeIGSM
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ'+CHR(13)+CHR(10)+;
  'ОТЧЕТ ИГС-М?'+CHR(13)+CHR(10),4+32,'ЭКСПЕРТИЗА')==7
  RETURN 
 ENDIF 
 
 m.igsmfolder = pMee + '\IGSM'
 IF !fso.FolderExists(m.igsmfolder)
  fso.CreateFolder(m.igsmfolder)
 ENDIF 

 IF !fso.FileExists(m.igsmfolder+'\catalog.dbf')
  CREATE TABLE &igsmfolder\catalog (RecId i AUTOINC NEXTVALUE 1 STEP 1, ddata date, cdata char(8))
  INDEX on recid TAG recid 
  USE IN catalog 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'mcod')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\mee2mgf', 'mee2mgf', 'shar', 'my_et')>0
  IF USED('mee2mgf')
   USE IN mee2mgf
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(m.igsmfolder+'\catalog', 'catalog', 'shar')>0
  IF USED('catalog')
   USE IN catalog
  ENDIF 
  IF USED('mee2mgf')
   USE IN mee2mgf
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT recid FROM catalog INTO CURSOR curr ORDER BY recid DESC 
 SELECT curr
 m.pr_name = ''
 SCAN 
  m.lastid = recid
  IF fso.FileExists(m.igsmfolder+'\'+'f'+PADL(m.lastid,6,'0')+'.dbf')
   m.pr_name = 'f'+PADL(m.lastid,6,'0')
   EXIT 
  ENDIF   
 ENDSCAN
 USE 
 
 IF !EMPTY(m.pr_name) 
  =OpenFile(m.igsmfolder+'\'+m.pr_name, 'prfile', 'shar', 'unik')
 ENDIF 
 
 =OpenFile(pTempl+'\igsm', 'igsm', 'shar')
 SELECT * FROM igsm INTO CURSOR wrkcurs READWRITE 
 USE IN igsm 

 FOR nperiod=18 TO 0 STEP -1 
  m.expperiod = STR(YEAR(GOMONTH(tdat1,-nperiod)),4)+PADL(MONTH(GOMONTH(tdat1,-nperiod)),2,'0')
  IF !fso.FolderExists(pbase+'\'+expperiod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+expperiod+'\aisoms.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+expperiod+'\aisoms', 'aisoms', 'shar')>0
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+expperiod+'\'+'nsi'+'\tarifn', 'tarif', 'shared', 'cod')>0
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP  
  ENDIF 
  IF OpenFile(pbase+'\'+expperiod+'\'+'nsi'+'\sookodxx', 'sookod', 'shar', 'er_c')>0
   IF USED('sookod')
    USE IN sookod
   ENDIF 
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP  
  ENDIF 
  IF OpenFile(pbase+'\'+expperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
   IF USED('sprlpu')
    USE IN sprlpu
   ENDIF 
   IF USED('sookod')
    USE IN sookod
   ENDIF 
   IF USED('tarif')
    USE IN tarif
   ENDIF 
   IF USED('aisoms')
    USE IN aisoms
   ENDIF 
   LOOP  
  ENDIF 
  m.lIsLpuTpn=.f.
  IF fso.FileExists(pbase+'\'+expperiod+'\'+'nsi'+'\lputpn.dbf')
   m.lIsLpuTpn=.t.
  ENDIF 
  IF m.lIsLpuTpn=.t.
   IF OpenFile(pbase+'\'+expperiod+'\'+'nsi'+'\lputpn', 'lputpn', 'shar', 'lpu_id')>0
    IF USED('lputpn')
     USE IN lputpn
    ENDIF 
    IF USED('sprlpu')
     USE IN sprlpu
    ENDIF 
    IF USED('sookod')
     USE IN sookod
    ENDIF 
    IF USED('tarif')
     USE IN tarif
    ENDIF 
    IF USED('aisoms')
     USE IN aisoms
    ENDIF 
    LOOP  
   ENDIF 
  ENDIF 
  
  m.period   = m.expperiod

  SELECT aisoms
  WAIT "ОБРАБОТКА ПЕРИОДА "+expperiod+" ..." WINDOW NOWAIT 
  SCAN 
   m.mcod = mcod
   m.IsPilot = IIF(SEEK(m.mcod, 'pilot'), .t., .f.)
   m.lpuid = lpuid
   m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
   m.IsTpn = .f.
   IF m.lIsLpuTpn
    m.IsTpn = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
   ENDIF 
   m.d_schet = TTOD(recieved)

   IF !fso.FolderExists(pbase+'\'+expperiod+'\'+m.mcod)
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+expperiod+'\'+m.mcod+'\talon.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+expperiod+'\'+m.mcod+'\people.dbf')
    LOOP 
   ENDIF 
   IF !fso.FileExists(pbase+'\'+expperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+expperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+expperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('people')
     USE IN people
    ENDIF 
    LOOP 
   ENDIF 
   IF OpenFile(pbase+'\'+expperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar')>0
    IF USED('merror')
     USE IN merror
    ENDIF 
    IF USED('talon')
     USE IN talon
    ENDIF 
    IF USED('people')
     USE IN people
    ENDIF 
    LOOP 
   ENDIF 

   SELECT Talon
   SET RELATION TO sn_pol INTO people
   SELECT merror 
   SET RELATION TO recid INTO talon 
   
   SCAN 
    IF !EMPTY(err_mee) AND e_period = m.gcperiod
     SCATTER MEMVAR 
     
     m.vmp = 0 
     DO CASE 
      CASE IsPlk(m.cod)
       m.vmp = 2
      CASE IsDst(m.cod)
       m.vmp = 4
      CASE IsGsp(m.cod)
       m.vmp = 3
      CASE Is02(m.cod)
       m.vmp = 1
     ENDCASE 
     
     m.sn_pol = talon.sn_pol
     m.c_i    = talon.c_i
     m.ds     = talon.ds
     m.d_u    = talon.d_u
     m.pcod   = talon.pcod
     m.d_type = talon.d_type

     m.lpu_id   = m.lpuid
     m.lpu_name = IIF(SEEK(m.lpu_id, 'sprlpu', 'lpu_id'), sprlpu.fullname, '')
     m.fil_id   = talon.fil_id
     m.IsFilTpn = .f.
     IF m.IsTpn
      m.IsFilTpn = IIF(SEEK(m.fil_id, 'lputpn', 'fil_id'), .t., .f.)
     ENDIF 
     m.recid    = PADL(recid,6,'0')
     m.iotd     = talon.otd
*     m.period   = m.gcperiod
     m.period_e = expperiod
     m.s_opl    = talon.s_all
     m.er_c     = err_mee
     m.osn230   = IIF(SEEK(LEFT(UPPER(m.er_c),2), 'sookod'), sookod.osn230, '0.0.0.')
     m.et       = IIF(EMPTY(m.et), '2', m.et)
     m.evid     = IIF(INLIST(m.et,'2','3','7','8'), '3', '4')
     m.tip_e    = e_tip
     m.cod_e    = e_cod
     m.k_u_e    = e_ku
     m.ns_all   = 0
     m.delta    = 0
     
     m.s_sank = s_2
     
     m.prmcod = people.prmcod
     m.lpu_prik = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
     m.prmcods = IIF(!EMPTY(FIELD('prmcods', 'people')), people.prmcods, '')
     m.lpu_priks = IIF(SEEK(m.prmcods, 'sprlpu'), sprlpu.lpu_id, 0)
     
     m.expert = docexp
     m.smoexp = m.usr

     IF m.koeff <= 0 && старый механизм

      IF EMPTY(e_cod) AND EMPTY(e_tip) AND EMPTY(e_ku)
       m.delta = s_all
      ENDIF  && Полное снятие!
      IF (!EMPTY(e_cod) AND cod != e_cod) OR (!EMPTY(e_ku) AND k_u != e_ku) OR (!EMPTY(e_tip) AND e_tip != tip)
       m.ns_all = fsumm(e_cod, e_tip, e_ku, m.IsVed, GOMONTH(tdat1,-nperiod))
       m.delta = s_all - m.ns_all
      ENDIF && Частичное снятие
     
     ELSE 

      m.delta = ROUND(talon.s_all * m.koeff,2)

     ENDIF 

     m.s_opl_e = m.delta
     m.s_opl_e = s_1
     
     IF err_mee='W0'
      m.er_c     = '99'
      m.osn230   = '0.0.0.'

      m.ds_e     = ''
      m.tip_e    = ''
      m.cod_e    = 0
      m.codnom_e = ''
      m.k_u_e    = 0
      m.s_opl_e  = 0
      
*      m.s_sank   = 0
     ENDIF 
    
     m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)
     m.lpu_ord = IIF(!EMPTY(FIELD('lpu_ord')), lpu_ord, 0)
     m.paztip = TipOfPaz(m.mcod, m.prmcod) && 0 (не прикреплен),1 (прикреплен по месту обращения),2 (к пилоту),3 (не к пилоту)

     IF IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod) OR IsEKO(m.cod) OR IsPat(m.cod)
      m.f_type = 'ft'
      IF IsPat(m.cod) OR IsEKO(m.cod)
       m.f_type = 'fh'
      ENDIF 
     ELSE 
      DO CASE 
       CASE TipOfPr(m.mcod, m.prmcod) = 0 && неприкреплен
        m.f_type = 'fp'
        IF m.IsFilTpn = .T.
         IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
          m.f_type = 'fh'
         ENDIF 
        ENDIF 
        CASE TipOfPr(m.mcod, m.prmcod) = 2 && к пилоту
        IF !EMPTY(m.lpu_ord) OR (EMPTY(m.lpu_ord) AND (m.lIs02=.T. OR INLIST(SUBSTR(talon.otd,2,2),'08','92')))
         m.f_type = 'vz'
        ELSE 
         m.f_type = 'fp'
        ENDIF 
        IF m.IsFilTpn = .T.
         m.fil_id = fil_id
         IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
          m.f_type = 'fh'
         ENDIF 
        ENDIF 
        CASE TipOfPr(m.mcod, m.prmcod) = 3 && свой
        m.f_type = 'fp'
        IF m.IsFilTpn = .T.
         m.fil_id = fil_id
         IF !SEEK(m.fil_id, 'lputpn', 'fil_id')
          m.f_type = 'fh'
         ENDIF 
        ENDIF 
        OTHERWISE 
        m.f_type = ''
      ENDCASE 
     ENDIF 
   
     IF !m.IsPilot
      IF IsKDS(m.cod)
       m.f_type=' '
      ELSE 
       m.f_type='ft'
      ENDIF 
     ENDIF 

     m.et_old = m.et
*     m.et = IIF(SEEK(m.et, 'mee2mgf'), mee2mgf.mgf_et, m.et)

     m.extip = IIF(INLIST(m.et_old,'2','3','7'),'1','2')
     m.videxp = IIF(INLIST(m.et_old,'2','4','6','7'),'1','2') && 1 - плановая, в т.ч. тематическая.
     m.podtip = IIF(INLIST(m.et_old,'2','3','7'),'0','1') && это неправильно! сделано временно!
*     e_period+period+mcod+STR(codexp,1)+docexp

     m.vvir = m.gcperiod+m.period_e+m.mcod+m.et_old+m.docexp

     m.act = n_akt
     m.d_a = d_akt
     
     IF USED('prfile')
      m.unik = STR(m.lpu_id,6)+m.recid
      IF !SEEK(m.unik, 'prfile')
       INSERT INTO wrkcurs FROM MEMVAR 
      ENDIF 
     ELSE 
      INSERT INTO wrkcurs FROM MEMVAR 
     ENDIF 
     
     m.et = m.et_old
    ENDIF 

   ENDSCAN 
   SET RELATION OFF INTO talon
   SELECT talon 
   SET RELATION OFF INTO people
   USE IN talon 
   USE IN people
   USE IN merror


   SELECT aisoms
   
  ENDSCAN 
  
  USE IN aisoms
  USE IN sookod
  USE IN tarif
  USE IN sprlpu
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 
  WAIT CLEAR 
  
 NEXT 

 USE IN pilot
 USE IN mee2mgf
 
 IF RECCOUNT('wrkcurs')>0
  
  INSERT INTO catalog (ddata,cdata) VALUES (DATE(), STRTRAN(DTOC(DATE()),'.',''))
  m.id    = GETAUTOINCVALUE()
  USE IN catalog 
  m.fname = 'f'+PADL(m.id,6,'0')
 
  fso.CopyFile(pTempl+'\igsm.dbf', m.igsmfolder+'\'+m.fname+'.dbf')
 
  =OpenFile(m.igsmfolder+'\'+m.fname, 'mefile', 'excl')>0
  SELECT mefile
  INDEX on STR(lpu_id,6)+recid TAG unik
  
  SELECT wrkcurs
  SCAN 
   SCATTER MEMVAR 
   RELEASE nrec 
   INSERT INTO mefile FROM MEMVAR 
  ENDSCAN 
  USE 
 
  LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
  m.dotname = 'igsm.xls'
  m.docname = m.igsmfolder+'\'+m.fname
  m.lcTmpName = pTempl+'\'+m.dotname
  m.lcRepName = m.docname+'.xls'
  m.IsVisible = .T.

  m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
  
  USE IN mefile

 ELSE 
  USE IN catalog 
  IF USED('prfile')
   USE IN prfile
  ENDIF 
  MESSAGEBOX('НОВЫХ ДАННЫХ НЕ ОБНАРУЖЕНО!',0+64,'')
 ENDIF 
 
RETURN
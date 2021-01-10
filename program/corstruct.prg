PROCEDURE CorStruct

 DO CASE 
  CASE m.ffoms = 77011
   m.LoadPeriod={01.02.2012}+5*365
  CASE m.ffoms = 77002
   m.LoadPeriod={15.02.2012}+5*365 && 13.02.2017
  CASE m.ffoms = 77008
   m.LoadPeriod={17.02.2012}+5*365 && 15.02.2017
  CASE m.ffoms = 77013
   m.LoadPeriod={15.03.2012}+5*365 && 14.03.2017
  CASE m.ffoms = 77012
   m.LoadPeriod={17.03.2012}+5*365 && 16.03.2017
  OTHERWISE 
   m.LoadPeriod={17.03.2012}+5*365 && 16.03.2017

 ENDCASE 

 IF DATE()>m.LoadPeriod
*  =ChkDirsBrief()
 ENDIF 
 
 IF MESSAGEBOX('ВЫ ХОТИТЕ ПРОВЕСТИ '+CHR(13)+CHR(10)+;
               'КОРРЕКТИРОВКУ СТРУКТУРЫ БД?!'+CHR(13)+CHR(10)+;
               '',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('ВЫ АБСОЛЮТНО УВЕРЕНЫ В СВОИХ ДЕЙСТВИЯХ?',4+48,'') != 6
  RETURN 
 ENDIF 

 ppriod = STR(tYear,4)+PADL(tMonth,2,'0')
 spriod = PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)

 ppdir  = pbase+'\'+ppriod
 IF !fso.FolderExists(ppdir)
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+ppdir,0+16,'')  
  RETURN
 ENDIF 
 
 aisfile = ppdir+'\AisOms'
 IF !fso.FileExists(aisfile+'.dbf')
  MESSAGEBOX('ОТСУТСТВУЕТ ФАЙЛ '+aisfile,0+16,'')  
  RETURN
 ENDIF 
 
 IF OpenFile(aisfile, 'AisOms', 'shared', 'mcod')>0
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\UsrLpu', "UsrLpu", "shar", "mcod") > 0
  USE IN aisoms
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\tarifn', "tarif", "shar", "cod") > 0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  USE IN aisoms
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\profus', "profus", "shar", "cod") > 0
  IF USED('profus')
   USE IN profus
  ENDIF 
  USE IN tarif
  USE IN aisoms
  RETURN
 ENDIF 
 IF OpenFile(pcommon+'\dspcodes', "dspcodes", "shar", "cod") > 0
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  USE IN profus
  USE IN tarif
  USE IN aisoms
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\pilot', "pilot", "shar", "mcod") > 0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  USE IN dspcodes
  USE IN profus
  USE IN tarif
  USE IN aisoms
  RETURN
 ENDIF 
 
 *IF fso.FileExists(pbase+'\'+gcperiod+'\dsp.dbf')
 * IF OpenFile(pbase+'\'+gcperiod+'\dsp', "dsp", "excl") > 0
 *  IF USED('dsp')
 *   USE IN dsp
 *  ENDIF 
 * ELSE 
 *  SELECT dsp
 *  IF FIELD('tip')!='TIP'
 *   ALTER TABLE dsp ADD COLUMN Tip n(1)
 *  ENDIF 
 *  SET RELATION TO cod INTO dspcodes
 *   REPLACE ALL tip WITH dspcodes.tip
 *  SET RELATION OFF INTO dspcodes
 *  USE IN dspcodes 
 *  INDEX on mcod+sn_pol+PADL(tip,1,"0") TAG NewExpTag
 *  USE IN dsp 
 * ENDIF 
 *ENDIF 

 SELECT AisOms
 SCAN
  m.mcod = mcod
  m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
  m.lpuid = STR(lpuid,4)
  m.nvfile = 'nv'+m.lpuid
  m.IsPilot  = IIF(SEEK(m.mcod, 'pilot'), .T., .F.)

  WAIT m.mcod WINDOW NOWAIT 

  IF !fso.FolderExists(ppdir+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppdir+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppdir+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\people', 'people', 'excl')>0
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\talon', 'talon', 'excl')>0
   USE IN People
   IF USED('talon')
    USE IN talon
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\doctor', 'doctor', 'excl')>0
   USE IN People
   USE IN talon
   IF USED('doctor')
    USE IN doctor
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\m'+m.mcod, 'merror', 'excl')>0
   USE IN People
   USE IN talon
   USE IN doctor
   IF USED('merror')
    USE IN merror
   ENDIF 
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\e'+m.mcod, 'eerror', 'excl')>0
   USE IN merror
   USE IN People
   USE IN talon
   USE IN doctor
   IF USED('eerror')
    USE IN eerror
   ENDIF 
   LOOP 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\people.bak')
   fso.DeleteFile(ppdir+'\'+m.mcod+'\people.bak')
  ENDIF 
  IF fso.FileExists(ppdir+'\'+m.mcod+'\talon.bak')
   fso.DeleteFile(ppdir+'\'+m.mcod+'\talon.bak')
  ENDIF 
  IF fso.FileExists(ppdir+'\'+m.mcod+'\otdel.bak')
   fso.DeleteFile(ppdir+'\'+m.mcod+'\otdel.bak')
  ENDIF 
  IF fso.FileExists(ppdir+'\'+m.mcod+'\doctor.bak')
   fso.DeleteFile(ppdir+'\'+m.mcod+'\doctor.bak')
  ENDIF 
  
  SELECT doctor
  IF FIELD('d_ser2')!=UPPER('d_ser2')
   ALTER table doctor ADD COLUMN d_ser2 d
   m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)
   m.bfile = 'b' + m.mcod + '.' + m.mmy
   m.nvItem   = 'NV'+m.lpuid+'.'+m.mmy
   IF fso.FileExists(ppdir+'\'+m.mcod+'\'+m.bfile)
    UnzipOpen(ppdir+'\'+m.mcod+'\'+m.bfile)
    UnzipGotoFileByName(nvItem)
    UnzipClose()
    IF fso.FileExists(ppdir+'\'+m.mcod+'\'+m.nvItem)
     IF OpenFile(ppdir+'\'+m.mcod+'\'+m.nvItem, 'nvitem', 'excl')<=0
      SELECT nvItem
      INDEX on pcod TAG pcod 
      SET ORDER TO pcod 
      SELECT doctor 
      SET RELATION TO pcod INTO nvItem
      REPLACE ALL d_ser2 WITH IIF(FIELD('d_ser2','nvItem')='D_SERV2', nvItem.d_ser2, {})
      SET RELATION OFF INTO nvItem
      SELECT nvItem
      DELETE TAG ALL 
      USE IN nvItem
      fso.DeleteFile(ppdir+'\'+m.mcod+'\'+m.nvItem)
     ENDIF 
    ENDIF 
   ENDIF 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\onk_ls'+m.qcod+'.dbf')
   IF OpenFile(ppdir+'\'+m.mcod+'\onk_ls'+m.qcod, 'onk', 'excl')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ELSE 
    SELECT onk
    IF FIELD('sqlid') != 'SQLID'
     ALTER TABLE onk ADD COLUMN sqlid i
    ENDIF 
    IF FIELD('sqldt') != 'SQLDT'
     ALTER TABLE onk ADD COLUMN sqldt t
    ENDIF 
    IF FIELD('s_all')!='S_ALL'
     ALTER TABLE onk ADD COLUMN s_all n(11,2)
    ENDIF 
    IF FIELD('oms')!='OMS'
     ALTER TABLE onk ADD COLUMN oms l
    ENDIF 
    USE IN onk
   ENDIF 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\onk_diag'+m.qcod+'.dbf')
   IF OpenFile(ppdir+'\'+m.mcod+'\onk_diag'+m.qcod, 'onk', 'excl')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ELSE 
    SELECT onk
    IF FIELD('sqlid') != 'SQLID'
     ALTER TABLE onk ADD COLUMN sqlid i
    ENDIF 
    IF FIELD('sqldt') != 'SQLDT'
     ALTER TABLE onk ADD COLUMN sqldt t
    ENDIF 
    USE IN onk
   ENDIF 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\onk_prot'+m.qcod+'.dbf')
   IF OpenFile(ppdir+'\'+m.mcod+'\onk_prot'+m.qcod, 'onk', 'excl')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ELSE 
    SELECT onk
    IF FIELD('sqlid') != 'SQLID'
     ALTER TABLE onk ADD COLUMN sqlid i
    ENDIF 
    IF FIELD('sqldt') != 'SQLDT'
     ALTER TABLE onk ADD COLUMN sqldt t
    ENDIF 
    USE IN onk
   ENDIF 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\onk_cons'+m.qcod+'.dbf')
   IF OpenFile(ppdir+'\'+m.mcod+'\onk_cons'+m.qcod, 'onk', 'excl')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ELSE 
    SELECT onk
    IF FIELD('sqlid') != 'SQLID'
     ALTER TABLE onk ADD COLUMN sqlid i
    ENDIF 
    IF FIELD('sqldt') != 'SQLDT'
     ALTER TABLE onk ADD COLUMN sqldt t
    ENDIF 
    USE IN onk
   ENDIF 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\onk_usl'+m.qcod+'.dbf')
   IF OpenFile(ppdir+'\'+m.mcod+'\onk_usl'+m.qcod, 'onk', 'excl')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ELSE 
    SELECT onk
    IF FIELD('sqlid') != 'SQLID'
     ALTER TABLE onk ADD COLUMN sqlid i
    ENDIF 
    IF FIELD('sqldt') != 'SQLDT'
     ALTER TABLE onk ADD COLUMN sqldt t
    ENDIF 
    USE IN onk
   ENDIF 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\onk_sl'+m.qcod+'.dbf')
   IF OpenFile(ppdir+'\'+m.mcod+'\onk_sl'+m.qcod, 'onk', 'excl')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ELSE 
    SELECT onk
    IF FIELD('sqlid') != 'SQLID'
     ALTER TABLE onk ADD COLUMN sqlid i
    ENDIF 
    IF FIELD('sqldt') != 'SQLDT'
     ALTER TABLE onk ADD COLUMN sqldt t
    ENDIF 
    USE IN onk
   ENDIF 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\onk_napr_v_out'+m.qcod+'.dbf')
   IF OpenFile(ppdir+'\'+m.mcod+'\onk_napr_v_out'+m.qcod, 'onk', 'excl')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ELSE 
    SELECT onk
    IF FIELD('sqlid') != 'SQLID'
     ALTER TABLE onk ADD COLUMN sqlid i
    ENDIF 
    IF FIELD('sqldt') != 'SQLDT'
     ALTER TABLE onk ADD COLUMN sqldt t
    ENDIF 
    USE IN onk
   ENDIF 
  ENDIF 

  IF fso.FileExists(ppdir+'\'+m.mcod+'\ho'+m.qcod+'.dbf')
   IF OpenFile(ppdir+'\'+m.mcod+'\ho'+m.qcod, 'ho', 'excl')>0
    IF USED('ho')
     USE IN ho
    ENDIF 
   ELSE 
    SELECT ho
    IF FSIZE('c_i')!=30
     ALTER table ho alter COLUMN c_i c(30)
    ENDIF 
    USE IN ho 
   ENDIF 
  ENDIF 

  SELECT eerror
  IF FIELD('detail')!='DETAIL'
   ALTER TABLE eerror ADD COLUMN detail c(1)
  ENDIF 
  IF FIELD('comment')!='COMMENT'
   ALTER TABLE eerror ADD COLUMN "comment" c(250)
  ENDIF 
  IF FIELD('sqlid') != 'SQLID'
   ALTER TABLE eerror ADD COLUMN sqlid i
  ENDIF 
  IF FIELD('sqldt') != 'SQLDT'
   ALTER TABLE eerror ADD COLUMN sqldt t
  ENDIF 
  IF FIELD('et') != 'ET'
   ALTER TABLE eerror ADD COLUMN et n(1)
   REPLACE ALL et WITH 1
  ENDIF 
  
  SELECT merror
  IF FIELD('subet')!='SUBET'
   ALTER TABLE merror ADD COLUMN SubEt n(1)
  ENDIF 
  IF FIELD('reason')!='REASON'
   ALTER TABLE merror ADD COLUMN reason c(1)
  ENDIF 
  IF FIELD('n_akt')!='N_AKT'
   ALTER TABLE merror ADD COLUMN n_akt c(15)
  ELSE 
   IF VARTYPE(n_akt) != 'C'
    ALTER TABLE merror drop COLUMN n_akt
    ALTER TABLE merror ADD COLUMN n_akt c(15)
   ENDIF 
  ENDIF 
  IF FIELD('d_akt')!='D_AKT'
   ALTER TABLE merror ADD COLUMN d_akt d
  ENDIF 
  IF FIELD('t_akt')!='T_AKT'
   ALTER TABLE merror ADD COLUMN t_akt c(2)
  ENDIF 
  IF FIELD('d_edit')!='D_EDIT'
   ALTER TABLE merror ADD COLUMN d_edit d
  ENDIF 
  
  SELECT talon 
  
  IF FIELD('nsif') != 'NSIF'
   ALTER TABLE Talon ADD COLUMN nsif n(1)
  ENDIF 

  IF FIELD('kd_fact') != 'KD_FACT'
   ALTER TABLE Talon ADD COLUMN kd_fact n(3)
  ENDIF 

  IF FSIZE('kd_fact')<>4
   ALTER table Talon alter COLUMN kd_fact n(4)
  ENDIF 

  IF FIELD('prcell') != 'PRCELL'
   ALTER TABLE Talon ADD COLUMN prcell c(3)
  ENDIF 

  IF FIELD('dop_r') != 'DOP_R'
   ALTER TABLE Talon ADD COLUMN dop_r n(2)
   SET ORDER to sn_pol IN people 
   SET RELATION TO sn_pol INTO people 
   SCAN 
   m.cod     = cod 
   m.ds      = ds 
   m.d_type  = d_type
   m.otd     = otd
   m.ord     = ord
   m.lpu_ord = lpu_ord
   m.dop_r   = 0
   IF !m.IsDental(m.cod, INT(VAL(m.lpuid)), m.mcod, m.ds)
    DO CASE 
     CASE IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)
      m.dop_r = 1
     CASE m.d_type='s'
      m.dop_r = 2
     CASE m.IsPilot AND INLIST(SUBSTR(m.otd,2,2),'08')
      m.dop_r = 3
     CASE INLIST(SUBSTR(m.otd,2,2),'70','73','90','93') AND IsStac(m.mcod)
      m.dop_r = 4
     CASE SUBSTR(m.otd,2,2)='01' AND IsStac(m.mcod)
      m.dop_r = 5
     CASE m.ord=7 AND m.lpu_ord=7665
      m.dop_r = 6
     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),29,129,49,149) AND people.mcod!=people.prmcod AND people.tip_p=3 && только для ничьих и чужих!
      m.dop_r = 7 && 29,129,49,149 раздел для госпитализированных чужих
     ** Добавлено 16.04.2019 по требованию Согаза
    ENDCASE 
   ELSE 
    DO CASE 
     CASE IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)
      m.dop_r = 8
     CASE SUBSTR(m.otd,2,2)='08'
      m.dop_r = 9
     CASE INLIST(SUBSTR(m.otd,2,2),'70','73','93') AND IsStac(m.mcod)
      m.dop_r = 10
    ENDCASE 
   ENDIF 
   REPLACE dop_r WITH m.dop_r
   ENDSCAN 
   SET RELATION OFF INTO people 
   SELECT people 
   SET ORDER to 
   SELECT talon 
  ENDIF 

  IF FIELD('vz') = 'VZ' AND VARTYPE(vz) != 'N'
   ALTER TABLE Talon drop COLUMN vz
   ALTER TABLE Talon ADD COLUMN vz n(1)
  ENDIF 
  IF FIELD('vz') != 'VZ'
   ALTER TABLE Talon ADD COLUMN vz n(1)
   SCAN 
    IF !(Mp='p' AND Typ='2' )
     LOOP 
    ENDIF 
    m.cod = cod
    m.lpu_ord = lpu_ord
    m.Is02    = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
    m.profil = profil
    m.otd = SUBSTR(otd,2,2)
    
    DO CASE 
     CASE m.lpu_ord>0 && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
      m.vz = 1
     CASE m.Is02 && vz=2, неотложная помощь (по реестру медицинских услуг)
      m.vz = 2
     CASE m.profil='100' AND INLIST(m.otd,'00','92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
      m.vz = 3
     CASE m.otd='08' && vz=4, услуги ЖК
      m.vz = 4
     CASE m.otd='91' && vz=5, услуги ЦЗ
      m.vz = 5
     OTHERWISE 
      m.vz = 0 && то, что должно попасть в up-файл
    ENDCASE 
    
    REPLACE vz WITH m.vz
    
   ENDSCAN 
  ENDIF 

  IF FIELD('sqlid') != 'SQLID'
   ALTER TABLE Talon ADD COLUMN sqlid i
  ENDIF 
  IF FIELD('sqldt') != 'SQLDT'
   ALTER TABLE Talon ADD COLUMN sqldt t
  ENDIF 
  
  
  * Удаляетя на основании статистика заполения файлов 01.06.2019
  IF FIELD('codnom') = 'CODNOM'
   ALTER TABLE Talon drop COLUMN codnom
  ENDIF 
  IF FIELD('napr_usl') = 'NAPR_USL'
   ALTER TABLE Talon drop COLUMN napr_usl
  ENDIF 
  IF FIELD('vid_vme') = 'VID_VME'
   ALTER TABLE Talon drop COLUMN vid_vme
  ENDIF 
  IF FIELD('tipgr') = 'TIPGR'
   ALTER TABLE Talon drop COLUMN tipgr
  ENDIF 

  *IF FIELD('mm') = 'MM'
  * ALTER TABLE Talon drop COLUMN mm
  *ENDIF 
  IF FIELD('f_type') = 'F_TYPE'
   ALTER TABLE Talon drop COLUMN f_type
  ENDIF 
  * Удаляетя на основании статистика заполения файлов 01.06.2019

  IF FIELD('napr_v_in') != UPPER('napr_v_in')
   ALTER TABLE Talon ADD COLUMN napr_v_in n(1)
  ENDIF 

  IF FIELD('ispr') != 'ISPR'
   ALTER TABLE Talon ADD COLUMN ispr l
  ENDIF 

  IF FSIZE('k_u')!=4
   ALTER TABLE talon ALTER COLUMN k_u n(4,0)
   DELETE TAG sn_pol
   INDEX on sn_pol TAG sn_pol
  ENDIF 

  IF FIELD('S_LEK') != 'S_LEK'
   ALTER TABLE Talon ADD COLUMN s_lek n(11,2)
  ENDIF 

  IF FIELD('TYP') != 'TYP'
   ALTER TABLE Talon ADD COLUMN typ c(1)
  ENDIF 

  IF FSIZE('c_i')!=30
   ALTER TABLE talon ALTER COLUMN c_i c(30)
   DELETE TAG sn_pol
   INDEX on sn_pol TAG sn_pol
  ENDIF 

  *IF FIELD('f_type')!='F_TYPE'
  * ALTER TABLE Talon ADD COLUMN f_type c(2)
  *ELSE 
  * IF FSIZE('f_type')!=2
  *  ALTER TABLE talon ALTER COLUMN f_type c(2)
  * ENDIF 
  *ENDIF 

  IF FIELD('lpu_ord')!='LPU_ORD'
   ALTER TABLE talon ADD COLUMN lpu_ord n(6)
  ENDIF 
  IF FIELD('date_ord')!='DATE_ORD'
   ALTER TABLE talon ADD COLUMN date_ord d
  ENDIF 

  IF FIELD('n_kd')!='N_KD'
   WAIT "ДОБАВЛЕНИЕ N_KD..." WINDOW NOWAIT 
   ALTER TABLE talon ADD COLUMN n_kd n(3)
   SCAN 
    m.tip = tip 
    IF EMPTY(m.tip)
     LOOP 
    ENDIF 
    m.cod = cod 
    IF !SEEK(m.cod, 'tarif')
     LOOP 
    ENDIF 
    m.n_kd = tarif.n_kd
    REPLACE n_kd WITH m.n_kd
   ENDSCAN 
   WAIT CLEAR 
  ENDIF 
  IF FIELD('mp')!='MP'
   WAIT "ДОБАВЛЕНИЕ MP..." WINDOW NOWAIT 
   ALTER TABLE talon ADD COLUMN mp c(1)
   WAIT CLEAR 
  ENDIF 
  
  SELECT people

  DELETE TAG sn_pol
  INDEX on sn_pol TAG sn_pol

  IF FIELD('prmcods')!='PRMCODS'
   ALTER TABLE People ADD COLUMN prmcods c(7)
  ENDIF 
  IF FIELD('IsPr')!='ISPR'
   ALTER TABLE People ADD COLUMN IsPr L
  ENDIF 
  IF FIELD('s_all')!='S_ALL'
   ALTER TABLE People ADD COLUMN s_all n(11,2)
  ENDIF 
  IF FIELD('fil_id')!='FIL_ID'
   ALTER TABLE People ADD COLUMN fil_id n(6)
  ENDIF 
  IF FIELD('prmcod')!='PRMCOD'
   ALTER TABLE People ADD COLUMN prmcod c(7)
  ENDIF 
  IF FIELD('tipp')!='TIPP'
   ALTER TABLE People ADD COLUMN tipp c(1)
   SCAN 
    DO CASE 
     CASE IsEnp(sn_pol)
      REPLACE tipp WITH 'П'
     CASE IsKms(sn_pol)
      REPLACE tipp WITH 'С'
     CASE IsVs(sn_pol)
      REPLACE tipp WITH 'С'
     OTHERWISE 
      REPLACE tipp WITH 'С'
    ENDCASE 
   ENDSCAN 
  ENDIF 
*  USE 

  IF !fso.FileExists(ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf')
   IF fso.FileExists(ppdir+'\'+m.mcod+'\'+m.nvfile+'.'+spriod)
    fso.CopyFile(ppdir+'\'+m.mcod+'\'+m.nvfile+'.'+spriod, ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf')
    oSettings.CodePage(ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf', 866, .t.)
    IF OpenFile(ppdir+'\'+m.mcod+'\'+nvfile, 'nvfile', 'excl') == 0
     SELECT nvfile 
     INDEX ON pcod TAG pcod 
     USE 
    ENDIF 
   ENDIF 
  ELSE 
*   fso.DeleteFile(ppdir+'\'+m.mcod+'\'+m.nvfile+'.dbf')
*   fso.DeleteFile(ppdir+'\'+m.mcod+'\'+m.nvfile+'.cdx')
  ENDIF 

  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  IF USED('doctor')
   USE IN doctor
  ENDIF 
  IF USED('merror')
   USE IN merror
  ENDIF 
  IF USED('eerror')
   USE IN eerror
  ENDIF 

  SELECT aisoms

 ENDSCAN 

 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('usrlpu')
  USE IN UsrLpu
 ENDIF 
 IF USED('tarif')
  USE IN tarif
 ENDIF 
 IF USED('profus')
  USE IN profus
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 
 IF USED('pilot')
  USE IN pilot
 ENDIF 
 
 WAIT CLEAR 

 MESSAGEBOX('OK!', 0+64, '')

RETURN 
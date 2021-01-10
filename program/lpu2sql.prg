PROCEDURE Fact2sql

 IF MESSAGEBOX('ИМПОРТИРОВАТЬ БД В SQL?',4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 

 nHandl = SQLCONNECT("local")
 IF nHandl <= 0
  nHandl = SQLCONNECT("lpu", "sa", "admin")
 ENDIF 
 IF nHandl <= 0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'Cannot make connection')
  RETURN 
 ENDIF

 =SetSession()
 
 IF SQLEXEC(nHandl, "USE lpu") = -1
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'Cannot use lpu')
  m.lResult = .F.
 ENDIF 
 
 DO OpenBases
 
 m.d_beg = SECONDS()
 
 DO aisoms

 SELECT aisoms
 SCAN 
  m.mcod  = mcod
  m.lpuid = lpuid
  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   USE IN talon 
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf', 'err', 'shar', 'rid')>0
   USE IN talon 
   USE IN people
   IF USED('err')
    USE IN err
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  m.IsPilot  = IIF(SEEK(m.lpuid, 'pilot'), .T., .F.)
  m.IsPilotS = IIF(SEEK(m.lpuid, 'pilots'), .T., .F.)
  m.IsHorLpu = IIF(SEEK(m.lpuid, 'horlpu'), .T., .F.)
  m.IsHorLpuS = IIF(SEEK(m.lpuid, 'horlpus'), .T., .F.)

  m.IsStac  = IIF(VAL(SUBSTR(m.mcod,3,2))>40,.t.,.f.)
  
  WAIT m.mcod + '...' WINDOW NOWAIT 
  
  lnGoodRecs = 0
  SELECT talon 
  SET RELATION TO sn_pol INTO people
  SET RELATION TO cod INTO tarif ADDITIVE 
  SET RELATION TO recid INTO err ADDITIVE 
  SET RELATION TO SUBSTR(otd,2,2) INTO profot ADDITIVE 
  IF USED('onk_sl')
   SET RELATION TO recid_lpu INTO onk_sl
  ENDIF

  *DO FactServices
  
  SET RELATION OFF INTO people
  SET RELATION OFF INTO tarif
  SET RELATION OFF INTO err
  SET RELATION OFF INTO profot 
  IF USED('onk_sl')
   SET RELATION OFF INTO onk_sl
  ENDIF 
  USE IN err 
  
  *DO FactMEK
  
  DO FactCases
  
  DO FactConsiliums
  
  DO FactReferrals

  DO FactOncoServices

  DO FactDenials

  DO FactOncoDiag
  
  DO FactDrugs

  USE IN talon 
  USE IN people

  SELECT aisoms
  
 ENDSCAN 

 USE IN aisoms 
 USE IN tarif
 USE IN profot
 USE IN periods
 USE IN mkb_c
 USE IN pilot
 USE IN pilots
 USE IN horlpu
 USE IN horlpus
 IF USED('tarion')
  USE IN tarion 
 ENDIF 
 IF USED('medx')
  USE IN medx
 ENDIF 
 IF USED('medpack')
  USE IN medpack
 ENDIF 
 IF USED('mfc')
  USE IN mfc
 ENDIF 

 m.d_end = SECONDS()
 WAIT CLEAR 
 
 IF USED('transid')
  USE IN transid
 ENDIF 
 
 MESSAGEBOX("ВРЕМЯ ОБРАБОТКИ: "+TRANSFORM(m.d_end-m.d_beg,'999999999') +' сек.',0+64,'')

 *IF SQLEXEC(nHandl, "ALTER DATABASE kms SET MULTI_USER")==-1
 * MESSAGEBOX("БД KMS НЕ УДАЛОСЬ ПЕРЕВЕСТИ"+CHR(13)+CHR(10)+;
 * "В МНОГОПОЛЬЗОВАТЕЛЬСКИЙ РЕЖИМ!!", 0+64, "")
 *ENDIF 

 =SQLDISCONNECT(nHandl)

RETURN 


PROCEDURE Dims2sql

 IF MESSAGEBOX('ИМПОРТИРОВАТЬ РАЗМЕРНОСТИ В SQL?',4+32,'')=7
  RETURN 
 ENDIF 

 IF !fso.FolderExists(pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\NSI')
  RETURN 
 ENDIF 
 
 nHandl = SQLCONNECT("local")
 IF nHandl <= 0
  nHandl = SQLCONNECT("lpu", "sa", "admin")
 ENDIF 
 IF nHandl <= 0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'Cannot make connection')
  RETURN 
 ENDIF

 =SetSession()
 
 IF SQLEXEC(nHandl, "USE lpu") = -1
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'Cannot use lpu')
  m.lResult = .F.
 ENDIF 

 m.d_beg = SECONDS()
 
 IF OpenFile(pCommon+'\periods', 'periods', 'shar')>0
  IF USED('periods')
   USE IN periods
  ENDIF 
 ELSE 
  WAIT "periods..." WINDOW NOWAIT 
  SELECT periods
  SCAN 
  
   m.code = id
   m.name = period
   
   cmd01 = 'INSERT INTO dim.periods '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 


 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\hopff', 'hopff', 'shar', 'codho')>0
  IF USED('hopff')
   USE IN hopff
  ENDIF 
 ELSE 
  WAIT "hopff..." WINDOW NOWAIT 
  SELECT hopff 
  SCAN 
  
   SCATTER MEMVAR 
   
   m.code = codho
   
   cmd01 = 'INSERT INTO dim.Surgeries '
   cmd02 = '(code, name, ms)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name, ?m.ms)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
 ELSE 
  WAIT "tarif..." WINDOW NOWAIT 
  SELECT tarif 
  SCAN 
  
   SCATTER MEMVAR 
   
   m.code = cod
   
   cmd01 = 'INSERT INTO dim.Tarif '
   cmd02 = '(code, name, tpn, uet1, uet2, n_kd)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name, ?m.tpn, ?m.uet1, ?m.uet2, ?m.n_kd)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\hopff', 'hopff', 'shar')>0
  IF USED('hopff')
   USE IN hopff
  ENDIF 
 ELSE 
  WAIT "hopff..." WINDOW NOWAIT 
  SELECT hopff 
  SCAN 
  
   SCATTER MEMVAR 
   
   m.code = codho
   
   cmd01 = 'INSERT INTO dim.hopff '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\tarion', 'tarif', 'shar')>0
  IF USED('tarif')
   USE IN tarif
  ENDIF 
 ELSE 
  WAIT "tarion..." WINDOW NOWAIT 
  SELECT tarif
  SCAN 
  
   SCATTER MEMVAR 
   
   m.code        = m.cod
   m.p_mas_value = m.p_mas_valu
   
   cmd01 = 'INSERT INTO dim.TariOn '
   cmd02 = '(code, name , forlek, dozlp, mass_value, mass_unit, vol_value, vol_unit, pr_v_value, '
   cmd03 = 'pr_v_unit, p_mas_value, p_mas_unit, edizm)'
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name, ?m.forlek, ?m.dozlp, ?m.mass_value, ?m.mass_unit, ?m.vol_value, ?m.vol_unit, ?m.pr_v_value, '
   cmd08 = '?m.pr_v_unit, ?m.p_mas_value, ?m.p_mas_unit, ?m.edizm)'
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\medicament', 'med', 'shar')>0
  IF USED('med')
   USE IN med
  ENDIF 
 ELSE 
  WAIT "medicament..." WINDOW NOWAIT 
  SELECT med
  SCAN 
  
   SCATTER MEMVAR 
   
   cmd01 = 'INSERT INTO dim.Medicaments '
   cmd02 = '(dd_sid, dd_name, is_oms, mass_value, mass_unit, vol_value, vol_unit, gd_sid, gd_name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.dd_sid, ?m.dd_name, ?m.is_oms, ?m.mass_value, ?m.mass_unit, ?m.vol_value, ?m.vol_unit, ?m.gd_sid, ?m.gd_name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\medpack', 'med', 'shar')>0
  IF USED('med')
   USE IN med
  ENDIF 
 ELSE 
  WAIT "medpack..." WINDOW NOWAIT 
  SELECT med
  SCAN 
  
   SCATTER MEMVAR 
   
   cmd01 = 'INSERT INTO dim.medpack '
   cmd02 = '(r_up, name, pmp_id, type, qty_value, qty_unit, mass_value, mass_unit, vol_value, vol_unit)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.r_up, ?m.name, ?m.pmp_id, ?m.type, ?m.qty_value, ?m.qty_unit, ?m.mass_value, ?m.mass_unit, ?m.vol_value, ?m.vol_unit)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\medicament_mfc', 'med', 'shar')>0
  IF USED('med')
   USE IN med
  ENDIF 
 ELSE 
  WAIT "medicament_mfc..." WINDOW NOWAIT 
  SELECT med
  SCAN 
  
   SCATTER MEMVAR 
   *m.mfc_country = m.mfc_countr
   *m.own_country = m.own_countr
   m.d_issued    = IIF(!EMPTY(m.d_issued), m.d_issued, NULL)
   m.d_end       = IIF(!EMPTY(m.d_end), m.d_end, NULL)
   
   cmd01 = 'INSERT INTO dim.medicament_mfc '
   cmd02 = '(dd_id, md_id, name, n_ru, d_issued, d_end, '
   cmd03 = 'version_id)'
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.dd_id, ?m.md_id, ?m.m_name, ?m.n_ru, ?m.d_issued, ?m.d_end,'
   cmd08 = '?m.version_id)'
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\mkb10', 'mkb', 'shar', 'ds')>0
  IF USED('mkb')
   USE IN mkb
  ENDIF 
 ELSE 
  WAIT "mkb..." WINDOW NOWAIT 
  SELECT mkb
  SCAN 
  
   SCATTER MEMVAR 
   m.code = ds
   m.name = name_ds
   m.sex  = sex
   
   cmd01 = 'INSERT INTO dim.Ds '
   cmd02 = '(code, name, sex)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name, ?m.sex)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\profot', 'profot', 'shar', 'otd')>0
  IF USED('profot')
   USE IN profot
  ENDIF 
 ELSE 
  WAIT "profot..." WINDOW NOWAIT 
  SELECT profot
  SCAN 
  
   m.code = otd
   m.name = full_name
   
   cmd01 = 'INSERT INTO dim.ProfOt '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pCommon+'\prv002xx', 'prv', 'shar', 'profil')>0
  IF USED('prv')
   USE IN prv
  ENDIF 
 ELSE 
  WAIT "prv002xx..." WINDOW NOWAIT 
  SELECT prv
  SCAN 
  
   m.code = profil
   m.name = pr_name
   
   cmd01 = 'INSERT INTO dim.Prv002 '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\kpresl', 'kpresl', 'shar', 'tip')>0
  IF USED('kpresl')
   USE IN kpresl
  ENDIF 
 ELSE 
  WAIT "kpresl..." WINDOW NOWAIT 
  SELECT kpresl
  SCAN 
  
   m.code = tip
   m.name = name_tip
   
   cmd01 = 'INSERT INTO dim.Tip '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\ososch', 'ososch', 'shar')>0
  IF USED('ososch')
   USE IN ososch
  ENDIF 
 ELSE 
  WAIT "ososch..." WINDOW NOWAIT 
  SELECT ososch
  SCAN 
  
   m.code = d_type
   m.name = n_type
   
   cmd01 = 'INSERT INTO dim.ososch '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onmrf_xx', 'onc', 'shar')>0
  IF USED('onc')
   USE IN onc
  ENDIF 
 ELSE 
  WAIT "onmrf_xx..." WINDOW NOWAIT 
  SELECT onc
  SCAN 
  
   m.code = cod_mrf
   m.name = ALLTRIM(name_mrf)
   
   cmd01 = 'INSERT INTO dim.mrf '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onmrfrxx', 'onc', 'shar')>0
  IF USED('onc')
   USE IN onc
  ENDIF 
 ELSE 
  WAIT "onmrfrxx..." WINDOW NOWAIT 
  SELECT onc
  SCAN 
  
   m.code     = id_r_m
   m.mrf_code = cod_mrf
   m.name     = ALLTRIM(name_r_m)
   
   cmd01 = 'INSERT INTO dim.mrf_rslt '
   cmd02 = '(code, mrf_code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.mrf_code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 * В DWH справочник Rslt
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\rsv009', 'rsv', 'shar', 'rslt')>0
  IF USED('rsv')
   USE IN rsv
  ENDIF 
 ELSE 
  WAIT "rsv..." WINDOW NOWAIT 
  SELECT rsv
  SCAN 
  
   m.code = rslt
   m.name = rs_name
   
   cmd01 = 'INSERT INTO dim.Rslt '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 
 * В DWH справочник Rslt

 * В DWH справочник Prvs
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\spv015', 'spv', 'shar', 'code')>0
  IF USED('spv')
   USE IN spv
  ENDIF 
 ELSE 
  WAIT "spv..." WINDOW NOWAIT 
  SELECT spv
  SCAN 
  
   m.code = code
   m.name = name
   
   cmd01 = 'INSERT INTO dim.Prvs '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 
 * В DWH справочник Prvs

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\isv012', 'isv', 'shar', 'ishod')>0
  IF USED('isv')
   USE IN isv
  ENDIF 
 ELSE 
  WAIT "isv..." WINDOW NOWAIT 
  SELECT isv
  SCAN 
  
   m.code = ishod
   m.name = is_name
   
   cmd01 = 'INSERT INTO dim.Ishod '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onpcelxx', 'p_cel', 'shar')>0
  IF USED('p_cel')
   USE IN p_cel
  ENDIF 
 ELSE 
  WAIT "p_cel..." WINDOW NOWAIT 
  SELECT p_cel
  SCAN 
  
   m.code = cod_pc
   m.name = name_pc
   
   cmd01 = 'INSERT INTO dim.p_cel '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onczabxx', 'c_zab', 'shar')>0
  IF USED('c_zab')
   USE IN c_zab
  ENDIF 
 ELSE 
  WAIT "c_zab..." WINDOW NOWAIT 
  SELECT c_zab
  SCAN 
  
   m.code = cod_cz
   m.name = name_cz
   
   cmd01 = 'INSERT INTO dim.c_zab '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onreasxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onreasxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code = cod_reas
   m.name = name_reas
   
   cmd01 = 'INSERT INTO dim.ds1_t '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onstadxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onstadxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code = cod_st
   m.ds   = ds
   m.st   = st
   
   cmd01 = 'INSERT INTO dim.stad '
   cmd02 = '(code, ds, st)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.ds, ?m.st)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\ontum_xx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "ontum_xx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code = cod_t
   m.ds   = ds
   m.t    = t
   m.name = name_t
   
   cmd01 = 'INSERT INTO dim.onk_t '
   cmd02 = '(code, ds, t, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.ds, ?m.t, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onnod_xx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onnod_xx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code = cod_n
   m.ds   = ds
   m.n    = n
   m.name = name_n
   
   cmd01 = 'INSERT INTO dim.onk_n '
   cmd02 = '(code, ds, n, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.ds, ?m.n, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onmet_xx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onmet_xx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code = cod_m
   m.ds   = ds
   m.m    = m
   m.name = name_m
   
   cmd01 = 'INSERT INTO dim.onk_m '
   cmd02 = '(code, ds, m, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.ds, ?m.m, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onconsxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onconsxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code = cod_cons
   m.name = name_cons
   
   cmd01 = 'INSERT INTO dim.consiliumreason '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onigh_xx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onigh_xx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code = cod_igh
   m.name = ALLTRIM(name_igh)
   m.igh  = ALLTRIM(igh)
   
   cmd01 = 'INSERT INTO dim.igh '
   cmd02 = '(code, name, igh)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name, ?m.igh)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onigrtxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onigrtxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code     = id_r_i
   m.igh_code = cod_igh
   m.r_i      = ALLTRIM(r_i)
   m.name     = ALLTRIM(name_r_i)
   
   cmd01 = 'INSERT INTO dim.igh_rslt '
   cmd02 = '(code, igh_code, r_i, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.igh_code, ?m.r_i, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onnaprxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onnaprxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code     = cod_vn
   m.name     = ALLTRIM(name_vn)
   
   cmd01 = 'INSERT INTO dim.onnapr '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onprotxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onprotxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code     = cod_prot
   m.name     = ALLTRIM(name_prot)
   
   cmd01 = 'INSERT INTO dim.onprot '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onlechxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onlechxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code     = cod_tlech
   m.name     = ALLTRIM(name_tlech)
   
   cmd01 = 'INSERT INTO dim.onlech '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onhir_xx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onhir_xx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code     = cod_thir
   m.name     = ALLTRIM(name_thir)
   
   cmd01 = 'INSERT INTO dim.onhir '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onleklxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onleklxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code     = cod_tlek_l
   m.name     = ALLTRIM(nametlek_l)
   
   cmd01 = 'INSERT INTO dim.onlekl '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onlekvxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onlekvxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code     = cod_tlek_v
   m.name     = ALLTRIM(nametlek_v)
   
   cmd01 = 'INSERT INTO dim.onlekv '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\onluchxx', 'on', 'shar')>0
  IF USED('on')
   USE IN on
  ENDIF 
 ELSE 
  WAIT "onluchxx..." WINDOW NOWAIT 
  SELECT on
  SCAN 
  
   m.code     = cod_tluch
   m.name     = ALLTRIM(name_tluch)
   
   cmd01 = 'INSERT INTO dim.onluch '
   cmd02 = '(code, name)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 IF OpenFile(pCommon+'\dopreason', 'dopr', 'shar')>0
  IF USED('dopr')
   USE IN dopr
  ENDIF 
 ELSE 
  WAIT "dopreason..." WINDOW NOWAIT 
  SELECT dopr
  SCAN 
  
   m.code = code
   m.name = ALLTRIM(name)
   m.Typ  = Typ
   
   cmd01 = 'INSERT INTO dim.dopreason '
   cmd02 = '(code, name, typ)'
   cmd03 = ''
   cmd04 = ''
   cmd05 = ''
   cmd06 = 'VALUES '
   cmd07 = '(?m.code, ?m.name, ?m.Typ)'
   cmd08 = ''
   cmd09 = ''
   cmd10 = ''
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10
   IF SQLEXEC(nHandl, cmdAll)!=-1
   ENDIF 

  ENDSCAN 
  USE 

  WAIT CLEAR 
  
 ENDIF 

 m.d_end = SECONDS()
 WAIT CLEAR 
 
 IF USED('transid')
  USE IN transid
 ENDIF 
 
 MESSAGEBOX("ВРЕМЯ ОБРАБОТКИ: "+TRANSFORM(m.d_end-m.d_beg,'999999999') +' сек.',0+64,'')

 *IF SQLEXEC(nHandl, "ALTER DATABASE kms SET MULTI_USER")==-1
 * MESSAGEBOX("БД KMS НЕ УДАЛОСЬ ПЕРЕВЕСТИ"+CHR(13)+CHR(10)+;
 * "В МНОГОПОЛЬЗОВАТЕЛЬСКИЙ РЕЖИМ!!", 0+64, "")
 *ENDIF 

 =SQLDISCONNECT(nHandl)

RETURN 

FUNCTION SetSession()
 IF SQLEXEC(nHandl, "SET QUOTED_IDENTIFIER ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET QUOTED_IDENTIFIER ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET CONCAT_NULL_YIELDS_NULL ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET CONCAT_NULL_YIELDS_NULL ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET ANSI_NULLS ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET ANSI_NULLS ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET ANSI_PADDING ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET ANSI_PADDING ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET ANSI_WARNINGS ON")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET ANSI_WARNINGS ON')
  RETURN 
 ENDIF 
 IF SQLEXEC(nHandl, "SET NUMERIC_ROUNDABORT OFF")<=0
  =AERROR(errarr)
  =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'SET NUMERIC_ROUNDABORT OFF')
  RETURN 
 ENDIF 
RETURN 

FUNCTION DropDataBase()
RETURN

FUNCTION CreateDataBase()
RETURN 

FUNCTION OpenFoxDb()
 
 CREATE CURSOR transid (foxid i, sqlid i)
 INDEX ON foxid TAG foxid
 SET ORDER TO foxid

RETURN .T.

PROCEDURE OpenBases
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\tarifn', 'tarif', 'shar', 'cod')>0
  USE IN aisoms
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\profot', 'profot', 'shar', 'otd')>0
  USE IN aisoms
  USE IN tarif
  IF USED('profot')
   USE IN profot
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pCommon+'\periods', 'periods', 'shar', 'period')>0
  USE IN aisoms
  USE IN tarif
  USE IN profot
  IF USED('periods')
   USE IN periods
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pCommon+'\mkb_c', 'mkb_c', 'shar', 'ds')>0
  USE IN aisoms
  USE IN tarif
  USE IN profot
  USE IN periods
  IF USED('mkb_c')
   USE IN mkb_c
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  USE IN aisoms
  USE IN tarif
  USE IN profot
  USE IN periods
  USE IN mkb_c
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\pilots', 'pilots', 'shar', 'lpu_id')>0
  USE IN aisoms
  USE IN tarif
  USE IN profot
  USE IN periods
  USE IN mkb_c
  USE IN pilot
  IF USED('pilots')
   USE IN pilots
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\horlpu', 'horlpu', 'shar', 'lpu_id')>0
  USE IN aisoms
  USE IN tarif
  USE IN profot
  USE IN periods
  USE IN mkb_c
  USE IN pilot
  USE IN pilots
  IF USED('horlpu')
   USE IN horlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\horlpus', 'horlpus', 'shar', 'lpu_id')>0
  USE IN aisoms
  USE IN tarif
  USE IN profot
  USE IN periods
  USE IN mkb_c
  USE IN pilot
  USE IN pilots
  USE IN horlpu
  IF USED('horlpus')
   USE IN horlpus
  ENDIF 
  RETURN 
 ENDIF 
 

 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\Tarion.dbf')
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\Tarion', 'tarion', 'shar', 'cod')>0
   IF USED('tarion')
    USE IN tarion 
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\medicament.dbf')
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\medicament', 'medx', 'shar', 'dd_sid')>0
   IF USED('medx')
    USE IN medx
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\medpack.dbf')
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\medpack', 'medpack', 'shar', 'r_up')>0
   IF USED('medpack')
    USE IN medpack
   ENDIF 
  ENDIF 
 ENDIF 
 IF fso.FileExists(pbase+'\'+gcperiod+'\'+'nsi'+'\medicament_mfc.dbf')
  IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\medicament_mfc', 'mfc', 'shar', 'dd_id')>0
   IF USED('mfc')
    USE IN mfc
   ENDIF 
  ENDIF 
 ENDIF 
RETURN 

PROCEDURE FactServices
  SELECT Talon
  SCAN 
   SCATTER FIELDS EXCEPT recid MEMVAR 
   IF m.sqlid>0
    *LOOP 
   ENDIF 
   
   m.s_id   = m.recid_lpu
   m.sqlid  = 0
   m.sqldt  = {}
   m.tal_d  = {}
   m.ds_onk = 0
   m.p_cel  = ''
   m.dn     = 0
   m.reab   = 0
   m.c_zab  = 0
   
   m.period    = m.gcPeriod
   m.period_id = IIF(SEEK(m.period, 'periods'), periods.id, 0)
   
   m.otd1   = SUBSTR(m.otd,1,1)
   m.otd23  = SUBSTR(m.otd,2,2)
   m.otd456 = SUBSTR(m.otd,4,3)
   m.otdn  = SUBSTR(m.otd,7,2)
   
   m.date_ord = IIF(!EMPTY(m.date_ord), m.date_ord, NULL)
   m.tal_d    = IIF(!EMPTY(m.tal_d), m.tal_d, NULL)

   m.sex = people.w
   m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
   m.ages = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)
   
   m.tarif = tarif.tarif
   m.ismek = IIF(!EMPTY(err.c_err), .T., .F.)
   
   DO CASE 
    CASE IsUsl(m.cod) 
     m.usltip = 1
    CASE IsEKO(m.cod)
     m.usltip = 3
    CASE IsKDS(m.cod) OR IsKDP(m.cod)
     m.usltip = 2
    CASE IsMes(m.cod)
     m.usltip = 4
    CASE IsVMP(m.cod)
     m.usltip = 5
    OTHERWISE 
     m.usltip = 0
   ENDCASE 
   
   m.usl_ok = INT(VAL(profot.usl_ok))
   
   m.IsOnk = IIF(LEFT(m.ds,1)='C' OR BETWEEN(LEFT(m.ds,3), 'D00', 'D09') OR ;
   	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
   m.IsOnk2 = IIF(INLIST(SUBSTR(otd,4,3),'018','060'), .T., .F.)
   
   m.IsDental = IsDental(m.cod, m.lpuid, m.mcod, m.ds)
   
   m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)
   m.Is02      = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
   m.prmcod    = people.prmcod
   m.prmcods   = people.prmcods
   
   m.dopreason = NULL   
   IF !m.IsDental
    DO CASE 
     CASE IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)
      m.dopreason = 1 && Tpn=r в тарифе
     CASE m.d_type='s'
      m.dopreason = 2 && Симультанное хирургическое вмешательство
     CASE m.IsPilot AND INLIST(SUBSTR(m.otd,2,2),'08')
      m.dopreason = 3 && Женская консультация при медицинском учреждении только пилоты!
     CASE INLIST(SUBSTR(m.otd,2,2),'70','73','90','93') AND IsStac(m.mcod) && Может быть только в стационарах
      m.dopreason = 4 && приемные отделения с коечным/без коечного фонда,,выездная бригада
     CASE SUBSTR(m.otd,2,2)='01' AND IsStac(m.mcod) && КДО, параклиника (только для ничьих и чужих)
      m.dopreason = 5 && приемные отделения с коечным/без коечного фонда
     CASE m.ord=7 AND m.lpu_ord=7665 && УМО (только для ничьих и чужих)
      m.dopreason = 6 && УМО
     ** Добавлено 16.04.2019 по требованию Согаза
     CASE INLIST(INT(m.cod/1000),29,129,49,149) AND people.mcod!=people.prmcod AND people.tip_p=3 && только для ничьих и чужих!
      m.dopreason = 7 && 29,129,49,149 раздел для госпитализированных чужих
     ** Добавлено 16.04.2019 по требованию Согаза
    ENDCASE 
   ELSE 
    DO CASE 
     CASE IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)
      m.dopreason = 8 && Tpn=r в тарифе, стоматология
     CASE SUBSTR(m.otd,2,2)='08'
      m.dopreason = 9 && Женская консультация при медицинском учреждении
     CASE INLIST(SUBSTR(m.otd,2,2),'70','73','93') AND IsStac(m.mcod)
      m.dopreason = 10 && приемные отделения с коечным/без коечного фонда,,выездная бригада
    ENDCASE 
   ENDIF 

   IF m.tdat1<{01.09.2019}

   m.Typ = NULL
   m.Mp  = NULL
   m.vz  = NULL
   
  * для 201901
  * в oms6cn Mp сбрасывается, после чего 4 для допуслуг терапии, p для АПП терапии для стомат тоже может быть p
  * typ = p для подушевых
  * модулем makepr4n заполняется поле Mm='P' - в ud-файл, 'Y' - в s_bad
  m.Mp = IIF(!EMPTY(Mp), Mp, Typ)
  m.Mp = IIF(m.Mp='p' AND m.IsDental, 's', m.Mp)
  m.Mp = IIF(EMPTY(m.Mp) AND (IsMes(m.cod) OR IsVMP(m.cod)), 'm', m.Mp)

  IF EMPTY(m.Mp)

  IF m.IsDental
  DO CASE 
   CASE EMPTY(m.prmcods) && неприкрепленные
    m.Mp = 's'
   
   CASE m.mcod  = m.prmcods && свои пациенты
    DO CASE 
     CASE m.IsTpnR = .T. OR INLIST(m.otd23,'08') && tpn='r' - 3 услуги по июлю 2019, 08 - 4
      m.Mp = '8'
     CASE INLIST(m.otd23,'70','73') AND IsStac(m.mcod) && 23 услуги
      m.Mp = '8'
     CASE m.otd23='93' AND IsStac(m.mcod) && ни одной!
      m.Mp = '8'
     OTHERWISE 
       m.Mp = 's'
    ENDCASE 
    
   CASE m.mcod != m.prmcods && чужие пациенты
    m.Mp = 's'

   OTHERWISE 

  ENDCASE 

  ELSE && IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)

  DO CASE 
   CASE EMPTY(m.prmcod) && неприкрепленные
    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd23,'08')) && Добавление условия pilot ничего не меняет
      m.Mp = '4'
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.Mp = 'm'
     CASE INLIST(m.otd23,'01') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE INLIST(m.otd23,'70','73','90','93') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE m.ord=7 AND m.lpu_ord=7665
      m.Mp = '4'
     *CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=people.prmcod AND people.tip_p=3 
     * m.Mp = '4'
     *CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=people.prmcod AND people.tip_p=3 
     * m.Mp = '4'
     OTHERWISE 
       m.Mp = 'p'
    ENDCASE 
   
   CASE m.mcod  = m.prmcod && свои пациенты
    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd23,'08'))
      m.Mp = '4'
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.Mp = 'm'
     CASE INLIST(m.otd23,'70','73','93') AND IsStac(m.mcod)
      m.Mp = '4'
     OTHERWISE 
       m.Mp = 'p'
    ENDCASE 
    
   CASE m.mcod != m.prmcod && чужие пациенты
    DO CASE 
     CASE m.IsTpnR = .T. OR m.d_type='s' OR (m.IsPilot AND INLIST(m.otd23,'08'))
      m.Mp = '4'
     CASE IsMes(m.cod) OR IsVMP(m.cod)
      m.Mp = 'm'
     CASE INLIST(m.otd23,'01') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE INLIST(m.otd23,'70','73','90','93') AND IsStac(m.mcod)
      m.Mp = '4'
     CASE m.ord=7 AND m.lpu_ord=7665
      m.Mp = '4'
     *CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=people.prmcod AND people.tip_p=3 
     * m.Mp = '4'
     *CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=people.prmcod AND people.tip_p=3 
     * m.Mp = '4'
     OTHERWISE 
      m.Mp = 'p'
    ENDCASE 
   OTHERWISE 
  ENDCASE 

  ENDIF IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)

  ENDIF 

  * для 201901

  IF m.IsDental
   DO CASE 
    CASE EMPTY(m.prmcods) && неприкрепленные
     m.Typ = '0'
    CASE m.mcod  = m.prmcods && свои пациенты
     m.Typ = '1'
    CASE m.mcod != m.prmcods && чужие пациенты
     m.Typ = '2'
   ENDCASE 
  ELSE
   DO CASE 
    CASE EMPTY(m.prmcod) && неприкрепленные
     m.Typ = '0'
    CASE m.mcod  = m.prmcod && свои пациенты
     m.Typ = '1'
    CASE m.mcod != m.prmcod && чужие пациенты
     m.Typ = '2'
   ENDCASE 
  ENDIF

  IF FIELD('mm')='MM'
   m.Vz = IIF(Mm='P', 1, m.vz)
   m.Vz = IIF(Mm='Y', 0, m.vz)
  ELSE 
   IF m.Typ = '2' AND m.Mp='p'
    DO CASE 
     CASE !m.ispilot AND !m.ishorlpu
     CASE  m.otd23 = '08'
     OTHERWISE 
      IF (m.Is02 OR INLIST(m.otd23, '08', '91') OR (m.profil = '100' AND INLIST(m.otd23, '00', '92'))) OR m.lpu_ord > 0
       m.Vz = 1
      ELSE 
       m.vz = 0
      ENDIF 
    ENDCASE 
   ENDIF 
  ENDIF 
  
   IF m.Typ = '2' AND m.Mp='s'
    DO CASE 
     CASE !m.ispilots AND !m.ishorlpus
     CASE  m.otd23 = '08'
     OTHERWISE 
      IF (m.Is02 OR INLIST(m.otd23, '08', '91') OR (m.profil = '100' AND INLIST(m.otd23, '00', '92'))) OR m.lpu_ord > 0
       m.Vz = 1
      ELSE 
       m.vz = 0
      ENDIF 
    ENDCASE 
   ENDIF 

  IF m.Vz=1
   DO CASE 
       CASE m.lpu_ord>0 && vz=1, направление, в т.ч. договор с ДШО/ШО, договор на проведение вакцинопрофилактики и "актив" ССиНМП
        m.vz = 1
       CASE m.Is02 && vz=2, неотложная помощь (по реестру медицинских услуг)
        m.vz = 2
       CASE m.profil='100' AND INLIST(m.otd23,'00','92') && vz=3, услуги, оказанные в травмапункте (в дополнение к  коду 2)
        m.vz = 3
       CASE m.otd23='08' && vz=4, услуги ЖК
        m.vz = 4
       CASE m.otd23='91' && vz=5, услуги ЦЗ
        m.vz = 5
       OTHERWISE 
        m.vz = 9 && что-то иное!
   ENDCASE 
  ENDIF 
  
  ENDIF 

   cmd01 = 'INSERT INTO FactServices '
   cmd02 = '(mcod, lpuid, period, period_id, sn_pol, c_i, ds, ds_0, pcod, otd, otd1, otd23, otd456, otdn, cod, tip, usltip, d_u, k_u, kd_fact, n_kd, d_type, s_all,'
   cmd03 = 'ismek, s_lek, profil, rslt, prvs, ishod, kur, ds_2, ds_3,det, k2, vnov_m, novor,  n_u, n_vmp, ord, date_ord,'
   cmd04 = 'lpu_ord, s_id, fil_id, ds_onk, p_cel, dn, reab, tal_d, napr_v_in, c_zab, mp, typ, sex, ages, tarif, usl_ok,'
   cmd05 = 'ispilot,ispilots,ishorlpu,ishorlpus,' 
   cmd06 = 'isonk, isonk2, IsDental, vz, prmcod, prmcods, dopreason)'
   cmd07 = 'VALUES '
   cmd08 = '(?m.mcod, ?m.lpuid, ?m.period, ?m.period_id, ?m.sn_pol, ?m.c_i, ?m.ds, ?m.ds_0, ?m.pcod, ?m.otd, ?m.otd1, ?m.otd23, ?m.otd456, ?m.otdn, ?m.cod, ?m.tip, ?m.usltip, ?m.d_u, ?m.k_u, ?m.kd_fact, ?m.n_kd, ?m.d_type, ?m.s_all,'
   cmd09 = '?m.ismek, ?m.s_lek, ?m.profil, ?m.rslt, ?m.prvs, ?m.ishod, ?m.kur, ?m.ds_2, ?m.ds_3, ?m.det, ?m.k2, ?m.vnov_m, ?m.novor, ?m.n_u, ?m.n_vmp, ?m.ord, ?m.date_ord,'
   cmd10 = '?m.lpu_ord, ?m.s_id, ?m.fil_id, ?m.ds_onk, ?m.p_cel, ?m.dn, ?m.reab, ?m.tal_d, ?m.napr_v_in, ?m.c_zab, ?m.mp, ?m.typ, ?m.sex, ?m.ages, ?m.tarif, ?m.usl_ok,'
   cmd11 = '?m.ispilot, ?m.ispilots, ?m.ishorlpu, ?m.ishorlpus,' 
   cmd12 = '?m.isonk, ?m.isonk2, ?m.IsDental,?m.vz, ?m.prmcod, ?m.prmcods,?m.dopreason)'
   cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   IF SQLEXEC(nHandl, cmdAll)!=-1
    lnGoodRecs = lnGoodRecs + 1
   ELSE 
    =AERROR(errarr)
    =MESSAGEBOX(ALLTRIM(errarr(2)), 16, m.mcod+'FactServices')
    =MESSAGEBOX(ALLTRIM(errarr(3)), 16, m.mcod+'FactServices')
    EXIT 
    *LOOP    
   ENDIF 
   
   *IF SQLEXEC(nHandl, "select @@IDENTITY as newid", "cursid") != -1
   * m.sqlid = cursid.newid
   * USE IN cursid
   *ENDIF 
   
   *SELECT talon 
   *REPLACE sqlid WITH m.sqlid, sqldt WITH DATETIME()
   REPLACE sqldt WITH DATETIME()

  ENDSCAN 
RETURN 

PROCEDURE FactDrugs
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_LS'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_LS'+m.qcod, 'onk_ls', 'share', 'recid_s')>0
    IF USED('onk_ls')
     USE IN onk_ls
    ENDIF 
   ENDIF 
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod, 'onk_sl', 'share', 'recid')>0
    IF USED('onk_sl')
     USE IN onk_sl
    ENDIF 
   ENDIF 
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_USL'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_USL'+m.qcod, 'onk_usl', 'share', 'recid')>0
    IF USED('onk_sl')
     USE IN onk_sl
    ENDIF 
   ENDIF 
  ENDIF 
  
  IF USED('onk_ls') AND USED('onk_sl') AND USED('onk_usl')
   lnGoodRecs = 0
   SET ORDER TO recid_lpu IN talon 
   SELECT onk_sl
   SET RELATION TO recid_s INTO talon 
   SELECT onk_usl
   SET RELATION TO recid_sl INTO onk_sl 
   SELECT onk_ls
   SET RELATION TO recid_usl INTO onk_usl
   SET RELATION TO sn_pol INTO people ADDITIVE 
   SCAN 
    SCATTER MEMVAR 
    
    m.isokcod = IIF(INLIST(m.cod, 97158, 81094), .t., .f.)
    
    m.period      = m.gcPeriod
    m.period_id   = IIF(SEEK(m.period, 'periods'), periods.id, 0)
    m.serv_id     = talon.sqlid
    m.case_id     = onk_sl.sqlid
    m.atttyp      = talon.Typ
    m.ds          = talon.ds
    m.isokds      = IIF(SEEK(m.ds, 'mkb_c'), .T., .F.)

    m.d_u         = date_inj

    m.sex = people.w
    m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.ages = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)
    
    m.gd_sid  = IIF(SEEK(m.sid, 'medx'), ALLTRIM(medx.gd_sid), NULL)
    m.v_unit  = IIF(SEEK(m.sid, 'medx'), ALLTRIM(medx.vol_unit), '')
    m.edizm   = IIF(!EMPTY(m.v_unit), 'мл', 'мг')
    m.tarif   = IIF(SEEK(m.gd_sid, 'tarion', 'cod'), tarion.ston, 0)
   
    IF m.edizm = 'мл'
     m.en     = IIF(SEEK(LEFT(m.r_up,10), 'medpack'), medpack.vol_value, 0) && например, 4 мл
    ELSE && мг
     m.en     = IIF(SEEK(LEFT(m.r_up,10), 'medpack'), medpack.mass_value, 0) && например, 440 мг
    ENDIF 

    cmd01 = 'INSERT INTO FactDrugs '
    cmd02 = '(usl_id, case_id, mcod, lpuid, period_id, period, regnum, d_u, code_sh, n_par, r_up, tip_opl, '
    cmd03 = 'ds, isokds, n_ru, ot_d, dt_q, dt_d, sid, gd_sid, edizm, tarif, en, s_all, sn_pol, c_i, ages, sex, atttyp,'
    cmd04 = 'cod, isokcod'
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.recid_usl, 0, ?m.mcod, ?m.lpuid, ?period_id, ?m.period, ?m.regnum, ?m.d_u, ?m.code_sh, ?m.n_par, ?m.r_up, ?m.tip_opl, '
    cmd09 = '?m.ds, ?m.isokds, ?m.n_ru, ?m.ot_d, ?m.dt_q, ?m.dt_d, ?m.sid, ?m.gd_sid, ?m.edizm, ?m.tarif, ?m.en, ?m.s_all, ?m.sn_pol, ?m.c_i, ?m.ages, ?m.sex, ?m.atttyp,'
    cmd10 = '?m.cod, ?m.isokcod'
    cmd11 = ''
    cmd12 = ')'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     SET STEP ON ON 
     =AERROR(errarr)
     =MESSAGEBOX(ALLTRIM(errarr(2)), 16, m.mcod+'FactDrugs'+m.recid_usl)
     =MESSAGEBOX(ALLTRIM(errarr(3)), 16, m.mcod+'FactDrugs'+m.recid_usl)
     EXIT 
     *LOOP    
    ENDIF 

   *IF SQLEXEC(nHandl, "select @@IDENTITY as newid", "cursid") != -1
   * m.sqlid = cursid.newid
   * USE IN cursid
   *ENDIF 
   
   *SELECT onk_ls
   *REPLACE sqlid WITH m.sqlid, sqldt WITH DATETIME()
   REPLACE sqldt WITH DATETIME()

   ENDSCAN 
   SET RELATION OFF INTO talon 
   SET RELATION OFF INTO people 
  ENDIF 
  IF USED('onk_ls')
   USE IN onk_ls
  ENDIF 
  IF USED('onk_sl')
   USE IN onk_sl
  ENDIF 
  IF USED('onk_usl')
   USE IN onk_usl
  ENDIF 

  SELECT aisoms
 RETURN 
 
 PROCEDURE FactConsiliums
   IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_CONS'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_CONS'+m.qcod, 'onk_cons', 'share')>0
    IF USED('onk_cons')
     USE IN onk_cons
    ENDIF 
   ENDIF 
  ENDIF 
  
  IF USED('onk_cons')
   lnGoodRecs = 0
   SET ORDER TO recid_lpu IN talon 
   SELECT onk_cons
   SET RELATION TO recid_s INTO talon 
   SET RELATION TO sn_pol INTO people ADDITIVE 
   SCAN 
    SCATTER MEMVAR 
    
    IF EMPTY(m.pr_cons)
     LOOP 
    ENDIF 
    
    m.period    = m.gcPeriod
    m.period_id = IIF(SEEK(m.period, 'periods'), periods.id, 0)
    *m.services_id = talon.sqlid
    m.recid     = talon.sqlid
    m.atttyp    = talon.Typ
    m.ds        = talon.ds

    m.reason    = m.pr_cons
    m.d_u       = m.dt_cons

    m.sex       = people.w
    m.adj       = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.ages      = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)

    cmd01 = 'INSERT INTO FactConsiliums '
    cmd02 = '(s_id, period_id, period, mcod, lpuid, sn_pol, c_i, ages, sex, AttTyp, '
    cmd03 = 'ds, cod, reason, d_u'
    cmd04 = ''
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.recid_s, ?m.period_id, ?m.period, ?m.mcod, ?m.lpuid, ?m.sn_pol, ?m.c_i, ?m.ages, ?m.sex, ?m.AttTyp, '
    cmd09 = '?m.ds, ?m.cod, ?m.reason, ?m.d_u'
    cmd10 = ''
    cmd11 = ''
    cmd12 = ')'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     LOOP 
    ENDIF 

    *IF SQLEXEC(nHandl, "select @@IDENTITY as newid", "cursid") != -1
    * m.sqlid = cursid.newid
    * USE IN cursid
    *ELSE 
    * LOOP 
    *ENDIF 
   
    *SELECT onk_cons
    *REPLACE sqlid WITH m.sqlid, sqldt WITH DATETIME()

   ENDSCAN 
   SET RELATION OFF INTO talon 
   SET RELATION OFF INTO people 
   USE IN onk_cons
  ENDIF 
RETURN 

PROCEDURE FactMek
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'err', 'share')>0
    IF USED('err')
     USE IN err
    ENDIF 
   ENDIF 
  ENDIF 
  
  IF USED('err')
   SELECT talon 
   SET ORDER TO recid
   SET RELATION TO sn_pol INTO people
   SET RELATION TO SUBSTR(otd,2,2) INTO profot ADDITIVE 
   SELECT err 
   SET ORDER TO rid
   lnGoodRecs = 0
   SET RELATION TO rid INTO talon 
   SCAN 
    SCATTER MEMVAR 
    
    m.id        = IIF(m.f='S', talon.recid_lpu, people.recid_lpu)
    m.fil_id    = talon.fil_id
        
    m.period    = m.gcPeriod
    m.period_id = IIF(SEEK(m.period, 'periods'), periods.id, 0)
    m.atttyp    = talon.Typ
    m.s_all     = talon.s_all

    m.ds        = talon.ds
    m.cod       = talon.cod
    m.d_u       = talon.d_u
    
    m.sn_pol    = talon.sn_pol
    m.c_i       = talon.c_i
    
    m.Mp        = talon.mp

    m.sex = people.w
    m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.ages = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)

    DO CASE 
     CASE IsUsl(m.cod) 
      m.usltip = 1
     CASE IsEKO(m.cod)
      m.usltip = 3
     CASE IsKDS(m.cod) OR IsKDP(m.cod)
      m.usltip = 2
     CASE IsMes(m.cod)
      m.usltip = 4
     CASE IsVMP(m.cod)
      m.usltip = 5
     OTHERWISE 
      m.usltip = 0
    ENDCASE 
    
    m.usl_ok = INT(VAL(profot.usl_ok))
	m.osn230 = ''
    
    IF FIELD('tip')!='TIP'	
	 m.tip = 0
	ENDIF 
    IF FIELD('dt')!='DT'	
	 m.dt = {}
	ENDIF 
    IF FIELD('usr')!='USR'	
	 m.usr = ''
	ENDIF 

    cmd01 = 'INSERT INTO FactMek '
    cmd02 = '(period_id, period, mcod, lpuid, fil_id, sn_pol, c_i, ages, sex, AttTyp, '
    cmd03 = 'recid_lpu,'
    cmd04 = 'usltip, usl_ok, mp, cod, '
    cmd05 = 'f, c_err, detail, tip, dt, usr, comment, osn230, s_all'
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.period_id, ?m.period, ?m.mcod, ?m.lpuid, ?fil_id, ?m.sn_pol, ?m.c_i, ?m.ages, ?m.sex, ?m.AttTyp, '
    cmd09 = '?m.id,'
    cmd10 = '?m.usltip, ?m.usl_ok, ?m.Mp, ?m.cod, '
    cmd11 = ''
    cmd12 = '?m.f, ?m.c_err, ?m.detail, ?m.tip, ?m.dt, ?m.usr, ?m.comment, ?m.osn230, ?m.s_all)'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     =AERROR(errarr)
     =MESSAGEBOX(ALLTRIM(errarr(2)), 16, m.mcod+'FactMek')
     =MESSAGEBOX(ALLTRIM(errarr(3)), 16, m.mcod+'FactMek')
     EXIT 
    ENDIF 

    *IF SQLEXEC(nHandl, "select @@IDENTITY as newid", "cursid") != -1
    * m.sqlid = cursid.newid
    * USE IN cursid
    *ELSE 
    * LOOP 
    *ENDIF 
   
    SELECT err
    *REPLACE sqlid WITH m.recid, sqldt WITH DATETIME()
    REPLACE sqldt WITH DATETIME()

   ENDSCAN 
   SET RELATION OFF INTO talon 
   USE 
   SELECT talon 
   SET RELATION OFF INTO people 
   SET RELATION OFF INTO profot
  ENDIF 
RETURN 

PROCEDURE FactCases
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod, 'onk_sl', 'share')>0
    IF USED('onk_sl')
     USE IN onk_sl
    ENDIF 
   ENDIF 
  ENDIF 
  
  IF USED('onk_sl')
   lnGoodRecs = 0
   SET ORDER TO recid_lpu IN talon 
   SELECT onk_sl
   SET RELATION TO recid_s INTO talon 
   SET RELATION TO sn_pol INTO people ADDITIVE 
   SCAN 
    SCATTER MEMVAR 
    
    m.period    = m.gcPeriod
    m.period_id = IIF(SEEK(m.period, 'periods'), periods.id, 0)
    *m.services_id = talon.sqlid
    *m.recid     = talon.sqlid
    m.atttyp    = talon.Typ

    m.ds        = talon.ds
    m.cod       = talon.cod
    m.d_u       = talon.d_u

    m.sex = people.w
    m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.ages = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)

    m.ds1_t = IIF(!EMPTY(onk_sl.ds1_t), onk_sl.ds1_t, NULL)
    m.stad  = IIF(!EMPTY(onk_sl.stad), onk_sl.stad, NULL)
    m.onk_t = IIF(!EMPTY(onk_sl.onk_t), onk_sl.onk_t, NULL)
    m.onk_n = IIF(!EMPTY(onk_sl.onk_n), onk_sl.onk_n, NULL)
    m.onk_m = IIF(!EMPTY(onk_sl.onk_m), onk_sl.onk_m, NULL)
    m.mtstz = IIF(!EMPTY(onk_sl.mtstz), .T., .F.)
    m.sod   = IIF(!EMPTY(onk_sl.sod), onk_sl.sod, NULL)
    m.k_fr  = IIF(!EMPTY(onk_sl.k_fr), onk_sl.k_fr, NULL)
    m.wei   = IIF(!EMPTY(onk_sl.wei), onk_sl.wei, NULL)
    m.hei   = IIF(!EMPTY(onk_sl.hei), onk_sl.hei, NULL)
    m.bsa   = IIF(!EMPTY(onk_sl.bsa), onk_sl.bsa, NULL)

    cmd01 = 'INSERT INTO FactCases '
    cmd02 = '(s_id, sl_id, period_id, period, mcod, lpuid, sn_pol, c_i, ages, sex, AttTyp, '
    cmd03 = 'ds1_t, stad, onk_t, onk_n, onk_m, mtstz, sod, k_fr, wei, hei, bsa,'
    cmd04 = 'ds, cod, d_u'
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.recid_s, ?m.recid, ?m.period_id, ?m.period, ?m.mcod, ?m.lpuid, ?m.sn_pol, ?m.c_i, ?m.ages, ?m.sex, ?m.AttTyp, '
    cmd09 = '?m.ds1_t, ?m.stad, ?m.onk_t, ?m.onk_n, ?m.onk_m, ?m.mtstz, ?m.sod, ?m.k_fr, ?m.wei, ?m.hei, ?m.bsa,'
    cmd10 = '?m.ds, ?m.cod, ?m.d_u'
    cmd11 = ''
    cmd12 = ')'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     =AERROR(errarr)
     =MESSAGEBOX(ALLTRIM(errarr(2)), 16, m.mcod+'FactCases')
     =MESSAGEBOX(ALLTRIM(errarr(3)), 16, m.mcod+'FactCases')
     EXIT 
    ENDIF 

    *IF SQLEXEC(nHandl, "select @@IDENTITY as newid", "cursid") != -1
    * m.sqlid = cursid.newid
    * USE IN cursid
    *ELSE 
    * LOOP 
    *ENDIF 
   
    SELECT onk_sl
    *REPLACE sqlid WITH m.recid, sqldt WITH DATETIME()
    REPLACE sqldt WITH DATETIME()

   ENDSCAN 
   SET RELATION OFF INTO talon 
   SET RELATION OFF INTO people 
   USE IN onk_sl
  ENDIF 
RETURN 

PROCEDURE FactOncoDiag
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_DIAG'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_DIAG'+m.qcod, 'onk', 'share')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ENDIF 
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod, 'onk_sl', 'share', 'recid')>0
    IF USED('onk_sl')
     USE IN onk_sl
    ENDIF 
   ENDIF 
  ENDIF 
  
  IF USED('onk') AND USED('onk_sl')
   lnGoodRecs = 0
   SET ORDER TO recid_lpu IN talon 
   SELECT onk_sl
   SET RELATION TO recid_s INTO talon 
   SELECT onk
   SET RELATION TO sn_pol INTO people ADDITIVE 
   SET RELATION TO recid_sl INTO onk_sl ADDITIVE 
   SCAN 
    SCATTER MEMVAR 
    
    m.tip = m.diag_tip
    
    m.period      = m.gcPeriod
    m.period_id   = IIF(SEEK(m.period, 'periods'), periods.id, 0)
    m.serv_id     = talon.sqlid
    m.case_id     = onk_sl.sqlid
    m.atttyp      = talon.Typ
    m.ds          = talon.ds

    m.d_u         = m.diag_date

    m.sex = people.w
    m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.ages = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)
    
    m.mrf      = IIF(m.tip=1, m.diag_code, NULL)
    m.mrf_rslt = IIF(m.tip=1, m.diag_rslt, NULL)
    m.igh      = IIF(m.tip=2, m.diag_code, NULL)
    m.igh_rslt = IIF(m.tip=2, m.diag_rslt, NULL)
    
    m.rslt = rec_rslt

    cmd01 = 'INSERT INTO FactOncoDiag '
    cmd02 = '(period_id, period, mcod, lpuid, case_id, serv_id, sn_pol, c_i, ages, sex, AttTyp, '
    cmd03 = 'ds, cod, tip, mrf, mrf_rslt, igh, igh_rslt, d_u, rslt, met_issl, sl_id'
    cmd04 = ''
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.period_id, ?m.period, ?m.mcod, ?m.lpuid, ?m.case_id, ?m.serv_id, ?m.sn_pol, ?m.c_i, ?m.ages, ?m.sex, ?m.AttTyp, '
    cmd09 = '?m.ds, ?m.cod, ?m.tip, ?m.mrf, ?m.mrf_rslt, ?m.igh, ?m.igh_rslt, ?m.d_u, ?m.rslt, ?m.met_issl, ?m.recid_sl'
    cmd10 = ''
    cmd11 = ''
    cmd12 = ')'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     LOOP 
    ENDIF 

    IF SQLEXEC(nHandl, "select @@IDENTITY as newid", "cursid") != -1
     m.sqlid = cursid.newid
     USE IN cursid
    ELSE 
     LOOP 
    ENDIF 
   
    SELECT onk
    REPLACE sqlid WITH m.sqlid, sqldt WITH DATETIME()

   ENDSCAN 
   SET RELATION OFF INTO talon 
   SET RELATION OFF INTO people 
   SET RELATION OFF INTO onk_sl
  ENDIF 
  IF USED('onk')
   USE IN onk
  ENDIF 
  IF USED('onk_sl')
   USE IN onk_sl
  ENDIF 
RETURN 

PROCEDURE FactReferrals
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_NAPR_V_OUT'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_NAPR_V_OUT'+m.qcod, 'onk', 'share')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ENDIF 
  ENDIF 
  
  IF USED('onk')
   lnGoodRecs = 0
   SET ORDER TO recid_lpu IN talon 
   SELECT onk
   SET RELATION TO recid_s INTO talon 
   SET RELATION TO sn_pol INTO people ADDITIVE 
   SCAN 
    SCATTER MEMVAR 
    
    m.period    = m.gcPeriod
    m.period_id = IIF(SEEK(m.period, 'periods'), periods.id, 0)
    *m.services_id = talon.sqlid
    m.recid     = talon.sqlid
    m.atttyp    = talon.Typ
    m.ds        = talon.ds

    m.d_u       = IIF(UPPER(FIELD('napr_date'))='NAPR_DATE', m.napr_date, NULL)
    m.n_ref     = IIF(UPPER(FIELD('NAP_NUMBER'))='NAP_NUMBER', m.nap_number, NULL)

    m.sex       = people.w
    m.adj       = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.ages      = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)

    cmd01 = 'INSERT INTO FactReferrals '
    cmd02 = '(s_id, period_id, period, mcod, lpuid, sn_pol, c_i, ages, sex, AttTyp, '
    cmd03 = 'ds, reason, lpu_id, d_u, n_ref'
    cmd04 = ''
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.recid_s, ?m.period_id, ?m.period, ?m.mcod, ?m.lpuid, ?m.sn_pol, ?m.c_i, ?m.ages, ?m.sex, ?m.AttTyp, '
    cmd09 = '?m.ds, ?m.napr_v_out, ?m.napr_mo, ?m.d_u, ?m.n_ref'
    cmd10 = ''
    cmd11 = ''
    cmd12 = ')'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     LOOP 
    ENDIF 

    *IF SQLEXEC(nHandl, "select @@IDENTITY as newid", "cursid") != -1
    * m.sqlid = cursid.newid
    * USE IN cursid
    *ELSE 
    * LOOP 
    *ENDIF 
   
    *SELECT onk
    *REPLACE sqlid WITH m.sqlid, sqldt WITH DATETIME()

   ENDSCAN 
   SET RELATION OFF INTO talon 
   SET RELATION OFF INTO people 
   USE IN onk
  ENDIF 
RETURN 

PROCEDURE FactDenials
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_PROT'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_PROT'+m.qcod, 'onk', 'share')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ENDIF 
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod, 'onk_sl', 'share', 'recid')>0
    IF USED('onk_sl')
     USE IN onk_sl
    ENDIF 
   ENDIF 
  ENDIF 
  
  IF USED('onk') AND USED('onk_sl')
   lnGoodRecs = 0
   SET ORDER TO recid_lpu IN talon 
   SELECT onk_sl
   SET RELATION TO recid_s INTO talon 
   SELECT onk
   *SET RELATION TO recid_s INTO talon 
   SET RELATION TO sn_pol INTO people ADDITIVE 
   SET RELATION TO recid_sl INTO onk_sl ADDITIVE 
   SCAN 
    SCATTER MEMVAR 
    
    m.period      = m.gcPeriod
    m.period_id   = IIF(SEEK(m.period, 'periods'), periods.id, 0)
    *m.services_id = talon.sqlid
    m.case_id     = onk_sl.sqlid
    m.atttyp      = talon.Typ
    m.ds          = talon.ds
    
    m.d_u  = talon.d_u
    m.sex  = people.w
    m.adj  = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.ages = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)

    m.d_u  = IIF(UPPER(FIELD('d_prot'))='D_PROT', m.d_prot, NULL)
    m.code = IIF(UPPER(FIELD('prot'))='PROT', m.prot, NULL)
    m.c_i  = talon.c_i

    cmd01 = 'INSERT INTO FactDenials '
    cmd02 = '(period_id, period, mcod, lpuid, case_id, sn_pol, c_i, ages, sex, AttTyp, '
    cmd03 = 'ds, d_u, code, sl_id'
    cmd04 = ''
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.period_id, ?m.period, ?m.mcod, ?m.lpuid, ?m.case_id, ?m.sn_pol, ?m.c_i, ?m.ages, ?m.sex, ?m.AttTyp, '
    cmd09 = '?m.ds, ?m.d_u, ?m.code, ?m.recid_sl'
    cmd10 = ''
    cmd11 = ''
    cmd12 = ')'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     LOOP 
    ENDIF 

    *IF SQLEXEC(nHandl, "select @@IDENTITY as newid", "cursid") != -1
    * m.sqlid = cursid.newid
    * USE IN cursid
    *ELSE 
    * LOOP 
    *ENDIF 
   
    *SELECT onk
    *REPLACE sqlid WITH m.sqlid, sqldt WITH DATETIME()

   ENDSCAN 
   SET RELATION OFF INTO talon 
   SET RELATION OFF INTO people 
  ENDIF 
  IF USED('onk')
   USE IN onk
  ENDIF 
  IF USED('onk_sl')
   USE IN onk_sl
  ENDIF 
RETURN 

PROCEDURE FactOncoServices
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_USL'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_USL'+m.qcod, 'onk', 'share')>0
    IF USED('onk')
     USE IN onk
    ENDIF 
   ENDIF 
  ENDIF 
  IF fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod+'.dbf')
   IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\ONK_SL'+m.qcod, 'onk_sl', 'share', 'recid')>0
    IF USED('onk_sl')
     USE IN onk_sl
    ENDIF 
   ENDIF 
  ENDIF 
  
  IF USED('onk') AND USED('onk_sl')
   lnGoodRecs = 0
   SET ORDER TO recid_lpu IN talon 
   SELECT onk_sl
   SET RELATION TO recid_s INTO talon 
   SELECT onk
   SET RELATION TO sn_pol INTO people ADDITIVE 
   SET RELATION TO recid_sl INTO onk_sl ADDITIVE 
   SCAN 
    SCATTER MEMVAR 
    
    m.period      = m.gcPeriod
    m.period_id   = IIF(SEEK(m.period, 'periods'), periods.id, 0)
    *m.serv_id     = talon.sqlid
    *m.case_id     = onk_sl.sqlid
    m.atttyp      = talon.Typ
    m.ds          = talon.ds
    m.cod         = talon.cod
    m.d_u         = talon.d_u

    m.d_u  = talon.d_u
    m.sex  = people.w
    m.adj  = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.ages = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)
    
    m.pptr = IIF(m.pptr=1, .t., .f.)

    cmd01 = 'INSERT INTO FactOncoServices '
    cmd02 = '(period_id, period, mcod, lpuid, case_id, sn_pol, c_i, ages, sex, AttTyp, '
    cmd03 = 'ds, cod, d_u, onlech, onhir, onlekl, onlekv, onluch, pptr, sl_id, usl_id'
    cmd04 = ''
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.period_id, ?m.period, ?m.mcod, ?m.lpuid, 0, ?m.sn_pol, ?m.c_i, ?m.ages, ?m.sex, ?m.AttTyp, '
    cmd09 = '?m.ds, ?m.cod, ?m.d_u, ?m.usl_tip, ?m.hir_tip, ?m.lek_tip_l, ?m.lek_tip_v, ?m.luch_tip, ?m.pptr, ?m.recid_sl, ?m.recid'
    cmd10 = ''
    cmd11 = ''
    cmd12 = ')'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     =AERROR(errarr)
     =MESSAGEBOX(ALLTRIM(errarr(2)), 16, m.mcod+'FactOncoServices')
     =MESSAGEBOX(ALLTRIM(errarr(3)), 16, m.mcod+'FactOncoServices')
     EXIT 
    ENDIF 

    *IF SQLEXEC(nHandl, "select @@IDENTITY as newid", "cursid") != -1
    * m.sqlid = cursid.newid
    * USE IN cursid
    *ELSE 
    * LOOP 
    *ENDIF 
   
    SELECT onk
    *REPLACE sqlid WITH m.sqlid, sqldt WITH DATETIME()
    REPLACE sqldt WITH DATETIME()

   ENDSCAN 
   SET RELATION OFF INTO talon 
   SET RELATION OFF INTO people 
   SET RELATION OFF INTO onk_sl 
 ENDIF 
 IF USED('onk')
  USE IN onk
 ENDIF 
 IF USED('onk_sl')
  USE IN onk_sl
 ENDIF 
RETURN 

PROCEDURE AisOms
  IF USED('aisoms')
   lnGoodRecs = 0
   SELECT aisoms 
   SCAN 
    SCATTER MEMVAR 

    cmd01 = 'INSERT INTO dbo.Aisoms '
    cmd02 = '(period, lpuid, mcod, pazval, finval, pazvals, finvals, paz, nsch, s_pred, s_lek, '
    cmd03 = 's_mek, s_532, s_avans, s_pr_avans, s_avans2, s_pr_avans2, e_mee, e_ekmp, dolg_b'
    cmd04 = ''
    cmd05 = ''
    cmd06 = ')'
    cmd07 = 'VALUES '
    cmd08 = '(?m.gcperiod, ?m.lpuid, ?m.mcod, ?m.pazval, ?m.finval, ?m.pazvals, ?m.finvals, ?m.paz, ?m.nsch, ?m.s_pred, ?m.s_lek, '
    cmd09 = '?m.sum_flk, ?m.s_532, ?m.s_avans, ?m.s_pr_avans, ?m.s_avans2, ?m.pr_avans2, ?m.e_mee, ?m.e_ekmp, ?m.dolg_b'
    cmd10 = ''
    cmd11 = ''
    cmd12 = ')'
    cmdAll = cmd01+cmd02+cmd03+cmd04+cmd05+cmd06+cmd07+cmd08+cmd09+cmd10+cmd11+cmd12
   
    IF SQLEXEC(nHandl, cmdAll)!=-1
     lnGoodRecs = lnGoodRecs + 1
    ELSE 
     =AERROR(errarr)
     =MESSAGEBOX(ALLTRIM(errarr(2)), 16, 'Cannot make connection')
     =MESSAGEBOX(ALLTRIM(errarr(3)), 16, 'Cannot make connection')
     EXIT 
    ENDIF 

   ENDSCAN 
 ENDIF 
RETURN 
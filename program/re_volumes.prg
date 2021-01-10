PROCEDURE re_volumes
 IF MESSAGEBOX('ПЕРЕСЧИТАТЬ ОБЪЕМЫ?',4+32,'')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\nsif', 'nsif', 'shar', 'lpu_id')>0
  IF USED('nsif')
   USE IN nsif 
  ENDIF 
  USE IN aisoms 
  RETURN 
 ENDIF 

 SELECT nsif
 IF FIELD('n_kt')<>'N_KT'
  USE IN nsif 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\nsif', 'nsif', 'excl')>0
   IF USED('nsif')
    USE IN nsif 
   ENDIF 
   USE IN aisoms 
   RETURN 
  ELSE 
   SELECT nsif
   IF FIELD('n_kt_plan') != 'N_KT_PLAN'
    ALTER TABLE nsif ADD COLUMN n_kt_plan n(6)
   ENDIF 
   IF FIELD('n_kt') != 'N_KT' && факт
    ALTER TABLE nsif ADD COLUMN n_kt n(6)
   ENDIF 

   IF FIELD('n_ds_plan') != 'N_DS_PLAN' && факт
    ALTER TABLE nsif ADD COLUMN n_ds_plan n(6)
   ENDIF 
   IF FIELD('n_ds') != 'N_DS' && факт
    ALTER TABLE nsif ADD COLUMN n_ds n(6)
   ENDIF 
  
   IF FIELD('n_gem_plan') != 'N_GEM_PLAN' && факт
    ALTER TABLE nsif ADD COLUMN n_gem_plan n(6)
   ENDIF 
   IF FIELD('n_gem') != 'N_GEM' && факт
    ALTER TABLE nsif ADD COLUMN n_gem n(6)
   ENDIF 
  
   IF FIELD('n_eco_plan') != 'N_ECO_PLAN' && факт
    ALTER TABLE nsif ADD COLUMN n_eco_plan n(6)
   ENDIF 
   IF FIELD('n_eco') != 'N_ECO' && факт
    ALTER TABLE nsif ADD COLUMN n_eco n(6)
   ENDIF 
  
   IF FIELD('n_ks_plan') != 'N_KS_PLAN' && факт
    ALTER TABLE nsif ADD COLUMN n_ks_plan n(6)
   ENDIF 
   IF FIELD('n_ks') != 'N_KS' && факт
    ALTER TABLE nsif ADD COLUMN n_ks n(6)
   ENDIF 
   
   IF FIELD('prv') != 'PRV'
    ALTER TABLE nsif ADD COLUMN prv c(3)
   ENDIF 

   USE 
  
   IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\nsif', 'nsif', 'shar', 'lpu_id')>0
    IF USED('nsif')
     USE IN nsif 
    ENDIF 
    USE IN aisoms 
    RETURN 
   ENDIF 

  ENDIF 
 ENDIF 

 IF OpenFile(m.pCommon+'\gr_plan', 'gr_plan', 'shar', 'cod')>0
  IF USED('gr_plan')
   USE IN gr_plan
  ENDIF 
  USE IN nsif 
  USE IN aisoms 
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\profot', 'profot', 'shar', 'otd')>0
  IF USED('profot')
   USE IN profot
  ENDIF 
  USE IN gr_plan
  USE IN nsif 
  USE IN aisoms 
  RETURN 
 ENDIF 
 
 SELECT aisoms 
 SCAN 
  m.mcod  = mcod 
  m.lpuid = lpuid
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'err', 'shar', 'rid')>0
   IF USED('err')
    USE IN err 
   ENDIF 
   USE IN talon 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\hosp.dbf')
   IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\hosp', 'hosp', 'shar', 'c_i')>0
    IF USED('hosp')
     USE IN hosp 
    ENDIF 
   ENDIF 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  IF m.tmonth=1
  *IF m.tmonth=3
   UPDATE nsif SET n_ks=0, ks_fact=0, n_ds=0, ds_fact=0, app_fact=0, n_kt=0, ptkt_fact=0, ;
   	n_gem=0, gem_fact=0, n_eco=0, eco_fact=0 WHERE lpu_id=m.lpuid 
  ELSE 
   m.lcperiod = STR(m.tyear,4)+PADL(m.tmonth-1,2,'0')
   m.lppath = pbase+ '\'+m.lcperiod
   IF fso.FolderExists(m.lppath+'\nsi')
    IF fso.FileExists(m.lppath+'\nsi\nsif.dbf')
     IF OpenFile(m.lppath+'\nsi\nsif', 'p_nsif', 'shar', 'lpu_id')=0
      SELECT nsif 
      SET RELATION TO lpu_id INTO p_nsif
       REPLACE n_ks WITH p_nsif.n_ks, ks_fact WITH p_nsif.ks_fact, n_ds WITH p_nsif.n_ds, ds_fact WITH p_nsif.ds_fact, ;
       app_fact WITH p_nsif.app_fact, n_kt WITH p_nsif.n_kt, ptkt_fact WITH p_nsif.ptkt_fact, n_gem WITH p_nsif.n_gem, ;
       gem_fact WITH p_nsif.gem_fact, n_eco WITH p_nsif.n_eco, eco_fact WITH p_nsif.eco_fact FOR lpu_id = m.lpuid 
      SET RELATION OFF INTO p_nsif
      USE IN p_nsif
     ELSE 
      IF USED('p_nsif')
       USE IN p_nsif
      ENDIF 
      *UPDATE nsif SET ks_fact=0, ds_fact=0, app_fact=0, ptkt_fact=0, gem_fact=0, eco_fact=0 ;
  		WHERE lpu_id=m.lpuid 
      UPDATE nsif SET n_ks=0, ks_fact=0, n_ds=0, ds_fact=0, app_fact=0, n_kt=0, ptkt_fact=0, ;
   	   n_gem=0, gem_fact=0, n_eco=0, eco_fact=0 WHERE lpu_id=m.lpuid 
     ENDIF 
    ELSE 
     *UPDATE nsif SET ks_fact=0, ds_fact=0, app_fact=0, ptkt_fact=0, gem_fact=0, eco_fact=0 ;
  		WHERE lpu_id=m.lpuid
     UPDATE nsif SET n_ks=0, ks_fact=0, n_ds=0, ds_fact=0, app_fact=0, n_kt=0, ptkt_fact=0, ;
   	  n_gem=0, gem_fact=0, n_eco=0, eco_fact=0 WHERE lpu_id=m.lpuid 
    ENDIF 

    IF fso.FileExists(m.pBase+'\'+m.gcPeriod+'\nsi\nsif_r2.dbf')
     IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\nsif_r2', 'p_nsif', 'shar', 'lpu_id')=0
      SELECT nsif 
      SET RELATION TO lpu_id INTO p_nsif

      REPLACE n_ks WITH n_ks+p_nsif.n_ks, ks_fact WITH ks_fact+p_nsif.ks_fact, n_ds WITH n_ds+p_nsif.n_ds, ;
       ds_fact WITH ds_fact+p_nsif.ds_fact, app_fact WITH app_fact+p_nsif.app_fact, n_kt WITH n_kt+p_nsif.n_kt, ;
       ptkt_fact WITH ptkt_fact+p_nsif.ptkt_fact, n_gem WITH n_gem+p_nsif.n_gem, gem_fact WITH gem_fact+p_nsif.gem_fact,;
       n_eco WITH n_eco+p_nsif.n_eco, eco_fact WITH eco_fact+p_nsif.eco_fact FOR lpu_id = m.lpuid 

      SET RELATION OFF INTO p_nsif
      USE IN p_nsif
     ENDIF 
    ENDIF 

   ELSE 
    *UPDATE nsif SET ks_fact=0, ds_fact=0, app_fact=0, ptkt_fact=0, gem_fact=0, eco_fact=0 ;
  		WHERE lpu_id=m.lpuid
    UPDATE nsif SET n_ks=0, ks_fact=0, n_ds=0, ds_fact=0, app_fact=0, n_kt=0, ptkt_fact=0, ;
   	 n_gem=0, gem_fact=0, n_eco=0, eco_fact=0 WHERE lpu_id=m.lpuid 
   ENDIF 
  ENDIF 
  
  =SEEK(m.lpuid, 'nsif')
  
  CREATE CURSOR n_ks (c_i c(30))
  INDEX on c_i TAG c_i 
  SET ORDER TO c_i
  
  CREATE CURSOR n_ds (c_i c(30))
  INDEX on c_i TAG c_i 
  SET ORDER TO c_i

  CREATE CURSOR n_gem (c_i c(30))
  INDEX on c_i TAG c_i 
  SET ORDER TO c_i

  CREATE CURSOR n_eco (c_i c(30))
  INDEX on c_i TAG c_i 
  SET ORDER TO c_i
  
  m.nn_ks      = 0
  m.nks_fact   = 0 
  m.nn_ds      = 0
  m.nds_fact   = 0
  m.nn_gem     = 0 
  m.ngem_fact  = 0 
  m.nn_eco     = 0
  m.neco_fact  = 0
  m.nxt_fact   = 0 
  m.nlt_fact   = 0
  m.nvmp_fact  = 0 
  m.napp_fact  = 0 
  m.nptkt_fact = 0
  m.nn_kt      = 0 

  SELECT talon 
  SET RELATION TO recid INTO err
  SCAN 
   IF !EMPTY(err.c_err) AND err.c_err<>'PPA'
    LOOP 
   ENDIF 

   m.otd    = otd 
   m.cod    = cod
   m.usl_ok = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), profot.usl_ok, ' ')
   m.k_u    = k_u
   m.c_i    = c_i
   m.sn_pol = sn_pol
   m.d_u    = d_u
   m.s_all  = s_all+s_lek
   
   m.nsif = 0

   DO CASE 
    CASE m.usl_ok = '1' && стационар
     IF !SEEK(m.c_i, 'n_ks')
      INSERT INTO n_ks FROM MEMVAR 
      m.on_ks = nsif.n_ks
      m.nn_ks = m.on_ks + 1
      REPLACE n_ks WITH m.nn_ks IN nsif
     ENDIF 
     m.oks_fact = nsif.ks_fact
     m.nks_fact = m.oks_fact + m.s_all
     REPLACE ks_fact WITH m.nks_fact IN nsif
     m.nsif = 1
      
    CASE m.usl_ok = '2' && дневной стационар
     m.gr = IIF(SEEK(m.cod,'gr_plan'), gr_plan.gr_plan, '')
     IF !INLIST(m.gr, 'gem', 'eco')
      IF !SEEK(m.c_i, 'n_ds')
       INSERT INTO n_ds FROM MEMVAR 
       m.on_ds = nsif.n_ds
       m.nn_ds = m.on_ds + 1
       REPLACE n_ds WITH m.nn_ds IN nsif
      ENDIF 
      m.ods_fact = nsif.ds_fact
      m.nds_fact = m.ods_fact + m.s_all
      REPLACE ds_fact WITH m.nds_fact IN nsif
      m.nsif = 2
     ENDIF 
      
     IF m.gr = 'gem'
      IF !SEEK(m.c_i, 'n_gem')
       INSERT INTO n_gem FROM MEMVAR 
       m.on_gem = nsif.n_gem
       m.nn_gem = m.on_gem + 1
       REPLACE n_gem WITH m.nn_gem IN nsif
      ENDIF 
      m.ogem_fact = nsif.gem_fact
      m.ngem_fact = m.ogem_fact + m.s_all
      REPLACE gem_fact WITH m.ngem_fact IN nsif
      m.nsif = 4
     ENDIF 

     IF m.gr = 'eco'
      IF !SEEK(m.c_i, 'n_eco')
       INSERT INTO n_eco FROM MEMVAR 
       m.on_eco = nsif.n_eco
       m.nn_eco = m.on_eco + 1
       REPLACE n_eco WITH m.nn_eco IN nsif
      ENDIF 
      m.oeco_fact = nsif.eco_fact
      m.neco_fact = m.oeco_fact + m.s_all
      REPLACE eco_fact WITH m.neco_fact IN nsif
      m.nsif = 5
     ENDIF 
     
     IF m.gr = 'on_х'
      m.oxt_fact = nsif.xt_fact
      m.nxt_fact = m.oxt_fact + m.s_all
      REPLACE xt_fact WITH m.nxt_fact IN nsif
     ENDIF 

     IF m.gr = 'on_v'
      m.olt_fact = nsif.lt_fact
      m.nlt_fact = m.olt_fact + m.s_all
      REPLACE lt_fact WITH m.nlt_fact IN nsif
     ENDIF 

     IF FLOOR(m.cod/1000)=297
      m.ovmp_fact = nsif.vmp_fact
      m.nvmp_fact = m.ovmp_fact + m.s_all
      REPLACE vmp_fact WITH m.nvmp_fact IN nsif
     ENDIF 

    CASE m.usl_ok = '3' && АПП

     IF IsUsl(m.cod) AND IIF(USED('hosp'), SEEK(m.c_i, 'hosp'), .F.)
      m.oks_fact = nsif.ks_fact
      m.nks_fact = m.oks_fact + m.s_all
      REPLACE ks_fact WITH m.nks_fact IN nsif
      m.nsif = 1
     ELSE 
      m.gr = IIF(SEEK(m.cod,'gr_plan'), gr_plan.gr_plan, '')
      IF !INLIST(m.gr, 'kt')
       m.oapp_fact = nsif.app_fact
       m.napp_fact = m.oapp_fact + m.s_all
       REPLACE app_fact WITH m.napp_fact IN nsif
       m.nsif = 3
      ELSE 
       m.optkt_fact = nsif.ptkt_fact
       m.nptkt_fact = m.optkt_fact + m.s_all
       m.on_kt = nsif.n_kt
       m.nn_kt = m.on_kt + m.k_u
       REPLACE ptkt_fact WITH m.nptkt_fact, n_kt WITH m.nn_kt IN nsif
       m.nsif = 6
      ENDIF 
     ENDIF 
    OTHERWISE
   ENDCASE 
   
   REPLACE nsif WITH m.nsif

  ENDSCAN 
  SET RELATION OFF INTO err
  USE IN talon 
  USE IN err
  IF USED('hosp')
   USE IN hosp 
  ENDIF 
  USE IN n_ks
  USE IN n_ds
  USE IN n_gem
  USE IN n_eco
  
  WAIT CLEAR 
  
  SELECT aisoms 
  
 ENDSCAN 
 
 USE 
 USE IN nsif 
 USE IN gr_plan
 USE IN profot
 
 MESSAGEBOX('OK!',0+64,'')

RETURN 
PROCEDURE re_vols_s7
 IF MESSAGEBOX('œ≈–≈—◊»“¿“‹ Œ¡⁄≈Ã€?',4+32,'')=7
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
   IF FIELD('n_kt') != 'N_KT' && Ù‡ÍÚ
    ALTER TABLE nsif ADD COLUMN n_kt n(6)
   ENDIF 

   IF FIELD('n_ds_plan') != 'N_DS_PLAN' && Ù‡ÍÚ
    ALTER TABLE nsif ADD COLUMN n_ds_plan n(6)
   ENDIF 
   IF FIELD('n_ds') != 'N_DS' && Ù‡ÍÚ
    ALTER TABLE nsif ADD COLUMN n_ds n(6)
   ENDIF 
  
   IF FIELD('n_gem_plan') != 'N_GEM_PLAN' && Ù‡ÍÚ
    ALTER TABLE nsif ADD COLUMN n_gem_plan n(6)
   ENDIF 
   IF FIELD('n_gem') != 'N_GEM' && Ù‡ÍÚ
    ALTER TABLE nsif ADD COLUMN n_gem n(6)
   ENDIF 
  
   IF FIELD('n_eco_plan') != 'N_ECO_PLAN' && Ù‡ÍÚ
    ALTER TABLE nsif ADD COLUMN n_eco_plan n(6)
   ENDIF 
   IF FIELD('n_eco') != 'N_ECO' && Ù‡ÍÚ
    ALTER TABLE nsif ADD COLUMN n_eco n(6)
   ENDIF 
  
   IF FIELD('n_ks_plan') != 'N_KS_PLAN' && Ù‡ÍÚ
    ALTER TABLE nsif ADD COLUMN n_ks_plan n(6)
   ENDIF 
   IF FIELD('n_ks') != 'N_KS' && Ù‡ÍÚ
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
  
  IF m.tmonth=1
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
      UPDATE nsif SET n_ks=0, ks_fact=0, n_ds=0, ds_fact=0, app_fact=0, n_kt=0, ptkt_fact=0, ;
   	   n_gem=0, gem_fact=0, n_eco=0, eco_fact=0 WHERE lpu_id=m.lpuid 
     ENDIF 
    ELSE 
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
    UPDATE nsif SET n_ks=0, ks_fact=0, n_ds=0, ds_fact=0, app_fact=0, n_kt=0, ptkt_fact=0, ;
   	 n_gem=0, gem_fact=0, n_eco=0, eco_fact=0 WHERE lpu_id=m.lpuid 
   ENDIF 
  ENDIF 
  
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
  
  m.n_ks      = 0
  m.ks_fact   = 0 
  m.n_ds      = 0
  m.ds_fact   = 0
  m.n_gem     = 0 
  m.gem_fact  = 0 
  m.n_eco     = 0
  m.eco_fact  = 0
  m.xt_fact   = 0 
  m.lt_fact   = 0
  m.vmp_fact  = 0 
  m.app_fact  = 0 
  m.ptkt_fact = 0
  m.n_kt      = 0 

  SELECT talon 
  SET RELATION TO recid INTO err
  SCAN 
   IF !EMPTY(err.c_err) && AND err.c_err<>'PPA'
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
   m.tip    = tip
   
   m.nsif = 0
   
   DO CASE 
    CASE IsMes(m.cod) OR INLIST(FLOOR(m.cod/1000),200,300,397) OR INLIST(m.cod,56029,156003)&&  —
     IF !SEEK(m.c_i, 'n_ks')
      INSERT INTO n_ks FROM MEMVAR 
      m.n_ks = m.n_ks+1
     ENDIF 
     m.ks_fact = m.ks_fact+m.s_all
     m.nsif = 1
    
    CASE INLIST(FLOOR(m.cod/1000), 97,197,297) && ƒ—
     m.gr = IIF(SEEK(m.cod,'gr_plan'), gr_plan.gr_plan, '')
     *IF !INLIST(m.gr, 'gem', 'eco')
     IF !INLIST(m.gr, 'eco')
      IF !SEEK(m.c_i, 'n_ds')
       INSERT INTO n_ds FROM MEMVAR 
       m.n_ds = m.n_ds + 1
      ENDIF 
      m.ds_fact = m.ds_fact+m.s_all
      m.nsif = 2
     ENDIF 
      
     IF m.gr = 'gem'
      IF !SEEK(m.c_i, 'n_gem')
       INSERT INTO n_gem FROM MEMVAR 
       m.n_gem = m.n_gem +1 
      ENDIF 
      m.gem_fact = m.gem_fact + m.s_all
      m.nsif = 4
     ENDIF 

     IF m.gr = 'eco'
      IF !SEEK(m.c_i, 'n_eco')
       INSERT INTO n_eco FROM MEMVAR 
       m.n_eco = m.n_eco + 1
      ENDIF 
      m.eco_fact = m.eco_fact + m.s_all
      m.nsif = 5
     ENDIF 
     
     IF m.gr = 'on_ı'
      m.xt_fact = m.xt_fact + m.s_all
     ENDIF 

     IF m.gr = 'on_v'
      m.lt_fact = m.lt_fact + m.s_all
     ENDIF 

     IF FLOOR(m.cod/1000)=297
      m.vmp_fact = m.vmp_fact + m.s_all
     ENDIF 

    OTHERWISE && ¿œœ
     IF (USED('hosp') AND SEEK(m.c_i, 'hosp')) OR m.usl_ok='1' && INLIST(SUBSTR(m.otd,2,2),'70','73')
      m.ks_fact = m.ks_fact + m.s_all
      m.nsif = 1
     ELSE 
      m.gr = IIF(SEEK(m.cod,'gr_plan'), gr_plan.gr_plan, '')
      IF !INLIST(m.gr, 'kt')
       m.app_fact = m.app_fact + m.s_all
       m.nsif = 3
      ELSE 
       m.n_kt = m.n_kt + m.k_u
       m.ptkt_fact = m.ptkt_fact + m.s_all
       m.nsif = 6
      ENDIF 
     ENDIF 

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
  
  SELECT nsif 
  =SEEK(m.lpuid, 'nsif')
  
  m.on_ks      = n_ks
  m.oks_fact   = ks_fact
  m.on_ds      = n_ds
  m.ods_fact   = ds_fact
  m.on_gem     = n_gem
  m.ogem_fact  = gem_fact
  m.on_eco     = n_eco
  m.oeco_fact  = eco_fact
  m.oxt_fact   = xt_fact
  m.olt_fact   = lt_fact
  m.ovmp_fact  = vmp_fact
  m.oapp_fact  = app_fact
  m.optkt_fact = ptkt_fact
  m.on_kt      = n_kt

  REPLACE n_ks WITH m.on_ks+m.n_ks, ks_fact WITH m.oks_fact+m.ks_fact, n_ds WITH m.on_ds+m.n_ds, ds_fact WITH m.ods_fact+m.ds_fact, ;
  	n_gem WITH m.on_gem+m.n_gem, gem_fact WITH m.ogem_fact+m.gem_fact, n_eco WITH m.on_eco+m.n_eco, eco_fact WITH m.oeco_fact+m.eco_fact, ;
  	xt_fact WITH m.oxt_fact+m.xt_fact, lt_fact WITH m.olt_fact+m.lt_fact, vmp_fact WITH m.ovmp_fact+m.vmp_fact, ;
  	app_fact WITH m.oapp_fact+m.app_fact, ptkt_fact WITH m.optkt_fact+m.ptkt_fact, n_kt WITH m.on_kt+m.n_kt
  	
  WAIT CLEAR 
  
  SELECT aisoms 
  
 ENDSCAN 
 
 SUM (s_pred+s_lek)-(sum_flk+ls_flk) TO m.s_ais
 USE 
 
 SELECT nsif
 SUM app_fact+ks_fact+ds_fact+eco_fact+ptkt_fact TO m.s_nsif
 USE
 
 USE IN gr_plan
 USE IN profot
 
 IF m.s_ais=m.s_nsif
  MESSAGEBOX('OK!',0+64,'')
 ELSE 
  MESSAGEBOX('œŒ ¿«¿“≈À» –¿—◊»“¿Õ€ Õ≈¬≈–ÕŒ!',0+64,'')
 ENDIF 
 

RETURN 
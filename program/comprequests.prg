PROCEDURE CompRequests
 IF MESSAGEBOX('—–¿¬Õ»“‹'+CHR(13)+CHR(10)+'«¿œ–Œ—€?',4+32,'ÕŒ¬¿ﬂ ¬≈–—»ﬂ!')=7
  RETURN 
 ENDIF 
 
 snddir = ''
 snddir = GETDIR(pBase, '” ¿∆»“≈ Õ¿ ƒ»–≈“Œ–»ﬁ BASE', '¡¿«¿ — ¿»— «¿œ–Œ—¿Ã»', 0)
 IF EMPTY(snddir)
  MESSAGEBOX(CHR(13)+CHR(10)+'¬€ Õ»◊≈√Œ Õ≈ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 snddir = snddir + '\' + m.gcPeriod
 IF !fso.FolderExists(snddir)
  MESSAGEBOX('ƒ»–≈ “Œ–»ﬂ ' + snddir + ' Õ≈ Õ¿…ƒ≈Õ¿!'+CHR(13)+CHR(10), 0+16, '')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(snddir+'\aisoms.dbf')
  MESSAGEBOX('‘¿…À ' + snddir+'\aisoms.dbf' + ' Õ≈ Õ¿…ƒ≈Õ!'+CHR(13)+CHR(10), 0+16, '')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  MESSAGEBOX('‘¿…À ' + pbase+'\'+m.gcperiod+'\aisoms.dbf' + ' Õ≈ Õ¿…ƒ≈Õ!'+CHR(13)+CHR(10), 0+16, '')
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR diff_q (mcod c(7), recid c(6), recid2 c(6), s_pol c(6), s_pol2 c(6), n_pol c(16), n_pol2 c(16), ;
 	d_u c(8), q c(2), q2 c(2), fam c(25), fam2 c(25), im c(20), im2 c(20), ot c(20), ot2 c(20), ;
 	dr c(8), dr2 c(8), w n(1), w2 n(1), ans_r c(3), ans_r2 c(3), tip_d c(1), tip_d2 c(1), ;
 	lpu_id n(6), lpu_id2 n(6), st_id n(6), d_rq d, d_end d, err n(1))
 CREATE CURSOR diff_lpu_id (mcod c(7), recid c(6), recid2 c(6), s_pol c(6), s_pol2 c(6), n_pol c(16), n_pol2 c(16), ;
 	d_u c(8), q c(2), q2 c(2), fam c(25), fam2 c(25), im c(20), im2 c(20), ot c(20), ot2 c(20), ;
 	dr c(8), dr2 c(8), w n(1), w2 n(1), ans_r c(3), ans_r2 c(3), tip_d c(1), tip_d2 c(1), ;
 	lpu_id n(6), lpu_id2 n(6), st_id n(6), d_rq d, d_end d, err n(1))

 SELECT aisoms 
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\soapans.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 

  IF !fso.FolderExists(snddir+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(snddir+'\'+m.mcod+'\answer.dbf')
   LOOP 
  ENDIF 

  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\soapans', 'soapans', 'excl')>0
   IF USED('soapans')
    USE IN soapans
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(snddir+'\'+m.mcod+'\answer', 'answer', 'excl')>0
   IF USED('answer')
    USE IN answer
   ENDIF 
   USE IN soapans
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(snddir+'\'+m.mcod+'\people', 'people', 'shar', 'recid')>0
   IF USED('people')
    USE IN people
   ENDIF 
   USE IN soapans
   USE IN answer
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod WINDOW NOWAIT 
  
  SELECT soapans
  INDEX on n_pol TAG n_pol
  SET ORDER TO n_pol
  SELECT answer
  INDEX on INT(VAL(recid)) TAG recid
  SET ORDER TO recid

  SELECT people 
  SET RELATION TO recid INTO answer
  SET RELATION TO LEFT(sn_pol,16) INTO soapans ADDITIVE 
  SCAN 

   recid   = soapans.recid
   s_pol   = soapans.s_pol
   n_pol   = soapans.n_pol
   q       = soapans.q
   fam     = soapans.fam
   im      = soapans.im
   ot      = soapans.ot
   dr      = soapans.dr
   w       = soapans.w
   ans_r   = soapans.ans_r
   tip_d   = soapans.tip_d
   lpu_id  = soapans.lpu_id

   recid2  = answer.recid
   s_pol2  = answer.s_pol
   n_pol2  = answer.n_pol
   q2      = answer.q
   fam2    = answer.fam
   im2     = answer.im
   ot2     = answer.ot
   dr2     = answer.dr
   w2      = answer.w
   ans_r2  = answer.ans_r
   tip_d2  = answer.tip_d
   lpu_id2 = answer.lpu_id
   
   IF m.q != m.q2
    INSERT INTO diff_q FROM MEMVAR 
   ENDIF 
   IF m.lpu_id != m.lpu_id2
    INSERT INTO diff_lpu_id FROM MEMVAR 
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO answer
  SET RELATION OFF INTO soapans
  USE 
  SELECT soapans
  DELETE TAG ALL 
  USE
  SELECT answer 
  DELETE TAG ALL 
  USE 
  WAIT CLEAR 
  SELECT aisoms
 ENDSCAN 
 USE IN aisoms 
 
 WAIT "—Œ’–¿Õ≈Õ»≈ –≈«”À‹“¿“Œ¬..." WINDOW NOWAIT 
 SELECT diff_q
 COPY TO &pBase/&gcPeriod/diff_q
 SELECT diff_lpu_id
 COPY TO &pBase/&gcPeriod/diff_lpu_id
 WAIT CLEAR 
 
 MESSAGEBOX('Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!',0+64,'')

RETURN 
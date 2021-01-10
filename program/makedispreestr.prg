PROCEDURE MakeDispReestr
 IF MESSAGEBOX('ÑÎÁÐÀÒÜ ÄÈÑÏÀÍÑÅÐÍÛÉ ÐÅÅÑÒÐ?',4+32,'Äèñïàíñåðèçàöèÿ')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisom 
  ENDIF 
  RETURN 
 ENDIF 
 
 IF m.tMonth=1 
  CREATE CURSOR dsp (sn_pol c(25), w n(1), ages n(3), dn n(1), start d, dn_end n(1), end d, ds c(6))
  SELECT dsp 
  INDEX ON sn_pol+ds TAG unik
  SET ORDER TO unik
 ELSE 
  IF fso.FileExists(m.pBase+'\'+STR(m.tYear,4)+PADL(m.tMonth-1,2,'0')+'\dsp_r.dbf')
   IF OpenFile(m.pBase+'\'+STR(m.tYear,4)+PADL(m.tMonth-1,2,'0')+'\dsp_r', 'dsp', 'shar', 'unik')>0
    IF USED('dsp')
     USE IN dsp 
    ENDIF 
    CREATE CURSOR dsp (sn_pol c(25), w n(1), ages n(3), dn n(1), start d, dn_end n(1), end d, ds c(6))
    SELECT dsp 
    INDEX ON sn_pol+ds TAG unik
    SET ORDER TO unik
   ENDIF 
  ELSE 
   CREATE CURSOR dsp (sn_pol c(25), w n(1), ages n(3), dn n(1), start d, dn_end n(1), end d, ds c(6))
   SELECT dsp 
   INDEX ON sn_pol+ds TAG unik
   SET ORDER TO unik
  ENDIF 
 ENDIF 
 
 SELECT aisoms 
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar')>0
   USE IN people 
   IF USED('talon')
    USE IN talon 
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT talon 
  SET RELATION TO sn_pol INTO people 
  SCAN 
   m.dn = dn
   IF m.dn=0
    LOOP 
   ENDIF 
   
   m.d_u   = d_u
   m.start = {}
   m.end   = {}
   
   m.sn_pol = sn_pol
   m.ds   = ds
   m.w    = people.w
   m.dr   = people.dr
   m.adj  = CTOD(STRTRAN(DTOC(m.dr), STR(YEAR(m.dr),4), STR(YEAR(people.d_beg),4)))-people.d_beg
   m.ages = (YEAR(people.d_beg) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)
   
   IF INLIST(m.dn,1,2)
    m.start = m.d_u
   ELSE 
    m.dn_end = m.dn
    m.dn     = 0
    m.end    = m.d_u
   ENDIF 

   m.vir     = m.sn_pol + m.ds

   IF !SEEK(m.vir, 'dsp')
    INSERT INTO dsp FROM MEMVAR 
   ELSE 
    IF INLIST(m.dn,3,4)
     REPLACE end WITH m.d_u, dn_end WITH m.dn IN dsp
    ELSE 
     REPLACE start WITH m.d_u, dn WITH m.dn IN dsp
    ENDIF 
   ENDIF 
   
  ENDSCAN 
  SET RELATION OFF INTO people 
  USE IN people 
  USE IN talon 
  SELECT aisoms 
  
  WAIT CLEAR 
  
 ENDSCAN 
 USE IN aisoms 
 
 SELECT dsp
 COPY TO &pBase\&gcPeriod\dsp_r WITH cdx 
 USE 
 
 MESSAGEBOX('OK!',0+64,'')

RETURN 
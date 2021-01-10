PROCEDURE RestEFls2
 * Îòëè÷àåòñÿ îò ïåðâîé âåðñèè RestEFls àäðåñîì ïîèñêà ïåðñîò÷åòà
 * Çäåñü - â äèðåêòîðèè exchange.dir

 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÂÎÑÑÒÀÍÎÂÈÒÜ ÔÀÉËÛ ÎØÈÁÎÊ'+CHR(13)+CHR(10)+;
  'ÏÎ ÔÀÉËÀÌ ÏÅÐÑÎÒ×ÅÒÀ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pExpImp+'\r.dbf')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË '+UPPER(pExpImp)+'\r.dbf!', 0+16, '')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pExpImp+'\s.dbf')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË '+UPPER(pExpImp)+'\s.dbf!', 0+16, '')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pExpImp+'\c.dbf')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË '+CHR(13)+CHR(10)+UPPER(pExpImp)+'\c.dbf!', 0+16, '')
  RETURN 
 ENDIF 
 
 IF OpenFile(pExpImp+'\c.dbf', 'cf', 'shar')>0
  IF USED('cf')
   USE IN cf
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  USE IN cf
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT lpu_id, SPACE(7) as mcod FROM cf GROUP BY lpu_id INTO CURSOR c_cf READWRITE 
 SELECT c_cf 
 SET RELATION TO lpu_id INTO sprlpu
 REPLACE ALL mcod WITH sprlpu.mcod
 SET RELATION OFF INTO sprlpu
 USE IN sprlpu 
 SELECT c_cf 
 INDEX on mcod TAG mcod 
 SET ORDER TO mcod 

 SCAN 
  m.mcod   = mcod
  m.lpu_id = lpu_id
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people', 'people', 'shar', 'recid_lpu')>0
   IF USED('people')
    USE IN people
   ENDIF 
   SELECT c_cf
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid_lpu')>0
   USE IN people
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT c_cf
   LOOP 
  ENDIF 
  IF OpenFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\e'+m.mcod, 'err', 'excl')>0
   USE IN talon
   USE IN people
   IF USED('err')
    USE IN err
   ENDIF 
   SELECT c_cf
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT err
  ZAP 
  
  SELECT * FROM cf WHERE lpu_id = m.lpu_id ORDER BY er_f INTO CURSOR mo_err
  SELECT mo_err
  SCAN 
   m.f     = er_f
   m.c_err = er_c
   m.rid   = 0
   m.recid_lpu = recid
   IF m.f = 'R'
    m.rid = IIF(SEEK(m.recid_lpu, 'people'), people.recid, 0)
   ELSE 
    m.rid = IIF(SEEK(m.recid_lpu, 'talon'), talon.recid, 0)
   ENDIF 
   
   IF m.rid>0
    INSERT INTO err FROM MEMVAR 
   ENDIF 
   
  ENDSCAN 
  USE IN mo_err
  
  USE IN people
  USE IN talon
  USE IN err
  
  SELECT c_cf
  
  WAIT CLEAR 
  
 ENDSCAN 

 USE IN c_cf 
 
 USE IN cf
 
 MESSAGEBOX('OK!',0+64,'')
 
 RETURN
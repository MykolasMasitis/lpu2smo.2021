PROCEDURE KillBadEkmp
 IF MESSAGEBOX('сдюкхрэ меопнтхкэмше детейрш щйло?',4+32,'')=7
  RETURN 
 ENDIF 
 IF MESSAGEBOX('бш сбепемш?',4+32,'')=7
  RETURN 
 ENDIF 
 
 IF OpenFile(pCommon+'\explist', 'explist', 'shar', 'cod')>0
  IF USED('')
   USE IN 
  ENDIF 
  RETURN 
 ENDIF 
 
 IF FIELD('prv', 'explist')!='PRV'
  USE IN explist 
  MESSAGEBOX('б яопюбнвмхйе COMMON\EXPLIST.DBF'+CHR(13)+CHR(10)+;
  	'нрясрярбсер онке PRV C(3)'+CHR(13)+CHR(10)+'опнднкфемхе пюанрш мебнглнфмн!',0+16,'')
  RETURN 
 ENDIF 
 
 SELECT explist
 COUNT FOR !EMPTY(prv) TO m.nrecs
 
 IF m.nrecs <= 0
  USE IN explist 
  MESSAGEBOX('б яопюбнвмхйе COMMON\EXPLIST.DBF'+CHR(13)+CHR(10)+;
  	'ме гюонкмемн онке PRV C(3)'+CHR(13)+CHR(10)+'опнднкфемхе пюанрш мебнглнфмн!',0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  USE IN explist
  RETURN 
 ENDIF 
 
 CREATE CURSOR currs (mcod c(7), nrecs n(6))
 
 SELECT aisoms
 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\m'+m.mcod, 'err', 'shar')>0
   IF USED('err')
    USE IN err 
   ENDIF 
   SELECT aisoms 
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+m.gcPeriod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
   IF USED('err')
    USE IN err 
   ENDIF 
   SELECT aisoms 
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  m.nrecs = 0 
  SELECT err
  SET RELATION TO recid INTO talon 
  SET RELATION TO docexp INTO explist ADDITIVE
*  RECALL ALL 
  SCAN 
   m.et = et 
   IF !INLIST(m.et,'4','5','6','9')
    LOOP 
   ENDIF 
   IF EMPTY(explist.cod) 
    LOOP 
   ENDIF 
   IF EMPTY(explist.prv)
    LOOP 
   ENDIF 
   IF EMPTY(talon.profil)
    LOOP 
   ENDIF 
   
   IF talon.profil!=explist.prv
    m.nrecs = m.nrecs + 1
    DELETE 
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO talon 
  SET RELATION OFF INTO explist 
  IF m.nrecs>0
   INSERT INTO currs FROM MEMVAR 
  ENDIF 
  
  USE IN err 
  USE IN talon 
  SELECT aisoms 
  
  WAIT CLEAR 
  
 ENDSCAN 
 USE IN aisoms 
 USE IN explist
 
 SELECT currs
 IF RECCOUNT('currs')>0
  COPY TO &pbase\&gcperiod\delerrsekmp
  USE 
  MESSAGEBOX('напюанрйю гюйнмвемю!'+CHR(13)+CHR(10)+'пегскэрюрш бш лнфере онялнрперэ'+CHR(13)+CHR(10)+;
 	'б тюике '+UPPER(pbase+'\'+m.gcperiod)+'\DelErrsEKMP.dbf', 0+64, '')
 ELSE 
  MESSAGEBOX('напюанрйю гюйнмвемю!', 0+64, '')
 ENDIF 

RETURN 
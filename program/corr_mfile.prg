PROCEDURE Corr_mfile(lcPath)
 IF MESSAGEBOX('унрхре нрйнппейрхпнбюрэ M-тюик?',4+32,'')=7
  RETURN 
 ENDIF 

 oal     = ALIAS() && Aisoms
 m.mcod  = SUBSTR(lcpath,RAT('\',lcpath)+1)

 IF OpenFile(lcPath+'\talon', 'talon', 'shar', 'recid')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  SELECT (oal)
  RETURN 
 ENDIF 
 IF OpenFile(lcPath+'\m'+m.mcod, 'merror', 'shar')>0
  IF USED('merror')
   USE IN merror
  ENDIF 
  IF USED('talon')
   USE IN talon 
  ENDIF 
  SELECT (oal)
  RETURN 
 ENDIF 
 
 SELECT merror
 SET RELATION TO recid INTO talon 
 m.defs = 0 
 m.ok   = 0 
 SCAN 
  m.n_akt = ALLTRIM(n_akt)
  IF !EMPTY(m.n_akt)
   LOOP 
  ENDIF 
  m.defs = m.defs + 1
  m.err_mee = err_mee
  m.sn_pol  = talon.sn_pol

  m.et     = et
  m.docexp = docexp
  m.usr    = usr
  m.reason = reason
  
  m.IsSV = IIF(m.err_mee='W0', .T., .F.)
  m.IsSS = !m.IsSV
  
  m.IsMEE = IIF(INLIST(m.et,'2','3','7','8'), .T., .F.)
  m.IsEKMP = !m.IsMEE
  
  IF m.IsSV
   IF m.IsMEE
    SELECT n_akt, actdate as d_akt FROM svacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.et)) ;
    	AND smoexp=m.usr AND reason=m.reason INTO CURSOR rqwest NOCONSOLE  
   ELSE 
    SELECT n_akt, actdate as d_akt FROM svacts WHERE period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.et)) ;
    	AND docexp=m.docexp INTO CURSOR rqwest NOCONSOLE  
   ENDIF 
   m.n_akt = ''
   m.d_akt = {}
   IF RECCOUNT('rqwest')>0
    m.n_akt = rqwest.n_akt
    m.d_akt = TTOD(rqwest.d_akt)
   ENDIF 
   USE IN rqwest
   SELECT merror 
   IF !EMPTY(m.n_akt)
    REPLACE n_akt WITH m.n_akt, d_akt WITH m.d_akt, t_akt WITH IIF(m.IsSV, 'SV', 'SS')
    m.ok = m.ok + 1
   ENDIF 

  ELSE 
   IF m.IsMEE
    SELECT n_akt, actdate as d_akt FROM ssacts WHERE period=m.gcperiod AND mcod=m.mcod AND ;
    	codexp=INT(VAL(m.et)) AND sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) ;
  		INTO CURSOR rqwest NOCONSOLE 
   ELSE 
    SELECT n_akt, actdate as d_akt FROM ssacts WHERE ;
    	period=m.gcperiod AND mcod=m.mcod AND codexp=INT(VAL(m.et)) AND ;
    	sn_pol=PADR(STRTRAN(m.sn_pol,' ',''),25) AND docexp=m.docexp AND reason=m.reason ;
    	INTO CURSOR rqwest NOCONSOLE 
   ENDIF 
   m.n_akt = ''
   m.d_akt = {}
   IF RECCOUNT('rqwest')>0
    m.n_akt = rqwest.n_akt
    m.d_akt = TTOD(rqwest.d_akt)
   ENDIF 
   USE IN rqwest
   SELECT merror 
   IF !EMPTY(m.n_akt)
    REPLACE n_akt WITH m.n_akt, d_akt WITH m.d_akt, t_akt WITH IIF(m.IsSV, 'SV', 'SS')
    m.ok = m.ok + 1
   ENDIF 

  ENDIF 
  
  
 ENDSCAN 
 SET RELATION OFF INTO talon 
 USE IN talon 
 USE IN merror 
 
 SELECT (oal)
 
 MESSAGEBOX('намюпсфемн '+TRANSFORM(m.defs,'99999')+' осяршу юйрнб!'+CHR(13)+CHR(10)+;
 	'бняярюмнбкемн '+TRANSFORM(m.ok,'99999')+' !'+CHR(13)+CHR(10),0+64,'')
 
RETURN 
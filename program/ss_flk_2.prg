PROCEDURE ss_flk_2
 * 2-ой этап МЭК
 * UMA - gosp создается из меню
 * MPA - deads создается из меню
 * D2A, D6A - disp
 * DKA - disp
 M.UMA = .T.
 M.MPA = .T.
 M.D6A = .T.
 M.DKA = .T.
 M.HNA = .T. && Центры здоровья
 M.DNA = .T.
 
 IF M.UMA = .T.
  IF USED('gosp')
   m.otd    = otd
   m.sn_pol = sn_pol
   m.cod    = cod
   m.d_u    = d_u
   m.ds     = ds
   m.d_type = d_type
   m.recid  = recid
   m.c_i    = c_i
     
   SET ORDER TO sn_pol IN gosp
   IF SEEK(m.sn_pol, 'gosp', 'sn_pol')
    DO WHILE gosp.sn_pol = m.sn_pol
     DO CASE 
      *CASE m.recid = gosp.recid
      CASE IsMes(m.cod) OR IsVmp(m.cod)
      CASE IsMes(gosp.cod) AND INLIST(m.cod,1781,101781) AND m.mcod=gosp.mcod
      CASE IsMes(gosp.cod) AND INLIST(INT(m.cod/1000), 29,129)
      CASE m.lpuid = 1874 AND IsMes(gosp.cod) AND INLIST(INT(m.cod/1000),138)
      CASE IsMes(gosp.cod) AND m.d_type='s'
      CASE IsMes(gosp.cod) AND INLIST(INT(m.cod/1000), 49,149) AND m.d_type='2'
      CASE IsMes(gosp.cod) AND INLIST(m.cod,97010,197010) AND m.d_type='2'
      CASE IsMes(gosp.cod) AND m.d_type='w'
      *CASE IsMes(gosp.cod) AND INLIST(m.cod,36022,136022,36023,136023) AND m.d_type='2'
      CASE IsMes(Gosp.cod) AND INLIST(m.cod,36022,136022,36023,136023,36024,136024) AND m.d_type='2'
      CASE IsMes(gosp.cod) AND INLIST(INT(m.cod/1000),84)
        CASE IsMes(Gosp.cod) AND INLIST(INT(m.cod/1000), 297) AND m.lpuid=1872  && Лучевая терапия в Морозовской
        CASE INLIST(Gosp.cod,200518, 200519, 200520, 200522, 200523, 200524) AND ;
        	INLIST(m.cod,49011, 49012, 49013, 49024, 49030, 49035) AND m.d_type='2' && гемодиализ во время ВМП
      CASE OCCURS('#', m.c_i)=3
        
      CASE INLIST(INT(m.cod/1000),25,125,26,126,27,127,28,128,29,129,30,130) AND m.mcod<>gosp.mcod
      CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod=gosp.mcod
      
      OTHERWISE 
       IF gosp.k_u>1
        IF BETWEEN(m.d_u, gosp.d_u-gosp.k_u+1, gosp.d_u-1)
         m.recid = recid
         rval = InsError('S', 'UMA', m.recid, '1',;
       	  'Оказание услуг в период лечения по МЭС/СКП в МО '+gosp.mcod+;
      	  ', МЭС '+PADL(gosp.cod,6)+', d_u '+DTOC(gosp.d_u)+', k_u '+STR(gosp.k_u,3)+', (второй этап МЭК)')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       ENDIF 
      ENDCASE 
       
     SKIP IN gosp
    ENDDO 
   ENDIF 

  ENDIF 
 ENDIF 

 IF M.MPA == .T.
  IF USED('deads')
   m.cod    = cod
   m.c_i    = c_i
   IF SEEK(LEFT(m.sn_pol,17), 'deads') AND OCCURS('#', m.c_i)<3
    m.cod    = cod 
    m.d_date = deads.d_u
    IF m.d_u>m.d_date
    DO CASE 
     CASE INLIST(INT(m.cod/1000),59,159) AND m.d_u - m.d_date < 60
     CASE INLIST(INT(m.cod/1000),25,125,26,126,27,127,29,129,30,130) AND m.d_u - m.d_date < 60
     CASE INLIST(INT(m.cod/1000),28,128) AND m.d_u - m.d_date<60 AND m.lpuid=1795
     OTHERWISE 
      m.recid = recid
      rval =InsError('S', 'MPA', m.recid, '1', ;
    		'Услуга оказана после смерти пациента: mcod='+deads.mcod+', d_u=)' + DTOC(deads.d_u) + ;
    		', tip='+deads.tip+', d_type='+deads.d_type+', (второй этап МЭК)')
   	  m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDCASE 
    ENDIF 
   ENDIF
  ENDIF 
 ENDIF 

 IF M.D6A == .T.
    m.cod = cod
    m.dsptip = IIF(SEEK(m.cod,'dspcodes'), dspcodes.tip, 0)
    m.dsptip = IIF(!INLIST(m.cod, 25204, 35401), m.dsptip, 0)
    
    m.d_u = d_u
    m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.vozr = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)
    
    IF m.dsptip = 1
     *m.perem = m.mcod+LEFT(sn_pol,17)+STR(2,1)
     m.perem = LEFT(sn_pol,17)+STR(2,1)
     IF SEEK(m.perem, 'disp')
      *DO WHILE disp.mcod+LEFT(disp.sn_pol,17)+STR(disp.tip,1) = m.perem
      DO WHILE LEFT(disp.sn_pol,17)+STR(disp.tip,1) = m.perem
       IF YEAR(m.d_u) = YEAR(disp.d_u) AND EMPTY(disp.er)
        m.recid = recid
        m.cmnt = 'Застрахованному '+ALLTRIM(sn_pol)+' в том же году ('+DTOC(disp.d_u)+') была оказана услуга '+PADL(disp.cod,6,'0')+;
      	' (Профосмотр взрослого населения) в МО: '+disp.mcod
        rval = InsError('S', 'D6A', m.recid, '1', m.cmnt+', (второй этап МЭК)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      SKIP IN disp
      ENDDO 
     ENDIF  
    ENDIF 
    
    IF m.dsptip = 2
     *m.perem = m.mcod+LEFT(sn_pol,17)+STR(2,1)
     m.perem = LEFT(sn_pol,17)+STR(1,1)
     IF SEEK(m.perem, 'disp')
      *DO WHILE disp.mcod+LEFT(disp.sn_pol,17)+STR(disp.tip,1) = m.perem
      DO WHILE LEFT(disp.sn_pol,17)+STR(disp.tip,1) = m.perem
       IF YEAR(m.d_u) = YEAR(disp.d_u) AND EMPTY(disp.er)
        m.recid = recid
        m.cmnt = 'Застрахованному '+ALLTRIM(sn_pol)+' в том же году ('+DTOC(disp.d_u)+') была оказана услуга '+PADL(disp.cod,6,'0')+;
      	' (Профосмотр взрослого населения) в МО: '+disp.mcod
        rval = InsError('S', 'D6A', m.recid, '1', m.cmnt+', (второй этап МЭК)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      SKIP IN disp
      ENDDO 
     ENDIF  
    ENDIF 
    
    IF m.dsptip = 3
     *m.perem = m.mcod+LEFT(sn_pol,17)+STR(4,1)
     m.perem = LEFT(sn_pol,17)+STR(4,1)
     IF SEEK(m.perem, 'disp')
      *DO WHILE disp.mcod+LEFT(disp.sn_pol,17)+STR(disp.tip,1) = m.perem
      DO WHILE LEFT(disp.sn_pol,17)+STR(disp.tip,1) = m.perem
       IF YEAR(m.d_u) = YEAR(disp.d_u)  AND EMPTY(disp.er)
        m.recid = recid
        m.cmnt = 'Застрахованному '+ALLTRIM(sn_pol)+' в том же году ('+DTOC(disp.d_u)+') была оказана услуга '+PADL(disp.cod,6,'0')+;
      	' (Профосмотр взрослого населения) в МО: '+disp.mcod
        rval = InsError('S', 'D6A', m.recid, '1', m.cmnt+', (второй этап МЭК)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      SKIP IN disp
      ENDDO 
     ENDIF  
    ENDIF 
    

    IF m.dsptip > 0 && AND !INLIST(m.dsptip,2,4)
    
    DO CASE 
     CASE m.dsptip = 1  && Диспансеризция взрослых, первый этап, tip=1, 1925-1935
      m.lastt = IIF(m.vozr<40, dspcodes.last, 12)
     CASE m.dsptip = 2 && ПМО взрослых, tip=2, 1921-1924
      m.lastt = dspcodes.last
     CASE m.dsptip = 3 && Диспансеризация детей, tip=3, осталась только одна - 10952
      m.lastt = dspcodes.last
     CASE m.dsptip = 4 && ПМО детей, tip=4, 101933-101951
      m.lastt = dspcodes.last
     CASE m.dsptip = 5 && Предварительные, tip=5 - больше их нет
      m.lastt = dspcodes.last
     CASE m.dsptip = 6 && Периодические, tip=6 - больше их нет
      m.lastt = dspcodes.last
     OTHERWISE 
      m.lastt = 0
    ENDCASE
    
    IF m.lastt>0
     m.perem = LEFT(sn_pol,17)+PADL(m.dsptip,1,'0')

     IF SEEK(m.perem, 'disp')
      DO WHILE LEFT(disp.sn_pol,17)+PADL(disp.tip,1,'0') = m.perem
       *IF IIF(m.lastt>=12, YEAR(m.d_u)-YEAR(disp.d_u)<m.lastt/12, (m.d_u - disp.d_u)/30<m.lastt) AND EMPTY(disp.er)
       IF (m.d_u>disp.d_u OR (m.d_u=disp.d_u AND m.mcod<>disp.mcod)) AND EMPTY(disp.er)
         m.recid = recid
         m.cmnt = 'Застрахованному '+ALLTRIM(sn_pol)+' ранее ('+DTOC(disp.d_u)+') оказана услуга '+PADL(disp.cod,6,'0')+;
      	 ' из той же категории, возраст: '+STR(m.vozr,2)+' лет'
         rval = InsError('S', IIF(INLIST(m.dsptip,2,4), 'DKA', 'D2A'), m.recid, '1', m.cmnt+', (второй этап МЭК)')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       SKIP IN disp
      ENDDO 
     ENDIF  

    ENDIF 
    
   ENDIF && IF m.dsptip>0
  
 ENDIF && IF M.D6A == .T.

 IF M.DKA == .T. AND USED('disp')
  m.cod = cod
  m.dsptip = IIF(SEEK(m.cod,'dspcodes'), dspcodes.tip, 0)

  IF !(INLIST(m.dsptip,2,4) OR INLIST(m.cod,101927,101928))
  ELSE 
    
   m.d_u = d_u
   m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
   m.vozr = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)
    
   m.k_key = m.mcod + LEFT(m.sn_pol,17) + PADL(m.cod,6,"0")
    
   IF SEEK(m.k_key, 'disp', 'un_tag')
    m.dd_u = disp.d_u
    IF INLIST(m.cod, 101937, 101945)
     IF m.d_u - m.dd_u < 365 AND m.dd_u >= CTOD('01.01.'+STR(tYear,4))
      m.recid = recid
      rval = InsError('S', 'DKA', m.recid, '1',;
       	'Комплексная услуга профнаправления '+STR(m.cod,6)+' оказывалась ранее, -'+DTOC(disp.d_u)+', (второй этап МЭК)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ELSE 
     IF disp.d_u >= CTOD('01.01.'+STR(tYear,4))
      m.recid = recid
      rval = InsError('S', 'DKA', m.recid, '1',;
      	'Комплексная услуга профнаправления '+STR(m.cod,6)+' оказывалась ранее, -'+DTOC(disp.d_u)+', (второй этап МЭК)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 
  ENDIF 
 ENDIF && IF M.DKB == .T.
   
 IF M.HNA == .T.
  m.cod    = cod
  m.sn_pol = sn_pol
  m.d_u    = d_u
  IF INLIST(m.cod,15001,115001) AND SEEK(m.sn_pol, 'polic_h') AND m.d_u > polic_h.d_u
   m.pr_mcod = polic_h.mcod
   m.pr_d_u  = polic_h.d_u
   m.recid = recid
   rval = InsError('S', 'HNA', m.recid, '1',;
   	'Пациент ранее обращался в ЦЗ МО '+m.pr_mcod+' '+DTOC(m.pr_d_u)+', (второй этап МЭК)')
   m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
  ENDIF 
  RELEASE cod, sn_pol, d_u
 ENDIF 

 IF M.DNA = .T. && .T.
  m.cod  = cod 
  m.rslt = rslt
  m.d_u = d_u
  DO CASE 
   CASE INLIST(m.cod, 101003, 101028, 101030) AND INLIST(m.rslt,321,322,323,324,325)
    m.perem = LEFT(sn_pol,17) + '3' && ищем услугу 101952 - в этой категории только одна осталась
    IF SEEK(m.perem, 'dspp')
     m.is_ok = .F.
     DO WHILE LEFT(dspp.sn_pol,17)+'3' = m.perem
      IF m.d_u-dspp.d_u<=60 AND EMPTY(dspp.er) AND INLIST(dspp.rslt,365,366,367,368)
       m.is_ok = .T.
       EXIT 
      ENDIF 
     SKIP IN dspp
     ENDDO 
     IF m.is_ok = .F.
      m.recid = recid
      rval = InsError('S', 'DNA', m.recid, '1',;
    		'Услуга 101952 с результатом 365/366/367/368 не было оказано ранее')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ELSE && не было такой услуги
     m.recid = recid
     rval = InsError('S', 'DNA', m.recid, '1',;
    	'Услуга 101952 с результатом 365/366/367/368 не было оказано ранее')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF  

   CASE INLIST(m.cod, 101003, 101028, 101030) AND INLIST(m.rslt,347,348,349,350,351)
    m.perem = LEFT(sn_pol,17) + '3' && ищем услугу 101952 - в этой категории только одна осталась
    IF SEEK(m.perem, 'dspp')
     m.is_ok = .F.
     DO WHILE LEFT(dspp.sn_pol,17)+'3' = m.perem
      IF m.d_u-dspp.d_u<=60 AND EMPTY(dspp.er) AND INLIST(dspp.rslt,369,370,371,372)
       m.is_ok = .T.
       EXIT 
      ENDIF 
     SKIP IN dspp
     ENDDO 
     IF m.is_ok = .F.
      m.recid = recid
      rval = InsError('S', 'DNA', m.recid, '1',;
    		'Услуга 101952 с результатом 369,370,371,372 не было оказано ранее')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ELSE && не было такой услуги
     m.recid = recid
     rval = InsError('S', 'DNA', m.recid, '1',;
    	'Услуга 101952 с результатом 369,370,371,372 не было оказано ранее')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF  
     
   CASE INLIST(m.cod, 101028, 101030) AND INLIST(m.rslt,332,333,334,335,336)
    m.perem = LEFT(sn_pol,17) + '4' 
    IF SEEK(m.perem, 'dspp')
     m.is_ok = .F.
     DO WHILE LEFT(dspp.sn_pol,17)+'4' = m.perem
      IF m.d_u-dspp.d_u<=60 AND EMPTY(dspp.er) AND INLIST(dspp.rslt,361,362,363,364)
       m.is_ok = .T.
       EXIT 
      ENDIF 
     SKIP IN dspp
     ENDDO 
     IF m.is_ok = .F.
      m.recid = recid
      rval = InsError('S', 'DNA', m.recid, '1',;
    		'Услуга 101933-101945, 101951 с результатом 361,362,363,364 не было оказано ранее')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ELSE && не было такой услуги
     m.recid = recid
     rval = InsError('S', 'DNA', m.recid, '1',;
    	'Услуга 1101933-101945, 101951 с результатом 361,362,363,364 не было оказано ранее')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF  
     
   CASE INLIST(m.cod,1017,1807) AND INLIST(m.rslt,317,318,355,356)
    m.perem = LEFT(sn_pol,17) + '1' 
    IF SEEK(m.perem, 'dspp')
     m.is_ok = .F.
     DO WHILE LEFT(dspp.sn_pol,17)+'1' = m.perem
      IF m.d_u-dspp.d_u<=60 AND EMPTY(dspp.er) AND INLIST(dspp.rslt,353,357,358)
       m.is_ok = .T.
       EXIT 
      ENDIF 
     SKIP IN dspp
     ENDDO 
     IF m.is_ok = .F.
      m.recid = recid
      rval = InsError('S', 'DNA', m.recid, '1',;
    		'Услуга 1925-1935 с результатом 353,357,358 не было оказано ранее')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ELSE && не было такой услуги
     m.recid = recid
     rval = InsError('S', 'DNA', m.recid, '1',;
    	'Услуга 1925-1935 с результатом 353,357,358 не было оказано ранее')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF  
   
  ENDCASE 
 ENDIF 

RETURN 

FUNCTION InsError(WFile, cError, cRecId, cDetail, cComment)
 IF PARAMETERS()<5
  cComment = ''
 ENDIF 
 IF PARAMETERS()<4
  cDetail = ''
 ENDIF 
 IF WFile == 'R'
  IF !SEEK(cRecId, 'rError')
   INSERT INTO rError (f, et, c_err, rid, detail, comment) VALUES ('R', 2, cError, cRecId, cDetail, cComment)
  ELSE 
  ENDIF !SEEK(cRecId, 'rError')
 ENDIF 
 IF WFile == 'S'
  IF !SEEK(cRecId, 'sError')
   INSERT INTO rError (f, et, c_err, rid, detail, comment) VALUES ('S', 2, cError, cRecId, cDetail, cComment)
   RETURN .T.
  ELSE 
   IF cError != sError.c_err
    INSERT INTO rError (f, et, c_err, rid, detail, comment) VALUES ('S', 2, cError, cRecId, cDetail, cComment)
   ENDIF cError != sError.c_err 
  ENDIF !SEEK(cRecId, 'sError')
 ENDIF 
RETURN .F.

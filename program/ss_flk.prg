PROCEDURE ss_flk

   IF M.CVA && COVID с 202004
    m.cod = cod
    m.d_u = d_u
    m.k_u = k_u
    m.ord = ord
    m.lpu_ord = lpu_ord
    m.lpu_ord = IIF(m.lpu_ord>9999 AND FLOOR(m.lpu_ord/10000)=77, m.lpu_ord%10000, m.lpu_ord)
    m.recid = recid
    m.n_u   = n_u
    m.tip   = tip
    IF INLIST(m.cod,70150,70160,70170,170150,170151,170160,170161,170170,170171) AND ;
    	(!SEEK(m.lpuid,'sprnco') OR sprnco.pnv<>1 OR ;
    	m.d_u < IIF(FIELD('DATEBEG_2','sprnco')='DATEBEG_2', sprnco.datebeg_2, sprnco.datebeg) OR ;
    	m.d_u > IIF(FIELD('DATEEND_2','sprnco')='DATEEND_2', sprnco.dateend_2, sprnco.dateend))
     rval = InsError('S', 'CVA', m.recid, '',;
     	STR(m.lpuid)+' не найдено в sprnco или sprnco.pnv<>1 ('+STR(sprnco.pnv,1)+') или d_u ('+DTOC(m.d_u)+') <sprnco.datebeg ('+DTOC(IIF(FIELD('DATEBEG_2','sprnco')='DATEBEG_2', sprnco.datebeg_2, sprnco.datebeg))+')')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF INLIST(m.cod,61400,161400,161401) AND (!SEEK(m.lpuid, 'sprnco') OR sprnco.ncov<>1 OR ;
    	m.d_u < IIF(FIELD('DATEBEG_1','sprnco')='DATEBEG_1', sprnco.datebeg_1, sprnco.datebeg) OR ;
    	m.d_u > IIF(FIELD('DATEEND_1','sprnco')='DATEEND_1', sprnco.dateend_1, sprnco.dateend))
     rval = InsError('S', 'CVA', m.recid, '',;
     	STR(m.lpuid)+' не найдено в sprnco или sprnco.ncov<>1 ('+STR(sprnco.ncov,1)+') или d_u ('+DTOC(m.d_u)+') <sprnco.datebeg ('+DTOC(IIF(FIELD('DATEBEG_1','sprnco')='DATEBEG_1', sprnco.datebeg_1, sprnco.datebeg))+')')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF INLIST(m.cod,61410,161410,161411) AND (!SEEK(m.lpuid, 'sprnco') OR sprnco.dol<>1 OR ;
    	m.d_u < IIF(FIELD('DATEBEG_5','sprnco')='DATEBEG_5', sprnco.datebeg_5, sprnco.datebeg) OR ;
    	m.d_u > IIF(FIELD('DATEEND_5','sprnco')='DATEEND_5', sprnco.dateend_5, sprnco.dateend))
     rval = InsError('S', 'CVA', m.recid, '',;
     	STR(m.lpuid)+' не найдено в sprnco или sprnco.dol<>1 ('+STR(sprnco.dol,1)+') или d_u ('+DTOC(m.d_u)+') <sprnco.datebeg ('+DTOC(IIF(FIELD('DATEBEG_5','sprnco')='DATEBEG_5', sprnco.datebeg_5, sprnco.datebeg))+')')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF INLIST(m.cod,60010,160010)
     IF !SEEK(m.lpuid, 'sprnco') OR sprnco.trs<>1 OR ;
    	m.d_u < IIF(FIELD('DATEBEG_6','sprnco')='DATEBEG_6', sprnco.datebeg_6, sprnco.datebeg) OR ;
    	m.d_u > IIF(FIELD('DATEEND_6','sprnco')='DATEEND_6', sprnco.dateend_6, sprnco.dateend)
      rval = InsError('S', 'CVA', m.recid, '',;
     	STR(m.lpuid)+' не найдено в sprnco или sprnco.trs<>1 ('+STR(sprnco.trs,1)+') или d_u ('+DTOC(m.d_u)+') <sprnco.datebeg ('+DTOC(IIF(FIELD('DATEBEG_6','sprnco')='DATEBEG_6', sprnco.datebeg_6, sprnco.datebeg))+')')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ELSE 
      IF (SEEK(m.lpu_ord, 'sprnco') AND (sprnco.ncov=1 OR sprnco.pnv=1)) OR INLIST(m.lpu_ord,4708,502009)
      ELSE 
       rval = InsError('S', 'CVA', m.recid, '',;
     	'lpu_ord ('+STR(m.lpu_ord)+') не найдено в sprnco или sprnco.trs<>1 ('+STR(sprnco.trs,1)+') или d_u ('+DTOC(m.d_u)+') <sprnco.datebeg ('+DTOC(IIF(FIELD('DATEBEG_6','sprnco')='DATEBEG_6', sprnco.datebeg_6, sprnco.datebeg))+')')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
    
    IF INLIST(m.cod,61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171)
     DO CASE 
      CASE m.ord=2
       IF !INLIST(m.lpuid,4511,2293,4586)
        IF (INLIST(m.lpu_ord,4708) AND ;
    	 (ISDIGIT(SUBSTR(m.n_u,1,1)) AND ISDIGIT(SUBSTR(m.n_u,2,1)) AND ISDIGIT(SUBSTR(m.n_u,3,1)) AND ISDIGIT(SUBSTR(m.n_u,4,1)) AND ;
          ISDIGIT(SUBSTR(m.n_u,5,1)) AND ISDIGIT(SUBSTR(m.n_u,6,1)) AND ISDIGIT(SUBSTR(m.n_u,7,1)) AND ISDIGIT(SUBSTR(m.n_u,8,1)) AND ;
          ISDIGIT(SUBSTR(m.n_u,9,1)))) OR (INLIST(m.lpu_ord,502009) AND !EMPTY(m.n_u))
        ELSE 
         rval = InsError('S', 'CVA', m.recid, '',;
     	 	'МЭС '+STR(m.cod,6)+' c ord=2 и lpu_ord<>4708 ('+STR(m.lpu_ord,4)+') или некорректным n_u ('+m.n_u)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       ELSE 
        IF !SEEK(m.lpu_ord, 'sprlpu') AND !SEEK(m.lpu_ord, 'f003')
         rval = InsError('S', 'CVA', m.recid, '',;
     	 	'МЭС '+STR(m.cod,6)+' c ord=2 и lpu_ord ('+STR(m.lpu_ord,4)+') не из справочника sprlpu/f003')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       ENDIF 
      
      CASE INLIST(m.lpuid,4511,2293,4586) AND m.ord=3
      
      CASE INLIST(m.cod,61400,161400,161401,61430,161430,161431) AND m.ord=1

      OTHERWISE 
       rval = InsError('S', 'CVA', m.recid, '',;
     	'МЭС '+STR(m.cod,6)+' c ord<>2 ('+STR(m.ord,1)+') или lpu_ord<>4708 ('+STR(m.lpu_ord,4)+')')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDCASE 

    ENDIF 

    IF INLIST(m.cod,61420,161420,161421) AND !INLIST(m.ord,1,5,6)
       rval = InsError('S', 'CVA', m.recid, '',;
     	'МЭС (предварительная диагностика)'+STR(m.cod,6)+' в сочетании c ord не равным 1,5')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

   ENDIF 
   
   IF M.R4A==.T.
    m.prvs = prvs

    m.c_i    = c_i 
    m.sn_pol = sn_pol
    m.d_u    = d_u

    IF OCCURS('#', m.c_i)=3
     m.dr = CTOD(SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),5,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),1,4))
    ELSE 
     m.dr = IIF(SEEK(m.sn_pol, 'people'), people.dr, {})
    ENDIF 

    m.adj  = CTOD(STRTRAN(DTOC(m.dr), STR(YEAR(m.dr),4), STR(YEAR(people.d_beg),4)))-people.d_beg
    *m.vozr   = (YEAR(m.d_u) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)
    m.vozr   = (YEAR(people.d_beg) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)
    m.i_otd  = SUBSTR(otd,4,3)
    m.k_u    = k_u
    
    m.tip     = tip
    m.d_type = d_type
    m.ds     = ds
    
    DO CASE 
    *CASE INLIST(m.d_type,'1','2','5','6','e') OR m.tip='5'
    CASE INLIST(m.d_type,'1','5','6','e') OR m.tip='5'
    CASE m.k_u<=1 AND (SEEK(m.ds, 'nocodr', 'ds1') OR SEEK(m.ds, 'nocodr', 'ds2') OR SEEK(m.ds, 'nocodr', 'ds3'))
    
    OTHERWISE 
    
    IF INLIST(m.prvs,81,41,73,82,42,11,83,149,22,174) AND m.vozr>=18
     m.recid = recid
     rval = InsError('S', 'R4A', m.recid, '',;
     	'Пациент старше 18 лет при специальности исполнителя={17,18,19,20,21,68,86}, с 02.2020')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF INLIST(m.prvs,15) AND m.vozr>1
     m.recid = recid
     rval = InsError('S', 'R4A', m.recid, '',;
     	'Пациент страше 1 года при специальности исполнителя=15, с 02.2020')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    *IF INLIST(m.prvs,66,116) AND m.vozr<65
    IF INLIST(m.prvs,66,116) AND m.vozr<60
     m.recid = recid
     rval = InsError('S', 'R4A', m.recid, '',;
     	'Пациент моложе 60 лет при специальности исполнителя 66/116, с 12.2020')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF INLIST(m.prvs,118,260,17,145,31) AND m.vozr<18
     m.recid = recid
     rval = InsError('S', 'R4A', m.recid, '',;
     	'Пациент моложе 18 лет при специальности исполнителя={118,260,17,145,31}, с 02.2020')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    IF INLIST(m.prvs,27) AND m.vozr<15
     m.recid = recid
     rval = InsError('S', 'R4A', m.recid, '',;
     	'Пациент моложе 15 лет при специальности исполнителя 27, с 02.2020')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    ENDCASE 
   ENDIF 

   IF M.NWA = .T. AND INT(VAL(m.gcPeriod))>201910 && включена со счетов за ноябрь
   *IF !INLIST(m.mcod,'0343020','0343999','0343067','0343001')
    m.i_otd = SUBSTR(otd,4,3)
    m.s_otd = SUBSTR(otd,2,2)
    IF INLIST(m.i_otd,'135','139','008','010','140','141','142','027','144','145','036','174','176','149','148','050') OR ;
       INLIST(m.i_otd,'049','150','047','051','152','153','154','156','155','157','159','173','175','161','062','164','071') OR ;
       INLIST(m.i_otd,'073','165','074','080','170','091','172','093','094','095','104','178','107','110','117','180') OR ;
       INLIST(m.i_otd,'120','121','181','182') OR (m.i_otd='072' AND !INLIST(m.s_otd,'94','76'))
     m.recid = recid
     rval = InsError('S', 'NWA', m.recid, '',;
     	'Регистрация медуслуг в отделении с профилем (4-6 позиции), не включенным в ОМС')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    m.prvs = prvs
    IF INLIST(m.prvs,110,288,3200,256,159,211,160,161,162,213,169,212,287,163,1,2,279,236,237,238,239,240) OR ;
       INLIST(m.prvs,241,242,243,244,245,246,7,289,269,290,235,232,152,18,267,268,205,23,97,98,164,165,99) OR ;
       INLIST(m.prvs,153,26,194,234,100,184,186,229,185,4,218,285,29,214,154,210)
     m.recid = recid
     rval = InsError('S', 'NWA', m.recid, '',;
     	'Недопустимый профиль prvs (специальность исполнителя медицинской помощи)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
        
   *ENDIF 
   ENDIF 
   
   IF M.PGA = .T. AND INT(VAL(m.gcPeriod))>201910 && включена со счетов за ноябрь
    IF m.IsHor OR m.IsPilot && IIF(m.qcod='S7', m.IsPilot, 3=2) && Добавил 03.12.2019 проба!!
    
     ** INLIST(m.lpuid,1872,2290)
    
     m.lpu_ord = lpu_ord
     m.date_ord = date_ord
     m.ord     = ord
     m.cod     = cod
     m.Is02    = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
     m.IsTpnR  = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)
     m.p_cel   = p_cel
     m.recid   = recid
     m.otd     = otd
     m.usl_ok  = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), profot.usl_ok, '0')
     m.sn_pol  = sn_pol
     m.ds      = ds
     m.d_type  = d_type
     *m.c_i     = ALLTRIM(c_i)
     m.c_i     = c_i
     m.Mp      = Mp
     m.ds      = ds
     
     m.IsUslGosp = .F.
     IF USED('hosp')
      m.IsUslGosp = IIF(IsUsl(m.cod) AND SEEK(m.c_i, 'hosp'), .T., .F.)
     ENDIF 

     IF INLIST(m.lpuid,1872,2290)  && Убрал 31.08.2020 - не уверен!
      IF INLIST(INT(m.cod/1000),25,125,26,126,27,127,28,128,29,129,30,130) AND ;
      	(!INLIST(SUBSTR(m.otd,2,2),'00','70','73') AND m.usl_ok='3')
       IF !(INLIST(INT(m.cod/1000),29,129) AND SUBSTR(m.otd,2,2)='85')
        IF SEEK(m.sn_pol, 'people') AND !EMPTY(people.prmcod)
         *IF !((m.lpu_ord=2290 AND m.ord=4) OR (m.ord=8 AND m.lpu_ord=8888 AND YEAR(m.date_ord)<>0))
          IF !(m.lpuid=2290 AND INLIST(m.lpu_ord,2290,772290))
           rval = InsError('S', 'PGA', m.recid, '',;
      		'Оказание услуг из разелов 25,125,26,126,27,127,28,128,29,129,30,130 в отделении не 00 в МО 1872 или 2290')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ENDIF 
         *ENDIF 
        ENDIF 
       ENDIF 
      ENDIF 
     ENDIF 

     IF !IsDental(m.cod, m.lpuid, m.mcod, m.ds)

     IF IsUsl(m.cod) AND (SEEK(m.sn_pol, 'people') AND !EMPTY(people.prmcod) AND people.prmcod!=m.mcod)

     DO CASE 
     *CASE m.IsTpnR = .T. OR m.d_type='s' OR (INLIST(SUBSTR(m.otd,2,2),'08','91'))
     CASE m.IsTpnR = .T. 
     *CASE INLIST(SUBSTR(m.otd,2,2),'08','91')
     CASE INLIST(SUBSTR(m.otd,2,2),'08','85','91') && с 202002 по предложению ВТБ_МС
     CASE IsMes(m.cod) OR IsVMP(m.cod)
     CASE INLIST(SUBSTR(m.otd,2,2),'01') AND IsStac(m.mcod)
     CASE INLIST(SUBSTR(m.otd,2,2),'70','73','90','93') AND IsStac(m.mcod)
     CASE m.ord=7 AND m.lpu_ord=7665
     CASE m.IsUslGosp=.T.
     *CASE INLIST(INT(m.cod/1000),49,149) AND m.mcod!=people.prmcod AND people.tip_p=3 
     *CASE INLIST(INT(m.cod/1000),29,129) AND m.mcod!=people.prmcod AND people.tip_p=3 
     CASE BETWEEN(m.cod,1815,1819) OR  INLIST(m.cod,101523,101533,101543,101553) && патронаж
     CASE OCCURS('#', m.c_i)=3 && добавил 20.01.2020
     *CASE !INLIST(m.Mp,'p','s') && добавил 20.01.2020
     CASE INLIST(m.cod,56029,156003)
     CASE INLIST(m.cod,28211,128211) AND IIF(USED('sprnco') AND SEEK(m.lpuid,'sprnco') AND sprnco.diag=1 AND BETWEEN(m.d_u,m.d_b,m.d_e), .T., .F.)
     CASE INLIST(m.cod,28165,128165) AND IIF(USED('sprnco') AND SEEK(m.lpuid, 'sprnco') AND sprnco.ig=1 AND BETWEEN(m.d_u,m.d_b2,m.d_e2), .T., .F.)
     CASE INLIST(m.cod,37043,137043) AND (INLIST(m.ds, 'B34.2','J02','J04','J06','J20','U07.1','U07.2') OR ;
  	 	BETWEEN(LEFT(m.ds,3),'J09','J18'))
  	 CASE INLIST(m.cod,37043,37048,137043,37044,37049,137044,137049) AND ;
  		(m.ds='C' OR BETWEEN(LEFT(m.ds,3),'D00','D09'))
  	 CASE INLIST(m.cod,60010,160010) AND SEEK(m.lpuid,'sprnco') AND m.d_u>={23.03.2020}
  	 *CASE INLIST(m.lpuid,1872,2290) AND (m.lpu_ord=2290 AND m.ord=4)
  	 
     OTHERWISE 

      IF !INLIST(SUBSTR(m.otd,2,2), '01','85','08','81',IIF(m.qcod='S7', '90', 'AB'),'92')
       IF EMPTY(m.lpu_ord) OR m.ord=0 OR m.lpu_ord=m.lpuid && добавлено 20.01.2020 - направление сами себе!
        IF m.Is02
         IF !INLIST(SUBSTR(m.otd,2,2),'91','27','92')
          IF m.p_cel<>'1.1'
           rval = InsError('S', 'PGA', m.recid, '',;
      	  	'Оказание скоропомощной услуги в "чужом" МО при p_cel="'+m.p_cel+'" (не равном 1.1)')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ENDIF 
         ENDIF 
        ELSE 
         DO CASE 
          CASE EMPTY(m.lpu_ord) OR m.ord=0
           rval = InsError('S', 'PGA', m.recid, '',;
      		'Оказание нескоропомощной услуги в "чужом" МО без направления')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          CASE m.lpu_ord=m.lpuid
           rval = InsError('S', 'PGA', m.recid, '',;
      		'Оказание нескоропомощной услуги в "чужом" МО по направлению, выданному "самому себе"')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          OTHERWISE 
         ENDCASE 
        ENDIF 
       ENDIF 
      ENDIF 

     ENDCASE 

     ENDIF 
    ENDIF && IF !IsDental(m.cod, m.lpuid, m.mcod, m.ds) 
    ENDIF && IF m.IsHor && Добавил 03.12.2019 проба!!
   ENDIF && IF M.PGA = .T. AND INT(VAL(m.gcPeriod))>201910 && включена со счетов за ноябрь

   IF M.PHA = .T. && AND USED('ppr')
    IF m.IsHor OR m.IsPilot
     m.cod     = cod
     m.recid   = recid
     m.sn_pol  = sn_pol
     m.ds      = ds
     m.c_i     = c_i
     m.p_cel   = p_cel
     

     m.prmcod  = IIF(SEEK(m.sn_pol, 'people'), people.prmcod, '')
     m.otd     = otd
     m.profil  = profil
     m.lpu_ord = lpu_ord
     m.ord     = ord

	 m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r', .T., .F.)
	 m.IsTpnR    = IIF(INLIST(m.cod,28211,128211) AND IIF(USED('sprnco') AND SEEK(m.lpuid,'sprnco') AND sprnco.diag=1 AND BETWEEN(m.d_u, m.d_b, m.d_e),.T.,.F.), .T., m.IsTpnR)
     m.IsTpnR    = IIF(INLIST(m.cod,28165,128165) AND IIF(USED('sprnco') AND SEEK(m.lpuid,'sprnco') AND sprnco.ig=1 AND BETWEEN(m.d_u, m.d_b2, m.d_e2),.T.,.F.), .T., m.IsTpnR) && со счетов за май.
	 m.IsTpnR    = IIF(INLIST(m.cod,37043,137043) AND ;
	  	(INLIST(m.ds, 'B34.2','J02','J04','J06','J20','U07.1','U07.2') OR ;
	  	 BETWEEN(LEFT(m.ds,3),'J09','J18')), .T., m.IsTpnR)
	 m.IsTpnR    = IIF(INLIST(m.cod,37043,37048,137043,37044,37049,137044,137049) AND ;
	  	(m.ds='C' OR BETWEEN(LEFT(m.ds,3),'D00','D09')), .T., m.IsTpnR)
	 m.IsTpnR    = IIF(INLIST(m.cod,60010,160010) AND m.IsSprNCO AND m.d_u>=m.d_b, .T., m.IsTpnR)

     m.Is02    = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
     m.d_type  = d_type
    
     m.facotd  = SUBSTR(m.otd,2,2)
     m.profil  = profil
     m.lpu_ord = lpu_ord

     m.Mp      = ''
     m.dop_r   = 0
     m.vz      = 0

     m.IsUslGosp = .F.
     IF USED('hosp')
      m.IsUslGosp = IIF(IsUsl(m.cod) AND SEEK(m.c_i, 'hosp'), .T., .F.)
     ENDIF 

     IF IsDental(m.cod, m.lpuid, m.mcod, m.ds)
     ELSE 
	  DO CASE 
       CASE !EMPTY(m.prmcod) AND m.mcod  = m.prmcod && свои пациенты
       
       CASE !EMPTY(m.prmcod) AND m.mcod != m.prmcod && чужие пациенты

	   CASE EMPTY(m.prmcod) && неприкрепленные
	    m.Is02 = IIF(SEEK(m.cod, 'pervpr') AND m.p_cel='1.1', .T., .F.)
	    m.Typ = '0'
	    DO CASE 
	     CASE m.IsTpnR = .T.
	      IF m.IsPilot OR m.IsHor
	       m.dop_r = 1
	       m.Mp    = '4'
	      ENDIF 
	     CASE INLIST(m.facotd,'08','85')
	      IF m.IsPilot OR m.IsHor
	       m.dop_r = 3
	       m.Mp = '4'
	      ENDIF 
	     CASE INLIST(m.cod,56029,156003)
	      IF m.IsPilot OR m.IsHor
	       m.dop_r = 3
	       m.Mp = '4'
	      ENDIF 
	     CASE IsMes(m.cod) OR IsVMP(m.cod)
	      m.Mp = 'm'
	     CASE INLIST(m.facotd,'01') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor)
	      m.dop_r = 5
	      m.Mp = '4'
	     CASE INLIST(m.facotd,'70','73','93') AND IsStac(m.mcod) AND (m.IsPilot OR m.IsHor)
	      m.dop_r = 4
	      m.Mp = '4'
	     CASE m.ord=7 AND m.lpu_ord=7665 AND (m.IsPilot OR m.IsHor)
	      m.dop_r = 6
	      m.Mp = '4'
	     CASE m.IsUslGosp
	      IF m.IsPilot OR m.IsHor
	       m.dop_r = 2
	       m.Mp = '4'
	      ENDIF 
	     CASE INLIST(m.cod, 37043,37044,37048,37049,137043,137044,137049) AND LEFT(m.ds,1)='C'
	      IF m.IsPilot OR m.IsHor
	       m.Mp    = '4'
	       m.dop_r = 11
	      ENDIF 
	     OTHERWISE 
	       m.Mp = 'p'
	       * только эти услуги проверяем на PH
		   DO CASE 
		    CASE OCCURS('#', m.c_i)=3 && добавил 20.01.2020
		    CASE SUBSTR(m.otd,2,2)<>'00'
		    OTHERWISE 
		     IF !SEEK(m.cod,'pervpr')
		      *IF !INLIST(m.lpuid,1872,2290) 
		       rval = InsError('S', 'PHA', m.recid, '',;
		    		'Оказание терапевтических услуг неприкрепленным пациентам (с 04.2020)')
		       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
		      *ENDIF 
		    ENDIF 
		   ENDCASE 
	       * только эти услуги проверяем на PH
	    ENDCASE 
	  ENDCASE 
     ENDIF 
    ENDIF && IF m.IsHor && Добавил 03.12.2019 проба!!
   ENDIF && IF M.PHA = .T.

   IF M.UMA = .T.
    IF USED('Gosp')
     m.otd    = otd
     m.sn_pol = sn_pol
     m.cod    = cod
     m.d_u    = d_u
     m.ds     = ds
     m.d_type = d_type
     m.recid  = recid
     m.c_i    = c_i
     
     SET ORDER TO sn_pol IN Gosp
     IF SEEK(m.sn_pol, 'Gosp', 'sn_pol')
      DO WHILE Gosp.sn_pol = m.sn_pol
       DO CASE 
        CASE IsMes(m.cod) OR IsVmp(m.cod)
        CASE IsMes(Gosp.cod) AND INLIST(m.cod,1781,101781) AND m.mcod=Gosp.mcod
        CASE IsMes(Gosp.cod) AND INLIST(INT(m.cod/1000), 29,129)

        CASE IsMes(Gosp.cod) AND INLIST(INT(m.cod/1000),59,159) AND m.c_i=Gosp.c_i AND m.d_u=Gosp.d_u

        CASE m.lpuid = 1874 AND IsMes(Gosp.cod) AND INLIST(INT(m.cod/1000),138)
        CASE IsMes(Gosp.cod) AND m.d_type='s'
        CASE IsMes(Gosp.cod) AND INLIST(INT(m.cod/1000), 49,149) AND m.d_type='2'
        CASE IsMes(Gosp.cod) AND INLIST(m.cod,97010,197010) AND m.d_type='2'
        CASE IsMes(Gosp.cod) AND m.d_type='w'
        CASE IsMes(Gosp.cod) AND INLIST(m.cod,36022,136022,36023,136023,36024,136024) AND m.d_type='2'
        CASE IsMes(Gosp.cod) AND INLIST(INT(m.cod/1000),84)
        CASE IsMes(Gosp.cod) AND INLIST(INT(m.cod/1000), 297) AND m.lpuid=1872  && Лучевая терапия в Морозовской
        CASE INLIST(Gosp.cod,200518, 200519, 200520, 200522, 200523, 200524) AND ;
        	INLIST(m.cod,49011, 49012, 49013, 49024, 49030, 49035) AND m.d_type='2' && гемодиализ во время ВМП
        CASE OCCURS('#', m.c_i)=3
        
        OTHERWISE 
         IF Gosp.k_u>1
          IF BETWEEN(m.d_u, Gosp.d_u-Gosp.k_u+1, Gosp.d_u-1)
           m.recid = recid
           rval = InsError('S', 'UMA', m.recid, '',;
      	 	'Оказание услуг в период лечения по МЭС/СКП')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ENDIF 
         ELSE 
         ENDIF 
       ENDCASE 
       
       SKIP IN Gosp
      ENDDO 
     ENDIF 

    ENDIF 

   ENDIF 

   
   IF M.CVA = .T. && AND .F.
    IF USED('Gosp')
     m.otd    = otd
     m.sn_pol = sn_pol
     m.cod    = cod
     m.d_u    = d_u
     m.ds     = ds
     m.d_type = d_type
     m.recid  = recid
     m.c_i    = c_i
     m.k_u    = k_u
     m.recid  = recid
     
     IF !IsUsl(m.cod) && если не услуга
     
      SET ORDER TO sn_pol IN Gosp
      IF SEEK(m.sn_pol, 'Gosp')
       DO WHILE Gosp.sn_pol = m.sn_pol
        IF m.recid<>Gosp.recid AND ;
         (INLIST(Gosp.cod, 61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171) AND  ;
       	  INLIST(m.cod,61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171))
         IF m.c_i=Gosp.c_i && BETWEEN(m.d_u, Gosp.d_u-Gosp.k_u, Gosp.d_u)
          m.recid = recid
          rval = InsError('S', 'CVA', m.recid, '',;
      	   	'Два и более короновирусных МЭСа за госпитализацию')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 
        ENDIF        

        IF m.cod<>Gosp.cod  AND ;
       	 ((INLIST(Gosp.cod, 61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171,61420,161420,16142) AND  ;
       	  !INLIST(m.cod,61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171,61420,161420,16142)) OR ;
       	 (INLIST(Gosp.cod, 61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171,61420,161420,16142) AND  ;
       	  !INLIST(m.cod,61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171,61420,161420,16142)))
       	 
         IF m.c_i=Gosp.c_i && BETWEEN(m.d_u, Gosp.d_u-Gosp.k_u, Gosp.d_u)
          m.recid = recid
          rval = InsError('S', 'CVA', m.recid, '',;
      	   	'Короновирусный и некороновирусный МЭС за одну госпитализацию (1)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 
        ENDIF        

        *IF m.cod<>Gosp.cod AND (!INLIST(Gosp.cod, 61420,161420,161421) AND INLIST(m.cod, 61420,161420,161421))
        IF m.cod<>Gosp.cod AND ;
        	(!INLIST(Gosp.cod, 61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171) ;
        	AND INLIST(m.cod, 61420,161420,161421))
         IF m.c_i=Gosp.c_i AND m.d_u<Gosp.d_u
          m.recid = recid
          rval = InsError('S', 'CVA', m.recid, '',;
      	   	'Профильный МЭС после обсервационного (61420,161420,161421)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 
        ENDIF        

        SKIP IN Gosp
       ENDDO 
      ENDIF 

     ELSE && если услуга
     
      SET ORDER TO sn_pol IN Gosp
      IF SEEK(m.sn_pol, 'Gosp')
       DO WHILE Gosp.sn_pol = m.sn_pol
        IF INLIST(Gosp.cod, 61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171) && если короновирусный МЭС
         IF (INLIST(INT(m.cod/1000),49,149) AND !INLIST(m.cod,49007,49020)) AND m.c_i=Gosp.c_i
         ELSE 
         
         IF m.c_i=Gosp.c_i
          *IF BETWEEN(m.d_u, Gosp.d_u-Gosp.k_u, Gosp.d_u)
          IF !INLIST(INT(m.cod/1000),29,129,59,159) OR (INLIST(INT(m.cod/1000),29,129,59,159) AND m.d_u=Gosp.d_u)
           m.recid = recid
           rval = InsError('S', 'CVA', m.recid, '',;
      	   	'Услуга во время леченя по короновирусному МЭСу (совпадающие c_i) '+;
      	   		DTOC(Gosp.d_u-Gosp.k_u+IIF(Gosp.k_u>1,1,0)) + DTOC(Gosp.d_u-1))
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
           ENDIF 
          *ENDIF 
         ELSE 
          IF Gosp.k_u>1 AND BETWEEN(m.d_u, Gosp.d_u-Gosp.k_u+IIF(Gosp.k_u>1,1,0), Gosp.d_u-1)
           m.recid = recid
           rval = InsError('S', 'CVA', m.recid, '',;
      	   	'Услуга во время леченя по короновирусному МЭСу (разное c_i) '+;
      	   		DTOC(Gosp.d_u-Gosp.k_u+IIF(Gosp.k_u>1,1,0)) + DTOC(Gosp.d_u-1))
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ENDIF 
         ENDIF 

         
         ENDIF 

        ENDIF        
        SKIP IN Gosp
       ENDDO 
      ENDIF 

     ENDIF 
    ENDIF 
   ENDIF 

   IF M.SKA = .T.
    IF USED('curskp')
     m.otd    = otd
     m.sn_pol = sn_pol
     m.cod    = cod
     m.d_u    = d_u
     m.ds     = ds
     m.l_ds = LEN(ALLTRIM(m.ds)) && Считаем длину диагноза для последующей проверки
     m.recid   = recid
     IF SEEK(m.sn_pol, 'curskp')
      IF INLIST(SUBSTR(m.otd,2,2),'70','73') AND (!BETWEEN(m.cod,1741,1780) AND !INLIST(INT(m.cod/1000),84,184))
       *m.recid = recid
       rval = InsError('S', 'SKA', m.recid, '',;
      	'Оказание услуг вне диапазона 1741-1780 в приемном отделении (70,73) госпитализированному по СКП')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      *m.d_u, Gosp.d_u-Gosp.k_u+1, Gosp.d_u-1
      *IF m.cod != curskp.cod AND BETWEEN(m.d_u, curskp.d_u-(curskp.k_u+1), curskp.d_u)
      IF (m.cod != curskp.cod AND BETWEEN(m.d_u, curskp.d_u-curskp.k_u+1, curskp.d_u-1)) OR ;
      	(m.d_u=curskp.d_u AND INLIST(INT(m.cod/1000),29,129))
       *m.recid = recid
       rval = InsError('S', 'SKA', m.recid, '',;
      	'Оказание услуг в период госпитализации по СКП')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      
      IF IsUsl(m.cod) AND m.d_u > curskp.d_u AND (SUBSTR(m.otd,4,3)=SUBSTR(curskp.otd,4,3) OR ;
      	INLIST(m.cod,3013, 3014, 3015, 3018, 3030, 3031))
       *m.recid = recid
       rval = InsError('S', 'SKA', m.recid, '',;
      	'Оказание услуг полсе госпитализации по СКП')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      
      m.ln_month = m.tmonth
      m.ln_year  = m.tyear
      FOR m.i_i=1 TO 3
       m.ln_month = IIF(m.ln_month>1, m.ln_month-1, 12)
       m.ln_year  = IIF(m.ln_month=12, m.ln_year-1, m.ln_year)
       m.lcperiod = STR(m.ln_year,4)+PADL(m.ln_month,2,'0')
       IF fso.FolderExists(m.pBase+'\'+m.lcPeriod+'\'+m.mcod)
        IF fso.FileExists(m.pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon.dbf')
         IF OpenFile(m.pBase+'\'+m.lcPeriod+'\'+m.mcod+'\talon', 't_t', 'shar', 'sn_pol')>0
          IF USED('t_t')
           USE IN t_t
          ENDIF 
          *SELECT talon 
          SELECT c_talon 
          LOOP 
         ELSE 
          && Файл открыт
          *SELECT talon 
          SELECT c_talon 
          IF !SEEK(m.sn_pol, 't_t')
           USE IN t_t
           LOOP 
          ELSE 
           && Ищем госпитализации
           SELECT t_t
           DO WHILE sn_pol = m.sn_pol
            m.o_cod = cod
            m.o_d_u = d_u
            m.o_otd = otd
            m.o_ds  = ds
            IF IsMes(m.o_cod) OR IsVmp(m.o_cod)
             *IF SUBSTR(m.o_otd,2,2)<>'09' AND LEFT(m.o_ds,m.l_ds)=LEFT(m.ds,m.l_ds) AND m.d_u - m.o_d_u<90
             IF SUBSTR(m.o_otd,2,2)<>'09' AND LEFT(m.o_ds,m.l_ds)=LEFT(m.ds,m.l_ds) AND m.d_u - m.o_d_u<30
              *m.recid = recid
              rval = InsError('S', 'SKA', m.recid, '',;
      	      	'Пациент ранее лежал в отделении с тем же диагнозом: cod='+PADL(m.o_cod,6,'0')+', d_u='+DTOC(m.o_d_u)+;
      	      		', otd='+m.o_otd)
              m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
              EXIT 
             ENDIF 
            ENDIF 
            SKIP 
           ENDDO 
           USE IN t_t 
           *SELECT talon 
           SELECT c_talon 
           LOOP 
           && Ищем госпитализации
          ENDIF 
          && Файл открыт
         ENDIF 
        ENDIF 
       ENDIF 
      ENDFOR 

     ENDIF 
    ENDIF 
   ENDIF 

   IF M.H6A == .T. AND USED('hosp_p')
    m.c_i   = c_i
    m.cod   = cod 
    m.recid = recid
    m.tip   = Tip
    *IF !INLIST(INT(m.cod/1000),29,59,129,159)
    IF !IsUsl(m.cod)
     IF SEEK(m.c_i, 'hosp_p')
      rval = InsError('S', 'H6A', m.recid, '',;
      	'Истории болезни использовалась в предыдущем периоде: '+DTOC(hosp_p.d_vip))
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.WEA 
    m.cod = cod 
    m.d_u = d_u
    m.IsStomatUsl = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
    m.recid = recid
    IF m.IsStomatUsl AND m.d_u>{17.03.2020}
      rval = InsError('S', 'WEA', m.recid, '',;
      	'Услуга стоматологии, оказанная ГБУЗ "ЧЛГ для ВВ ДЗМ" позднее с 18.03.2020 письмо МГФОМС 07-01-05/7023 от 27 марта 2020')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
   ENDIF 

   IF M.H6A == .T. AND USED('curs_h6') && Алгоритм H6, дубли по c_i
    m.cod   = cod 
    m.recid = recid
    m.tip = tip
    IF (IsGsp(m.cod) OR IsDst(m.cod))
     IF SEEK(m.recid, 'curs_h6')
      rval = InsError('S', 'H6A', m.recid, '',;
      	'Повтор истории болезни')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.H6A == .T. AND USED('curs_h6p') && Алгоритм H6, дубли по c_i
    m.cod   = cod 
    m.recid = recid
    m.tip = tip
    *IF IsGsp(m.cod) OR IsDst(m.cod)
     IF SEEK(m.recid, 'curs_h6p')
      rval = InsError('S', 'H6A', m.recid, '',;
      	'Повтор номера карты')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    *ENDIF 
   ENDIF 

   IF M.H6A == .T. && Алгоритм H6
    * Ревизия от 01.11.2019
    * ПФО август 2019_1_8652455_v1
    m.polis=''
    DO CASE 
     CASE IsKms(sn_pol)
      m.polis = SUBSTR(sn_pol,8)
     CASE IsVs(sn_pol)
      m.polis = SUBSTR(sn_pol,7)
     OTHERWISE 
      m.polis = sn_pol
    ENDCASE 
    *m.polis = ALLTRIM(sn_pol)

    IF EMPTY(c_i) OR '0'=c_i
     m.recid = recid
     rval =InsError('S', 'H6A', m.recid, '', ;
     	'Пустое поле c_i (номер карты)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF

    IF (INLIST(cod,101927,101928,101951) OR BETWEEN(cod,101933,101945))
     *IF SUBSTR(c_i,1,6)!='ПРОФД_' OR (ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(m.polis) AND ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(sn_pol))
     IF SUBSTR(c_i,1,6)!='ПРОФД_' OR ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(sn_pol) && С 202003
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid, '', ;
     	'Неверная кодировка карты для профосмотров детей 101927,101928,101951,101933-101945 (д.б. ПРОФД_полис)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF INLIST(cod,101946,101947,101948)
     *IF SUBSTR(c_i,1,6)!='ПРЕДД_' OR (ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(m.polis) AND ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(sn_pol))
     IF SUBSTR(c_i,1,6)!='ПРЕДД_' OR ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(sn_pol) && С 202003
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid, '', ;
     	'Неверная кодировка карты для предварительных медосмотров 101946,101947,101948 (д.б.ПРЕДД_полис)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF INLIST(cod,101949,101950)
     *IF SUBSTR(c_i,1,4)!='ПОД_' OR (ALLTRIM(SUBSTR(c_i,5))!=ALLTRIM(m.polis) AND ALLTRIM(SUBSTR(c_i,5))!=ALLTRIM(sn_pol))
     IF SUBSTR(c_i,1,4)!='ПОД_' OR ALLTRIM(SUBSTR(c_i,5))!=ALLTRIM(sn_pol) && С 202003
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid, '', ;
     	'Неверная кодировка карты для периодических медосмотров 101949,101950 (д.б. ПОД_полис)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF (BETWEEN(cod,1900,1905) OR BETWEEN(cod,101929,101932))
     *IF !INLIST(SUBSTR(c_i,1,3),'ДД_','ДС_','ДУ_') OR ;
     	(ALLTRIM(SUBSTR(c_i,4))!=ALLTRIM(m.polis) AND ALLTRIM(SUBSTR(c_i,4))!=ALLTRIM(sn_pol))
     IF !INLIST(SUBSTR(c_i,1,3),'ДД_','ДС_','ДУ_') OR ;
     	ALLTRIM(SUBSTR(c_i,4))!=ALLTRIM(sn_pol)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid, '', ;
     	'Неверная кодировка карты для дипансеризации 1900-1905, 101929-101932 (д.б. ДД_/ДС_/ДУ_полис)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    *IF BETWEEN(cod,1906,1909)
    *IF BETWEEN(cod,1921,1924)
    IF BETWEEN(cod,1949,1954) OR BETWEEN(cod,1968,1973)
     *IF SUBSTR(c_i,1,6)!='ПРОФВ_' OR ;
     	(ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(m.polis) AND ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(sn_pol))
     IF SUBSTR(c_i,1,6)!='ПРОФВ_' OR ;
     	ALLTRIM(SUBSTR(c_i,7))!=ALLTRIM(sn_pol)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid, '', ;
     	'Неверная кодировка карты для профилактических медосмотров 1949-1954/1968-1973 (д.б. ПРОФВ_полис)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    *IF BETWEEN(cod,1925,1935)
    IF BETWEEN(cod,1936,1948) OR BETWEEN(cod,1955,1967) && с января 2020
     *IF !INLIST(SUBSTR(c_i,1,3),'ДД_') OR ;
     	(ALLTRIM(SUBSTR(c_i,4))!=ALLTRIM(m.polis) AND ALLTRIM(SUBSTR(c_i,4))!=ALLTRIM(sn_pol))
     IF !INLIST(SUBSTR(c_i,1,3),'ДД_') OR ;
     	ALLTRIM(SUBSTR(c_i,4))!=ALLTRIM(sn_pol)
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid, '', ;
     	'Неверная кодировка карты для дипансеризации 1936,1948/1955,1967 (д.б. ДД_)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF ord=7 AND lpu_ord=7665
     IF SUBSTR(c_i,1,3)<>'УМО'
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid, '', ;
     	'Неверная кодировка карты для УМО (д.б. УМО+номер ВКК диспансерного наблюдения спортсмена)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
    IF ((people.d_type='9' AND m.cod>100000) AND OCCURS('#', ALLTRIM(c_i))<>3) OR ;
    	(OCCURS('#', ALLTRIM(c_i))=3 AND people.d_type<>'9')
     m.d_type = d_type
     m.tpn = IIF(SEEK(m.cod, 'tarif'), tarif.tpn, '?')
     *IF !(INLIST(m.tpn,'q','r') AND INLIST(m.d_type,'2','e'))
     IF !(INLIST(m.tpn,'q','r') AND INLIST(m.d_type,'e'))
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid, '', ;
     	'Несоответсвие кодировки карты незарегистрированного новорожденного и d_type(9)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
    IF people.d_type='9' AND OCCURS('#', ALLTRIM(c_i))>0
     m.c_i = ALLTRIM(c_i)
     IF OCCURS('#', m.c_i)!=3
      m.recid = recid
      rval =InsError('S', 'H6A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ELSE 
      DO CASE 
       CASE LEN(SUBSTR(m.c_i,1,AT('#',m.c_i)-1))>12
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid, '', ;
     		'Неверная кодировка карты незарегистрированного новорожденного (длина номера истории болезни >12)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       CASE OCCURS('.',SUBSTR(m.c_i,1,AT('#',m.c_i)-1))>0
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid, '', ;
     		'Недопустимые символы в истории болезни новорожденного (".")')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       
       CASE !INLIST(SUBSTR(m.c_i,AT('#',m.c_i)+1,1),'1','2')
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid, '', ;
     		'Неверная кодировка карты незарегистрированного новорожденного (неверный пол новорожденного)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       CASE EMPTY(CTOD(SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),5,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),1,4)))
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid, '', ;
     		'Неверная кодировка карты незарегистрированного новорожденного (неверная дата рождения новорожденного)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       *CASE !INLIST(SUBSTR(m.c_i,AT('#',m.c_i,3)+1,1),'1','2','3','4','5','6','7','8','9')
       CASE !BETWEEN(INT(VAL(SUBSTR(m.c_i,AT('#',m.c_i,3)+1))),1,9)
        m.recid = recid
        rval =InsError('S', 'H6A', m.recid, '', ;
       	'Неверная кодировка карты незарегистрированного новорожденного (порядковый номер новорожденного <1 или >9)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       OTHERWISE 
      ENDCASE 
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.COA == .T. && Алгоритм CO 
    * Ревизия от 01.11.2019, ПФО август 2019_1_8652455_v1
    * OK!
    m.otd = ALLTRIM(otd )
    DO CASE 
     CASE EMPTY(m.otd)
      m.recid = recid
      rval =InsError('S', 'COA', m.recid, '', 'Пустое поле otd')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

     CASE LEN(m.otd)<>8
      m.recid = recid
      rval =InsError('S', 'COA', m.recid, '', 'Длина поля otd не равна 8 символам')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

     CASE !INLIST(LEFT(m.otd,1),'1','2','3')
      m.recid = recid
      rval =InsError('S', 'COA', m.recid, '', 'Первый символ кода отделения не равен 1,2,3')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

     CASE !SEEK(SUBSTR(m.otd,2,2), 'profot')
      m.recid = recid
      rval =InsError('S', 'COA', m.recid, '', '2,3 позиция кода отделения не соответствует справочнику profot')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

     CASE !SEEK(SUBSTR(m.otd,4,3), 'prv002')
      m.recid = recid
      rval =InsError('S', 'COA', m.recid, '', '4,5,6 позиции кода отделения не соответствует справочнику prv002')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      
     CASE SUBSTR(m.mcod,2,1)='1' AND LEFT(m.otd,1)='2'
      m.recid = recid
      rval =InsError('S', 'COA', m.recid, '', 'Детское отделение во взрослом МО (с 02.2020)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

     CASE SUBSTR(m.mcod,2,1)='2' AND LEFT(m.otd,1)='1'
      m.recid = recid
      rval =InsError('S', 'COA', m.recid, '', 'Взрослое отделение в детском МО (с 02.2020)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

     OTHERWISE 

    ENDCASE 
   ENDIF 

   IF M.MPA == .T.
    IF USED('deads')
     m.cod    = cod
     m.c_i    = c_i
     IF SEEK(LEFT(m.sn_pol,17), 'deads') AND OCCURS('#', m.c_i)<3
      m.cod    = cod 
      m.d_date = deads.d_u
      DO CASE 
       CASE INLIST(INT(m.cod/1000),59,159) AND m.d_u - m.d_date < 60
       CASE INLIST(INT(m.cod/1000),25,125,26,126,27,127,29,129,30,130) AND m.d_u - m.d_date < 60
       CASE INLIST(INT(m.cod/1000),28,128) AND m.d_u - m.d_date<60 AND m.lpuid=1795
       OTHERWISE 
        m.recid = recid
        rval =InsError('S', 'MPA', m.recid, '', ;
     		'Услуга оказана после смерти пациента: mcod='+deads.mcod+', d_u=)' + DTOC(deads.d_u) + ;
     		', tip='+deads.tip+', d_type='+deads.d_type)
      	m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDCASE 
     ENDIF
    ENDIF 
   ENDIF 
   
   IF M.HRA = .T.
    m.cod = cod 
    m.otd = otd
    m.sn_pol = sn_pol
    m.d_u = d_u
    IF SUBSTR(m.otd,2,2)='91' AND !INLIST(m.cod,15001,115001) AND (!SEEK(m.sn_pol, 'polic_h') OR ;
    	SEEK(m.sn_pol, 'polic_h') AND m.d_u<polic_h.d_u)
     m.recid = recid
     rval =InsError('S', 'HRA', m.recid, '', ;
     'Регистрация услуг центра здоровья, оказанных пациенту, не прошедшему первичную регистрацию (15001/115001)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
   ENDIF 
   
   IF M.D1A = .T. AND USED('dc_du')
    m.cod = cod
    m.c_i = c_i
    m.d_u = d_u

    IF (INLIST(LEFT(m.c_i,3),'ДД_','ДУ_','ДС_') OR INLIST(LEFT(m.c_i,5),'ПРОФВ','ПРОФД')) AND m.d_u>{03.04.2020}
     m.recid = recid
     rval =InsError('S', 'D1A', m.recid, '', ;
      	'Медицинская услуги по профилактическим осмотрам и/или диспансеризации оказана после {03.04.2020}')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF m.cod = 101952
     DO CASE 
      CASE LEFT(m.c_i,3)='ДУ_'
       IF !SEEK(m.lpuid, 'dc_du', 'du')
        m.recid = recid
        rval =InsError('S', 'D1A', m.recid, '', ;
        	'Услуга 101952 '+ALLTRIM(m.c_i)+'оказана в МО не из справочника dc_du.dbf')
      	m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      CASE LEFT(m.c_i,3)='ДС_'
       IF !SEEK(m.lpuid, 'dc_du', 'dc')
        m.recid = recid
        rval =InsError('S', 'D1A', m.recid, '', ;
        	'Услуга 101952 '+ALLTRIM(m.c_i)+'оказана в МО не из справочника dc_du.dbf')
      	m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
     ENDCASE 
    ENDIF 
    
   ENDIF 

   IF M.NDA == .T.
    IF USED('msext')
     m.cod = cod
     m.ord = ord
     m.d_type = d_type
     m.tip = tip
     m.c_i = c_i
     IF IsMes(m.cod) && OR IIF(VAL(m.gcPeriod)<=201910, 1=1, IsVmp(m.cod))
      IF m.ord=3 AND !((SEEK(m.cod, 'msext') OR INLIST(m.tip,'3','5') OR INLIST(m.d_type,'3','5')) OR (!SEEK(m.c_i, 'Gosp', 'c_i')))
       IF !(INLIST(m.lpuid,4511,2293,4586) AND ;
       	INLIST(m.cod,61400,161400,161401,61430,161430,161431,70150,70160,70170,170150,170151,170160,170161,170170,170171)) && Добавлено из-за covid!
        m.recid = recid
        rval =InsError('S', 'NDA', m.recid, '', ;
      	'Ord=3 при использовании МЭС ' + STR(m.cod,6)+', не включенном в справочник msext')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.MZA == .T.
    IF USED('stop')
     IF SEEK(LEFT(m.sn_pol,17), 'stop')
      m.cod    = cod 
      m.d_date = stop.date_end
      m.d_u    = d_u
      DO CASE 
       CASE INLIST(INT(m.cod/1000),59,159) AND m.d_u - m.d_date < 60
       CASE INLIST(INT(m.cod/1000),25,125,26,126,27,127,29,129,30,130) AND m.d_u - m.d_date < 60
       CASE INLIST(INT(m.cod/1000),28,128) AND m.d_u - m.d_date<60 AND m.lpuid=1795
       OTHERWISE 
        m.recid = recid
        rval =InsError('S', 'MZA', m.recid, '', ;
     	 'Услуга оказана после смерти пациента по данным ЦС ЕРЗ: ' + DTOC(m.d_date))
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDCASE 
     ENDIF
    ENDIF 
   ENDIF 
   
   IF M.NVA
    m.pcod = pcod 
    m.d_u = d_u 
    IF !SEEK(m.pcod, 'doctor')
     m.recid = recid
     rval =InsError('S', 'NVA', m.recid, '', ;
     	'Код специалиста (: ' + ALLTRIM(m.pcod)+') не найден в справочнике')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE 
     DO CASE 
      CASE EMPTY(doctor.d_prik)
       rval =InsError('S', 'NVA', m.recid, '', ;
     	'Неполнота данных о специалисте ('+ ALLTRIM(m.pcod)+'), не заполнено поле d_prik')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE doctor.d_prik>m.d_u
       rval =InsError('S', 'NVA', m.recid, '', ;
     	'Дата оказания услуги специалистом '+ ALLTRIM(m.pcod)+'раньше, чем приказ о его/её назначении'+DTOC(doctor.d_prik))
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      
     ENDCASE 
    ENDIF 
   ENDIF 

   IF M.PLA == .T.
    IF !EMPTY(vnov_m) AND !BETWEEN(vnov_m, 300, 2500)
     m.recid = recid
     rval =InsError('S', 'PLA', m.recid, '', ;
     	'VNOV_M заполнено, но не входи в диапазон 300 - 2500 (vnov_m='+STR(vnov_m,4)+')')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.H7A == .T.
    m.cod = cod 
    IF INLIST(FLOOR(m.cod/1000),300,397)
     m.recid = recid
     rval =InsError('S', 'H7A', m.recid, '', ;
     	'300/397 коды относятся к сверхбазовой программе ОМС, подача недопустима')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF

    IF INLIST(m.cod, 1719,8050,8051,8052,26281,28210,31001,31002,31003,40040,40041,40042,40043,40044,40045)
     m.recid = recid
     rval =InsError('S', 'H7A', m.recid, '', ;
     	'Услуга пренатальной диагностики сверх базовой программы ОМС')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.HCA == .T. && Алгоритм HC
    * Ревизия от 01.11.2019, ПФО август 2019_1_8652455_v1
    * OK!
    IF k_u <= 0
     m.recid = recid
     rval = InsError('S', 'HCA', m.recid, '',;
     	'Значение поля k_u='+ALLTRIM(STR(k_u))+' (меньше либо равно нулю)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
    IF !IsUSL(cod) AND kd_fact <= 0
     m.recid = recid
     rval = InsError('S', 'HCA', m.recid, '',;
     	'Значение поля kd_fact='+ALLTRIM(STR(kd_fact))+' (меньше либо равно нулю)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF

   ENDIF 

   m.o_otd   = SUBSTR(otd,2,2)
   m.usl_ok  = IIF(SEEK(m.o_otd, 'profot'), profot.usl_ok, '0')
   m.is_gsp  = IIF(m.usl_ok='1', .T., .F.)

   *IF M.OGA == .T. AND M.O0A == .T. AND m.is_gsp && Алгоритм OG
   IF M.OGA == .T. AND M.O0A == .T. && m.is_gsp включен со счетов за Апрель!
    m.recid  = recid
    m.ds_onk = ds_onk
    m.rslt   = rslt
    IF !INLIST(m.ds_onk,0,1)
     m.recid = recid
     rval = InsError('S', 'OGA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
    IF m.ds_onk=1 AND INLIST(m.rslt,317,321,332,343,347)
     m.recid = recid
     rval = InsError('S', 'OGA', m.recid, '',;
     	'ds_onk=1 при rslt='+STR(m.rslt,3)+' из перечня 317,321,332,343,347')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   *IF M.O4A == .T. AND M.O0A == .T.  AND m.is_gsp && Алгоритм O4
   IF M.O4A == .T. && AND M.O0A == .T. && ошибка включена со счетов за сентябрь по требованию Согаза
    m.recid  = recid
    m.p_cel  = p_cel
    m.o_otd  = SUBSTR(otd,2,2)
    m.usl_ok = IIF(SEEK(m.o_otd, 'profot'), profot.usl_ok, '0')
    m.is_amb = IIF(m.usl_ok='3', .T., .F.)

    IF m.is_amb AND !SEEK(m.p_cel, 'onpcel')
     m.recid = recid
     rval = InsError('S', 'O4A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   *IF M.OHA == .T. AND M.O0A == .T. AND m.is_gsp  && Алгоритм OH
   IF M.OHA == .T. && AND M.O0A == .T. ошибка включена со счетов за сентябрь по требованию Согаза
    m.recid = recid
    m.p_cel = p_cel
    m.dn    = dn
    IF m.p_cel = '1.3' AND !INLIST(m.dn,1,2,3,4)
     m.recid = recid
     rval = InsError('S', 'OHA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   *IF M.OIA == .T. AND M.O0A == .T. AND m.is_gsp  && Алгоритм OI
   IF M.OIA == .T. AND M.O0A == .T.
    m.recid = recid
    m.reab  = reab
    IF !INLIST(m.reab,0,1)
     m.recid = recid
     rval = InsError('S', 'OIA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.ENA == .T. && Алгоритм EN
    m.recid = recid
    m.cod   = cod
    m.tal_d = tal_d
    m.c_i   = c_i
    IF OCCURS('#', m.c_i)=3
     m.dr = CTOD(SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),5,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),1,4))
    ELSE 
     m.dr = IIF(SEEK(m.sn_pol, 'people'), people.dr, {})
    ENDIF 
    *m.d_beg = IIF(SEEK(m.sn_pol, 'people'), people.d_beg, {})
    m.d_u = d_u
    m.k_u = k_u
    m.d_beg = m.d_u - m.k_u
    
    IF IsVmp(m.cod)
     IF EMPTY(m.tal_d)
      m.recid = recid
      rval = InsError('S', 'ENA', m.recid, '',;
     	'Для ВМП (код '+PADL(m.cod,6,'0')+') не заполнено поле tal_d (дата выдачи талона-направления)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ELSE 
      IF m.tal_d<m.dr-1
       m.recid = recid
       rval = InsError('S', 'ENA', m.recid, '',;
     	'Для ВМП (код '+PADL(m.cod,6,'0')+') талон-направление выдан ('+DTOC(m.tal_d)+') ранее даты рождения пациента ('+DTOC(m.dr)+')')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      *IF m.tal_d>m.d_beg
      IF m.tal_d>m.d_u-IIF(m.k_u>1, m.k_u, 0)
       m.recid = recid
       rval = InsError('S', 'ENA', m.recid, '',;
     	'Для ВМП (код '+PADL(m.cod,6,'0')+') талон-направление выдан ('+DTOC(m.tal_d)+') после начала лечения ('+DTOC(m.d_beg)+')')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF
    ENDIF 
   ENDIF 

   *IF M.O3A == .T. AND M.O0A == .T. AND m.is_gsp  && Алгоритм O3
   IF M.O3A == .T. AND M.O0A == .T.
    m.recid = recid
    m.napr_v_in =  napr_v_in
    IF !EMPTY(m.napr_v_in) AND !SEEK(m.napr_v_in, 'onnapr')
     m.recid = recid
     rval = InsError('S', 'O3A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   *IF M.O5A == .T. AND M.O0A == .T. AND m.is_gsp  && Алгоритм O3
   IF M.O5A == .T. AND M.O0A == .T.
    m.ds   = ds
    m.ds_2 = ds_2
    m.ds_onk = ds_onk
    m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) OR ;
  		BETWEEN(LEFT(m.ds,3),'D00','D09') , .T., .F.)

    IF m.IsOnkDs
     m.recid = recid
     m.c_zab = c_zab
     *IF !EMPTY(m.c_zab) AND !SEEK(m.c_zab, 'onczab')
     IF !SEEK(m.c_zab, 'onczab')
      m.recid = recid
      rval = InsError('S', 'O5A', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 
   ENDIF 

   IF M.DUA == .T. && Част алгоритма DU
    m.c_i = ALLTRIM(c_i)
    m.d_u = d_u 
    m.cod = cod
    m.usl_ok  = IIF(SEEK(m.o_otd, 'profot'), profot.usl_ok, '0')

    IF IsUsl(m.cod)
     IF OCCURS('#', m.c_i)=3
      m.d_dr = CTOD(SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),5,2)+'.'+;
	 	SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),1,4))
	  IF !EMPTY(m.d_dr)
	   IF GOMONTH(m.d_dr,IIF(m.usl_ok='3',5,6))<m.d_u
        m.recid = recid
        rval = InsError('S', 'DUA', m.recid, '',;
     	 'Дата оказания услуги незарегистрированному новорожденному не соответствовует периоду новорожденности.')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	   ENDIF 
	  ENDIF 
	 ENDIF 
	 RELEASE d_dr
    ENDIF 

    IF (people.tip_p==1 AND MONTH(d_u)!=tMonth)
     m.recid = recid
     rval = InsError('S', 'DUA', m.recid, '',;
     	'Дата оказания не в периоде для амбулаторного пациента')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 

   IF M.H8A == .T. && Част алгоритма H8
    IF !SEEK(ds, 'mkb10')
     m.recid = recid
     rval = InsError('S', 'H8A', m.recid, '', ;
     	'Код диагноза не найден в справочнике mkb10')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE 
     IF FIELD('isoms','mkb10') = 'ISOMS'
      m.ds  = ds
      m.ds_2 = ds_2
      m.ds_3 = ds_3
      IF !mkb10.IsOms
       m.recid = recid
       rval = InsError('S', 'H8A', m.recid, '', ;
     	'Код диагноза не входит в Территориальную программу ОМС (с 01.09.2020)')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      IF !EMPTY(m.ds_2)
       IF !SEEK(m.ds_2, 'mkb10')
        m.recid = recid
        rval = InsError('S', 'H8A', m.recid, '', ;
     		'Код диагноза ds_2 '+ALLTRIM(m.ds_2)+' не найден в справочнике mkb10')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ELSE 
        IF !mkb10.IsOms
         m.recid = recid
         rval = InsError('S', 'H8A', m.recid, '', ;
     	  'Код диагноза ds_2 '+ALLTRIM(m.ds_2)+' не входит в Территориальную программу ОМС (с 01.09.2020)')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       ENDIF 
      ENDIF 
      IF !EMPTY(m.ds_3)
       IF !SEEK(m.ds_3, 'mkb10')
        m.recid = recid
        rval = InsError('S', 'H8A', m.recid, '', ;
     		'Код диагноза ds_3 '+ALLTRIM(m.ds_3)+' не найден в справочнике mkb10')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ELSE 
        IF !mkb10.IsOms
         m.recid = recid
         rval = InsError('S', 'H8A', m.recid, '', ;
     	  'Код диагноза ds_3 '+ALLTRIM(m.ds_3)+' не входит в Территориальную программу ОМС (с 01.09.2020)')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       ENDIF 
      ENDIF 
     ENDIF 
     IF INLIST(LEFT(ds,3),'B95','B96','B97')
      m.recid = recid
      rval = InsError('S', 'H8A', m.recid, '', ;
     	'Код диагноза входит в перечень неиспользуемых в ОМС (B95.xx,B96.xx,B97.xx)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     m.cod = cod
     m.ds  = ds
     IF (LEFT(m.ds,1)='Z' AND !INLIST(m.ds,'Z13.8','Z01.7','Z20','Z34','Z35','Z11.5','Z03.8','Z22.8')) AND ;
     	INLIST(FLOOR(m.cod/1000),25,26,27,28,29,30,125,126,127,128,129,130)
      m.recid = recid
      rval = InsError('S', 'H8A', m.recid, '', ;
     	'При Z-диагнозах кроме Z13.8,Z01.7,Z20,Z34,Z35,Z11.5,Z03.8,Z22.8 не может быть услуг из групп 25,26,27,28,29,30,125,126,127,128,129,130')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     IF EMPTY(Tip)
      IF LEN(ALLTRIM(ds))=3 AND SEEK(ALLTRIM(ds)+'.','mkb10')
       m.recid = recid
       rval = InsError('S', 'H8A', m.recid, '', ;
     	'Использование рубрики при наличии подрубрики')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
     
     m.c_i    = c_i
     m.sn_pol = sn_pol
     m.d_u    = d_u
     m.dr     = IIF(SEEK(m.sn_pol, 'people'), people.dr, {})
     m.adj    = 0
     m.vozr   = (YEAR(m.d_u) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)
     
     IF ds='P' AND  m.vozr>=18 AND OCCURS('#', m.c_i)<>3
       m.recid = recid
       rval = InsError('S', 'H8A', m.recid, '', ;
     	'Использование диагноза PXX.XX у совершеннолетнего (>= 18 лет)')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 

     m.ds_2 = ds_2
     m.ds_3 = ds_3
      
     DO CASE 
      CASE m.ds = m.ds_2
       m.recid = recid
       rval = InsError('S', 'H8A', m.recid, '', ;
     	'Основной диагноз (DS) '+m.ds+' совпадает с сопутcтвующим диагнозом (DS_2) '+m.ds_2)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      CASE m.ds = m.ds_3
       m.recid = recid
       rval = InsError('S', 'H8A', m.recid, '', ;
     	'Основной диагноз (DS) '+m.ds+' совпадает с диагнозом осложнения (DS_3) '+m.ds_3)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      CASE !EMPTY(m.ds_2) AND !EMPTY(m.ds_3) AND m.ds_2=m.ds_3
       m.recid = recid
       rval = InsError('S', 'H8A', m.recid, '', ;
     	'Сопутсвующий диагноз (DS_2) '+m.ds_2+' совпадает с диагнозом осложнения (DS_3) '+m.ds_3)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      OTHERWISE 
      
     ENDCASE 
    ENDIF
   ENDIF 

   IF M.HEA == .T.  && Алгоритм HE
    m.sex = IIF(OCCURS('#',c_i)==3, SUBSTR(c_i, AT('#',c_i,1)+1, 1), STR(people.w,1))
    IF (SEEK(ds, 'mkb10') AND !EMPTY(mkb10.sex)) AND m.sex != mkb10.sex
     m.d_type = d_type
     IF m.d_type<>'4'
      m.recid = recid
      rval = InsError('S', 'HEA', m.recid, '', ;
     	'Несоответствие пола пациента ('+m.sex+') его диагнозу по справочнику mkb10 ('+mkb10.sex+')')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF
   ENDIF 
   
   IF M.CSA == .T.
    IF !SEEK(cod, 'tarif')
     m.recid = recid
     rval = InsError('S', 'CSA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 
   
   IF M.KEA == .T.
    m.d_type = d_type
    m.cod    = cod
    m.tip    = tip
    m.otd    = otd
    m.c_i    = c_i
    m.sn_pol = sn_pol
    m.d_u    = d_u
    m.usl_ok  = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), profot.usl_ok, '0')
    m.lpu_ord = lpu_ord
    m.lpu_ord = IIF(m.lpu_ord>9999 AND FLOOR(m.lpu_ord/10000)=77, m.lpu_ord%10000, m.lpu_ord)
    m.rslt    = rslt
    
    IF m.d_type='s'
     DO CASE 
      CASE !INLIST(INT(m.cod/1000),51,52,53,54,55,151,152,153,154,155)
       m.recid = recid
       rval = InsError('S', 'KEA', m.recid, '', ;
       	'd_type=s при коде услуги ('+STR(m.cod,6)+') не из перечня симультанных услуг (51-55,151-155)')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      CASE !USED('Gosp') && это не стационар вообще!
       m.recid = recid
       rval = InsError('S', 'KEA', m.recid, '', ;
       	'd_type=s при коде услуги ('+STR(m.cod,6)+') при отсутствии госпитализации (не стационар)')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      CASE !SEEK(m.sn_pol, 'Gosp', 'sn_pol')
       m.recid = recid
       rval = InsError('S', 'KEA', m.recid, '', ;
       	'd_type=s при коде услуги ('+STR(m.cod,6)+') при отсутствии госпитализации')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      CASE !SEEK(m.c_i, 'Gosp', 'karta')
       m.recid = recid
       rval = InsError('S', 'KEA', m.recid, '', ;
       	'Несоответствие истории болезни симультанной услуги ('+STR(m.cod,6)+') и госпитализации')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      OTHERWISE 
       SET ORDER TO karta IN Gosp
       m.is_ok=.F.
       DO WHILE m.c_i=Gosp.c_i
        IF IsMES(Gosp.cod) OR IsVMP(Gosp.cod) OR INLIST(Gosp.cod,56029,156003)
         IF Gosp.k_u>1
          IF BETWEEN(m.d_u, Gosp.d_u-Gosp.k_u, Gosp.d_u)
           IF INLIST(Gosp.d_type,'1','3','5','6')
            m.is_ok=.T.
           ENDIF 
          ENDIF 
         ELSE 
          IF BETWEEN(m.d_u, Gosp.d_u-1, Gosp.d_u)
           IF INLIST(Gosp.d_type,'1','3','5','6')
            m.is_ok=.T.
           ENDIF 
          ENDIF 
         ENDIF 
        ENDIF 
        SKIP IN Gosp
       ENDDO 
       IF !m.is_ok
        m.recid = recid
        rval = InsError('S', 'KEA', m.recid, '', ;
       	 'Симультанная услуга оказана вне периода госпитализации либо d_type!=1')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 

     ENDCASE 
    ENDIF 
    
    IF (INLIST(m.cod,56029,156003) AND INLIST(SUBSTR(m.otd,4,3),'005','167')) AND !INLIST(m.d_type,'3','5')
     m.recid = recid
     rval = InsError('S', 'KEA', m.recid, '', ;
     	'Для услуги '+STR(m.cod,6)+' (реанимация <12 часов) d_type='+m.d_type+' и не равен 3/5')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF m.cod=1561 AND m.d_type<>'5'
     m.recid = recid
     rval = InsError('S', 'KEA', m.recid, '', ;
     	'Для услуги '+STR(m.cod,6)+' (констатация факта смерти) d_type='+m.d_type+' и не равен 5')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF m.d_type='w' AND SUBSTR(m.otd,2,3)<>'93'
     m.recid = recid
     rval = InsError('S', 'KEA', m.recid, '', ;
     	'd_type=w (услуга '+STR(m.cod,6)+' при фасетном коде отделения ('+SUBSTR(m.otd,2,3)+'), отличном от 93')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF SUBSTR(m.otd,2,3)='93' AND m.d_type<>'w'
     m.recid = recid
     rval = InsError('S', 'KEA', m.recid, '', ;
     	'd_type!=w (услуга '+STR(m.cod,6)+' при фасетном коде отделения '+SUBSTR(m.otd,2,3))
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF (INLIST(INT(m.cod/1000),49,149) OR INLIST(m.cod,97010,197010)) AND INLIST(m.usl_ok,'1') AND m.d_type<>'2'
     m.recid = recid
     rval = InsError('S', 'KEA', m.recid, '', ;
     	'd_type!=2 для услуги '+STR(m.cod,6))
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF (INLIST(m.cod,36022,36023,36024,136022,136023,136024)) AND m.d_type<>'2'
     m.recid = recid
     rval = InsError('S', 'KEA', m.recid, '', ;
     	'd_type!=2 для услуги '+STR(m.cod,6))
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
	IF m.d_type='R' AND !(INLIST(SUBSTR(m.otd,2,2),'69','82','87','88','89') AND SUBSTR(m.otd,4,3)='158')
     m.recid = recid
     rval = InsError('S', 'KEA', m.recid, '', ;
     	'd_type=r для услуги '+STR(m.cod,6)+'в непрофильном отделении (д.б.69,82,87,88,89 + 158)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	ENDIF
	IF m.d_type='6' AND m.lpu_ord<>m.lpuid
     m.recid = recid
     rval = InsError('S', 'KEA', m.recid, '', ;
     	'd_type=6 при lpu_ord ('+STR(m.lpu_ord,4)+'), отличным от lpuid ('+STR(m.lpuid,4))
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	ENDIF 
	IF IsMes(m.cod) AND (m.d_type='5' AND !INLIST(m.rslt,105,106, 205,206,313))
     m.recid = recid
     rval = InsError('S', 'KEA', m.recid, '', ;
     	'd_type=5 при rslt ('+STR(m.rslt,3)+'), отличным от 105,106, 205,206,313')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	ENDIF 
   ENDIF 

   IF M.TFA == .T.
    m.otd    = otd 
    m.tip    = tip
    m.k_u    = k_u
    m.cod    = cod
    m.sn_pol = sn_pol
    m.c_i    = c_i
    m.vir2   = m.sn_pol + m.c_i
    m.d_type = d_type
    m.rslt   = rslt

    IF SEEK(cod, 'tarif')
     IF !EMPTY(Tarif.n_kd) AND !SEEK(Tip, 'kpresl')
      m.recid = recid
      rval = InsError('S', 'TFA', m.recid, '',;
      	'Пустое поле Tip для МЭС')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 

    IF IsVMP(m.cod) AND m.tip != 'v'
     m.recid = recid
     rval = InsError('S', 'TFA', m.recid, '',;
     	'Tip='+m.tip+' для ВМП ('+STR(m.cod,6)+')')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF !IsVMP(m.cod) AND m.tip='v'
     m.recid = recid
     rval = InsError('S', 'TFA', m.recid, '',;
     	'Tip='+m.tip+' для видов МП, не относящимся к ВМП ('+STR(m.cod,6)+')')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF m.tip='5' AND !INLIST(m.rslt,105,106,205,206)
     m.recid = recid
     rval = InsError('S', 'TFA', m.recid, '',;
     	'Tip='+m.tip+' при rslt<>105,106,205,206 ('+STR(m.rslt,3)+')')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    *IF INLIST(m.cod, 61400,161400,161401,70150,70160,70170,170150,170151,170160,170161,170170,170171,61410,161410,161411) AND m.Tip='7'
    * m.recid = recid
    * rval = InsError('S', 'TFA', m.recid, '',;
    * 	'Tip='+m.tip+' недопустим для данного МЭС ('+STR(m.cod,6)+')')
    * m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    *ENDIF 

     DO CASE 
      CASE !(IsMes(m.cod) OR IsVMP(m.cod))
      CASE INLIST(m.cod, 76170, 76171, 76180, 76182, 76190, 76192, 76200, 76210, 76211, 76212, 76411, ;
      	76431, 76490, 76500, 76521, 76530, 76540, 76550)
      CASE INLIST(m.cod, 76560, 76570, 76581, 76582, 76640, 76650, 76660, 76670, 76681, 76683, 76691, ;
      	76711, 76721, 76810, 76830, 76870, 76880, 76891) && родовспоможение или прерывание беременности 
      CASE INLIST(m.cod,68260, 168260) && замена речевого процессора
      CASE INLIST(m.cod, 77180, 177180, 77181, 177181, 82010) 
      CASE INLIST(m.cod, 63020,80010,81094)  && три мэс по письму 
      CASE INLIST(m.cod, 68260,168260) AND d_type='5'
      CASE INLIST(INT(m.cod/1000), 84, 184, 92, 192, 97)
      
      CASE INLIST(m.cod, 64160, 69095, 69180, 69190, 69191, 72081, 72182, 72380, 72381, 72401, 76082, 76092, ;
      	76120, 168050, 168051, 168260, 172240, 172241, 173220, 173221, 190481, 190483)
      ** Доделать поиск симультанных услуг (51-55/151-155)!
      
      OTHERWISE 

       IF m.tip='A' AND !INLIST(SUBSTR(m.otd,4,3),'004','012','018','060','077')
        m.recid = recid
        rval = InsError('S', 'TFA', m.recid, '',;
     		'Tip=A в отделениях, не соответствующих допустимым профилям: химиотерапии, ;
     			гематологии и ревматологии (4-6 разряды фасетного кода отделения 004,012,018,060,077)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
       IF m.tip='R' AND !(INLIST(SUBSTR(m.otd,2,2),'69','82','87','88','89') AND SUBSTR(m.otd,4,3)='158')
        m.recid = recid
        rval = InsError('S', 'TFA', m.recid, '',;
     		'Tip=R в отделениях, не соответствующих профилю медицинской реабилитации стационара (2,3 разряды ;
     	 	фасетного кода отделения не равны 69,82,87,88,89, 4-6 разряды не равны 158)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
       IF m.tip='R' AND (INLIST(SUBSTR(m.otd,2,2),'69','82','87','88','89') AND SUBSTR(m.otd,4,3)='158') AND ;
       		m.k_u<14
        m.recid = recid
        rval = InsError('S', 'TFA', m.recid, '',;
     		'Tip=R ранее 14 дней лечения в отд. медицинской реабилитации  стационара')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
       IF m.tip='T' AND !INLIST(SUBSTR(m.otd,2,2),'00','01','08','09','10','69','70','73','80','81','82','85','87','88','99') AND ;
    		m.k_u<10
        m.recid = recid
        rval = InsError('S', 'TFA', m.recid, '',;
     		'Tip=T ранее 10 дней лечения пациента в профильном отделении стационара')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 

       IF SEEK(STR(m.cod,6)+' '+m.tip, 'tipnomes')
        m.recid = recid
        rval = InsError('S', 'TFA', m.recid, '',;
     		'Регистрация в счете сочетания Tip='+m.tip+' и МЭС ('+STR(m.cod,6)+'), включенного в справочник tipno_xx.dbf')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 

     ENDCASE 

   ENDIF 

   IF M.TLA == .T.
    m.recid = recid
    m.cod = cod
    m.k_u = k_u
    DO CASE 
     CASE m.cod = 83010 
      IF !BETWEEN(m.k_u,1,2)
       rval = InsError('S', 'TLA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     CASE m.cod = 83020
      IF !BETWEEN(m.k_u,3,4)
       rval = InsError('S', 'TLA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     CASE m.cod = 83030
      IF !BETWEEN(m.k_u,5,6)
       rval = InsError('S', 'TLA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     CASE m.cod = 83040
      IF !BETWEEN(m.k_u,7,8)
       rval = InsError('S', 'TLA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     CASE m.cod = 83050
      IF m.k_u<9
       rval = InsError('S', 'TLA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
    ENDCASE 
    IF (INLIST(m.cod,97041,97013,197013) OR BETWEEN(cod, 84000, 84999)) AND m.k_u>1
     m.recid = recid
     rval = InsError('S', 'TLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    IF INLIST(m.cod,56029,156003) AND m.k_u>1 && С декабря 2019
     m.recid = recid
     rval = InsError('S', 'TLA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
   ENDIF 

   IF M.THA == .T.
    m.cod = cod
    m.otd = otd
    m.c_i    = c_i
    m.sn_pol = sn_pol
    m.tip = tip

    IF IsMes(m.cod) AND m.IsOtdSkp
      IF !USED('ho')
       m.recid = recid
        rval = InsError('S', 'THA', m.recid, '',;
        	'МЭС применен в отделении СКП (2,3 позиция фасетного кода отделения 09) без оперативного пособия ;
        	(отсутствует файл ho)')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ELSE 
       m.vir = m.sn_pol + m.c_i + PADL(m.cod,6,'0')
       IF !SEEK(m.vir, 'ho')
        m.recid = recid
        rval = InsError('S', 'THA', m.recid, '',;
        	'МЭС применен в отделении СКП (2,3 позиция фасетного кода отделения 09) без оперативного пособия ;
        	(отсутствует соответствующая запись в файле ho)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
    ENDIF 

    IF INLIST(m.tip,'0','A','T','R')
     IF INLIST(FLOOR(m.cod/1000),72,73,75,76,79,82,85,90,172,173,175,176,179,182,185,190) ;
     	AND !SEEK(m.cod, 'noth')
     DO CASE 
      CASE INLIST(m.cod, 72040, 72070, 72071, 72090, 72140, 72150, 72190, 72291, 72320, 72380, 72400, 72420, 72460)
      CASE INLIST(m.cod, 73120, 73130, 75070, 75140, 75150, 76160, 76200)
      CASE BETWEEN(cod,76411,76891)
      CASE INLIST(cod, 79010, 79040, 79060, 79120, 79130, 79131, 79140, 79150, 79160, 79260, 79270, 79271)
      CASE INLIST(m.cod, 82010, 90460, 90470, 90480, 90490, 90500, 90510, 90520, 172030, 172031, 172070, 172071)
      CASE INLIST(m.cod, 172110, 172111, 172190, 172191, 172230, 172231, 172240, 172241, 172270, 172271, 172290)
      CASE INLIST(m.cod, 172291, 175010, 175011, 175020, 175021, 175060, 175061, 175100, 175101, 176100, 176101, 176070)
      CASE INLIST(m.cod, 176071, 176110, 176111, 176150, 176151, 179050, 179051, 179070, 179071, 179170, 179171)
      CASE INLIST(m.cod, 179290, 179291, 179330, 179331, 190220, 190221, 190260, 190261, 190460, 190461, 190480)
      CASE INLIST(m.cod, 190481, 190482, 190483, 190490, 190491, 190500, 190501, 190510, 190511, 190520, 190521, 190530, 190531)
      OTHERWISE 

       m.vir  = m.sn_pol + m.c_i + PADL(m.cod,6,'0')
       m.vir2 = m.sn_pol + m.c_i

       IF !USED('ho')
        m.recid = recid
        rval = InsError('S', 'THA', m.recid, '',;
       	 'Tip='+m.tip+' указан для МС разделов 72/172, 73/173, 75/175, 76/176, 79/179, 82/182, 85/185, 90/190  ;
       		при отсутствии выполненного оперативного пособия (отсутствует файл ho)')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ELSE 
        IF !SEEK(m.vir, 'ho') AND !SEEK(m.vir2, 'ho', 'snp_ci')
         m.recid = recid
         rval = InsError('S', 'THA', m.recid, '',;
       	  'Tip='+m.tip+' указан для МС разделов 72/172, 73/173, 75/175, 76/176, 79/179, 82/182, 85/185, 90/190  ;
       		при отсутствии выполненного оперативного пособия (отсутствует запись в файле ho)')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       ENDIF 
      ENDCASE 
     ENDIF 
    ENDIF 
   ENDIF 

   *IF M.HOA == .T.
   IF M.S1A == .T.
    IF IsMes(m.cod) AND m.IsOtdSkp
     IF USED('ho')
      m.sn_pol = sn_pol
      m.c_i    = c_i
      m.cod    = cod
      
      m.vir = m.sn_pol + m.c_i + PADL(m.cod,6,'0')
      IF SEEK(m.vir, 'ho')
       m.cod   = cod
       m.codho = ho.codho
       m.ds    = ds
       
       *m.vir   = PADL(m.cod,6,'0') + m.codho + LEFT(m.ds,5)
       m.vir   = PADL(m.cod,6,'0') + m.codho + IIF(!ISDIGIT(SUBSTR(m.ds,5,1)), LEFT(Ds,3)+'   ', LEFT(m.ds,5)+' ')

       IF !SEEK(m.vir, 'reeskp', 'unik')
        m.recid = recid
        rval = InsError('S', 'S1A', m.recid, '', ;
        	'Сочетание код СКП+код операции+диагноз '+PADL(m.cod,6,'0')+' '+m.codho+' '+LEFT(m.ds,5)+' не найдено в справочнике reeskp')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ELSE 
       m.recid = recid
       rval = InsError('S', 'S1A', m.recid, '', ;
       	'Сочетание полис+карта+код СКП '+ALLTRIM(m.sn_pol)+' '+ALLTRIM(m.c_i)+' '+PADL(m.cod,6,'0')+;
       		' не найдено в справочнике примененных операций (ho-файл)')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ELSE 
      m.recid = recid
      rval = InsError('S', 'S1A', m.recid, '', ;
      	'Отсутствует справочник ho (услуга оказан в отделени СКП, оперативное вмешательство должно быть')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.TVA == .T.
    m.tiplpu = SUBSTR(m.mcod,2,1)
    IF m.tiplpu!='3'
     IF SEEK(cod, 'codwdr') AND (!EMPTY(codwdr.kp) AND m.tiplpu!=codwdr.kp)
      m.recid = recid
      rval = InsError('S', 'TVA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ELSE 
     
     ENDIF
    ENDIF 
   ENDIF 

   IF M.TVA == .T.
    *IF INLIST(m.lpuid,2202,1873,1872,1871,1874) && Св. Владимира, Сперанского, Морозовская, Башляевой, Филатова (13)
    * c декабря 2019!
    IF INLIST(m.lpuid,1871,1872,1873,2202) && Св. Владимира, Сперанского, Морозовская, Башляевой, Филатова (13)
     m.profil  = profil
     IF (cod>=61010 AND cod<=99647) AND !INLIST(m.cod,70150,70160,70170,61400,61410) && AND m.profil!='034'
      m.recid = recid
      rval = InsError('S', 'TVA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF
    ENDIF 
   ENDIF 

   *IF M.FSA == .T. && Ошибка перенесена со счетов за Март 2019
   IF M.PFA == .T.
    m.p_cel = p_cel
    m.d_u   = d_u
    m.IsStomatUsl = IIF(INLIST(FLOOR(m.cod/1000), 9, 109), .T., .F.)
    IF !IsPilots AND !INLIST(m.mcod, '4344623','4344700','4344621','4344640','4344613','0343036','0244124') && С 01.02.2019 
     IF m.IsStomatUsl
      m.recid = recid
      *rval = InsError('S', 'FSA', m.recid)
      rval = InsError('S', 'PFA', m.recid, '',;
      	'Оказание стоматуслуг в МО, не являющимся "стомат пилотом" и не относящимся к МО из приложения 1.6.2.')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    *IF INLIST(m.mcod, '4344623','4344700','4344621','4344640','4344613','0343036','0244124') ;
    	AND EMPTY(lpu_ord)
    IF INLIST(m.mcod, '4344623','4344700','4344621','4344640','4344613','0343036','0244124') ;
    	AND (m.p_cel<>'1.1' AND m.d_u<{15.06.2020})
     IF m.IsStomatUsl
      m.recid = recid
      *rval = InsError('S', 'FSA', m.recid)
      *rval = InsError('S', 'PFA', m.recid, '',;
      	'Оказание стоматуслуг в МО из приложения 1.6.2 без направления')
      rval = InsError('S', 'PFA', m.recid, '',;
      	'Оказание нескоромощных (p_cel<>1.1)стоматуслуг в МО из приложения 1.6.2')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    IF m.IsPilots
     m.p_cel = p_cel
     IF IIF(SUBSTR(m.mcod,3,2)<>'07', m.IsStomatUsl, 1=1) AND ;
     	(!EMPTY(people.prmcods) AND people.prmcods<>people.mcod) AND !SEEK(m.cod, 'p_pr')
      m.otd = otd
      *IF m.IsStomatUsl AND !(IsStac(m.mcod) AND INLIST(SUBSTR(m.otd,2,2), '08','73')) && Закомментировано со счетов за декабрь!
       m.recid = recid
       rval = InsError('S', 'PFA', m.recid, '', ;
       	'Оказание стоматуслуг "чужим" пациентам в пилотах за исключением оказания стоматуслуг в 08 и 73 отделения стационара') 
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      *ENDIF 
     ENDIF 
     IF IIF(SUBSTR(m.mcod,3,2)<>'07', m.IsStomatUsl, 1=1) AND EMPTY(people.prmcods) AND m.p_cel<>'1.1' AND m.d_u<{15.06.2020}
      m.otd = otd
       m.recid = recid
       rval = InsError('S', 'PFA', m.recid, '', ;
       	'Оказание нескоромощных (p_cel<>1.1)стоматуслуг неприкрепленным пациентам') 
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.NLA == .T.
    m.tiplpu = IIF(VAL(SUBSTR(m.mcod,3,2))<41, 'p', IIF(VAL(SUBSTR(m.mcod,3,2))!=71, 's', 'b'))
    IF SEEK(cod, 'codwdr') AND !EMPTY(codwdr.stac) AND LOWER(codwdr.stac) != m.tiplpu
     DO CASE 
      CASE INLIST(m.cod, 29006, 129006, 29007, 129007) AND INLIST(SUBSTR(otd,2,2),'00','01','08','85','90','91','92','93')
      CASE INLIST(m.lpuid,1863,1891,1842,5009,5361) AND INLIST(m.cod, 29006, 29007)
      *CASE INLIST(m.lpuid,1912,1940,2049,1874,1909) AND INLIST(FLOOR(m.cod/1000), 146)
      CASE INLIST(m.lpuid,1874,1909) AND INLIST(FLOOR(m.cod/1000), 146) && Убрали с октября 2019
      OTHERWISE 
       m.recid = recid
       rval = InsError('S', 'NLA', m.recid, '', ;
       	'В МО типа '+m.tiplpu+' не допускается применение такой услуги/МЭС/койко-дня (допустимо для типа МО '+codwdr.stac)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDCASE 
    ENDIF

    * Неотложка в ЛПУ
    IF (INLIST(cod, 56031, 156002) AND m.tdat1>={01.10.2017}) AND LEFT(m.mcod,1)='0'
     m.recid = recid
     rval = InsError('S', 'NLA', m.recid, '', ;
     	'Услуги неотложной медпомощи 56031 и 156001 запрещены к оказанию в МО кроме СиНМП с 01.10.2017')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    * Неотложка в ЛПУ
    
    IF USED('sprnco') AND .f.
    IF INLIST(m.cod, 70150,70160,70170,170150,170151,170160,170161,170170,170171) AND ;
    	(!SEEK(m.lpuid, 'sprnco') OR sprnco.pnv<>1 OR ;
    		d_u<IIF(FIELD('DATEBEG_2','sprnco')='DATEBEG_2', sprnco.datebeg_2, sprnco.datebeg))
     m.recid = recid
     rval = InsError('S', 'NLA', m.recid, '', ;
     	'Нарушение применения оказания услуги 70150, 70160, 70170, 170150, 170151, 170160, 170161, 170170, 170171')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    IF INLIST(m.cod, 61400,161400,161401) AND (!SEEK(m.lpuid, "sprnco") OR sprnco.ncov<>1 OR ;
    	d_u < IIF(FIELD('DATEBEG_1','sprnco')='DATEBEG_1', sprnco.datebeg_1, sprnco.datebeg))
     m.recid = recid
     rval = InsError('S', 'NLA', m.recid, '', ;
     	'Нарушение применения оказания услуги 61400,161400,161401')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    ENDIF 
   ENDIF 

   IF M.MDA == .T.
    m.cod = cod
    m.d_type = d_type
    m.sex = IIF(OCCURS('#',c_i)==2, SUBSTR(c_i, AT('#',c_i,1)+1, 1), STR(people.w,1))
    IF !INLIST(SUBSTR(m.mcod,3,2),'27','67') AND m.d_type<>'4' ;
    	AND !INLIST(m.cod,61400,161400,161401,70150,70160,70170,170150,170151,170160,170161,170170,170171,61410,161410,;
    		161411,61420,161420,161421,61430,161430,161431)
     IF (SEEK(cod, 'codwdr') AND !EMPTY(codwdr.sex)) AND m.sex != codwdr.sex
      m.recid = recid
      rval = InsError('S', 'MDA', m.recid, '', ;
     	'Несоответствие пола пациента ('+m.sex+') разрешенному полу по справочнику codwdr ('+codwdr.sex+')')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF

     m.otd = otd
     IF m.sex='1' AND INLIST(SUBSTR(m.otd,4,3),'003','136','137','184')
      m.recid = recid
      rval = InsError('S', 'MDA', m.recid, '', ;
     	'Оказание услуг мужчине в отделении гинекологического профиля (4,5,6 позиция фасетного кода отд 003,136,137)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.H3A == .T.
    IF !SEEK(d_type, 'ososch')
     m.recid = recid
     rval = InsError('S', 'H3A', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF
   ENDIF 
   
   IF M.SOA == .T.
    * надо проверять по внутреннему справочнику отделений
   ENDIF 

   IF M.R1A == .T.
    m.ishod = ishod
    m.rslt  = rslt
    IF EMPTY(m.ishod)
     m.recid = recid
     rval = InsError('S', 'R1A', m.recid, '', ;
     	'Незаполнено поле ishod (исход заболевания)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
     IF !SEEK(m.ishod, 'isv012')
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid, '', ;
     	'Исход лечения (поле ishod) не соответствует кодификатору ISV012')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     
     IF USED('rsltishod')
      m.vir = STR(m.rslt,3)+STR(m.ishod,3)
      IF !SEEK(m.vir, 'rsltishod')
       m.recid = recid
       rval = InsError('S', 'R1A', m.recid, '', ;
     	'Сочетание rslt+ishod не соответствует справочнику rsltishod')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 

     m.c_i = ALLTRIM(c_i)
     m.cod = cod

     *IF (BETWEEN(m.cod,1925,1935) AND LEFT(m.c_i,3)='ДД_') AND m.ishod!=304
     IF ((BETWEEN(cod,1936,1948) OR BETWEEN(cod,1955,1967)) AND LEFT(m.c_i,3)='ДД_') AND m.ishod!=304
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid, '', ;
     	'Исход лечения (поле ishod) не равен 304 для диспансеризация взрослых (услуги 1910-1920,25204,35401,1017,1807 и карта "ДД_"')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     *IF ((BETWEEN(m.cod,1910,1920) OR INLIST(m.cod, 25204,35401,1017,1807 )) AND LEFT(m.c_i,3)='ДД_') AND m.ishod!=304
     * m.recid = recid
     * rval = InsError('S', 'R1A', m.recid, '', ;
     *	'Исход лечения (поле ishod) не равен 304 для диспансеризация взрослых (услуги 1910-1920,25204,35401,1017,1807 и карта "ДД_"')
     * m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     *ENDIF 

     IF (INLIST(m.cod,101952,101028,101030) AND LEFT(m.c_i,3)='ДУ_') AND m.ishod!=304
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid, '', ;
     	'Исход лечения (поле ishod) не равен 304 для диспансеризация детей (услуги 101952,101028,101030 и карта "ДУ_"')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     *IF (BETWEEN(m.cod,1906,1909) AND LEFT(m.c_i,6)='ПРОФВ_') AND m.ishod!=304
     * m.recid = recid
     * rval = InsError('S', 'R1A', m.recid, '', ;
     *	'Исход лечения (поле ishod) не равен 304 для проф. осмотры взрослых (услуги 1906-1909 и карта "ПРОФВ_"')
     * m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     *ENDIF 
     *IF (BETWEEN(m.cod,1921,1924) AND LEFT(m.c_i,6)='ПРОФВ_') AND m.ishod!=304
     IF ((BETWEEN(cod,1949,1954) OR BETWEEN(cod,1968,1973)) AND LEFT(m.c_i,6)='ПРОФВ_') AND m.ishod!=304
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid, '', ;
     	'Исход лечения (поле ishod) не равен 304 для проф. осмотры взрослых (услуги 1949-1954/1968-1973 и карта "ПРОФВ_"')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     IF ((BETWEEN(m.cod,101933,101945) OR INLIST(m.cod, 101951,101028,101030)) AND LEFT(m.c_i,3)='ПРОФД_') AND m.ishod!=304
      m.recid = recid
      rval = InsError('S', 'R1A', m.recid, '', ;
     	'Исход лечения (поле ishod) не равен 304 для проф. осмотров детей (услуги 101933-101945,101951,101028,101030 и карта "ПРОФД_"')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 

     && Проверка 14.01.2019 соответствия условиям оказания
     *m.otd = otd
     m.o_otd   = SUBSTR(otd,2,2)
     m.usl_ok  = IIF(SEEK(m.o_otd, 'profot'), profot.usl_ok, '0')
     DO CASE 
      CASE m.usl_ok='1' AND LEFT(STR(m.ishod,3),1) != '1'
      *CASE !(IsDstOtd(m.otd) OR IsPlkOtd(m.otd)) AND LEFT(STR(m.ishod,3),1) != '1'
      *CASE IsGsp(cod) AND LEFT(STR(ishod,3),1) != '1'
       m.recid = recid
       rval = InsError('S', 'R1A', m.recid, '', ;
       	'Значения поля ishod (исход лечения) в условиях круглосуточного стационара не соответствует маске 1ХХ')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.usl_ok='2' AND LEFT(STR(m.ishod,3),1) != '2'
      *CASE IsDstOtd(m.otd) AND LEFT(STR(ishod,3),1) != '2'
      *CASE IsDst(cod) AND LEFT(STR(ishod,3),1) != '2'
       m.recid = recid
       rval = InsError('S', 'R1A', m.recid, '', ;
       	'Значения поля ishod (исход лечения) в условиях дневного стационара  не соответствует маске 2ХХ')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.usl_ok='3' AND LEFT(STR(m.ishod,3),1) != '3'
      *CASE IsPlkOtd(m.otd) AND LEFT(STR(ishod,3),1) != '3'
      *CASE IsPlk(cod) AND LEFT(STR(ishod,3),1) != '3'
       m.recid = recid
       rval = InsError('S', 'R1A', m.recid, '', ;
       	'Значения поля ishod (исход лечения) в амбулаторных условиях не соответствует маске 3ХХ')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      OTHERWISE 
     ENDCASE 
     && Проверка 14.01.2019 соответствия условиям оказания
    ENDIF 

   ENDIF 

   IF M.R2A == .T.
    m.rslt  = rslt
    m.ishod = ishod
    m.tip   = tip
    m.d_type = d_type
    
    IF EMPTY(m.rslt)
     m.recid = recid
     rval = InsError('S', 'R2A', m.recid, '', ;
     	'Незаполнено поле rslt (результат лечения)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
     IF !SEEK(rslt, 'rsv009')
      m.recid = recid
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt (результат лечения) не соответствует кодификатору RSV009')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     DO CASE 
      CASE INLIST(m.rslt, 102,103,104,105,106,107,108,109) AND m.ishod=101
       m.recid = recid
       rval = InsError('S', 'R2A', m.recid, '', 'RSLT = {102, 103, 104, 105, 106, 107, 108, 109} и ISHOD = 101')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      CASE INLIST(m.rslt, 202,203,204,205,206,207,208) AND m.ishod=201
       m.recid = recid
       rval = InsError('S', 'R2A', m.recid, '', 'RSLT = {202, 203, 204, 205, 206, 207, 208} и ISHOD = 201')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      CASE INLIST(m.rslt,105,106) AND m.ishod<>104
       m.recid = recid
       rval = InsError('S', 'R2A', m.recid, '', 'RSLT = {105, 106} и ISHOD = 104')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      CASE INLIST(m.rslt,205,206) AND m.ishod<>204
       m.recid = recid
       rval = InsError('S', 'R2A', m.recid, '', 'RSLT = {205, 206} и ISHOD = 204')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      CASE m.rslt=313 AND m.ishod<>305
       m.recid = recid
       rval = InsError('S', 'R2A', m.recid, '', 'RSLT = {313} и ISHOD = 305')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      OTHERWISE 
     ENDCASE 
    ENDIF 
    
    IF INLIST(m.rslt,105,106,205,206,313) AND !(m.tip='5' OR m.d_type='5')
      m.recid = recid
      rval = InsError('S', 'R2A', m.recid, '', 'RSLT = {105,106,205,206,313} и TIP<>5 или d_type<>5')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    m.c_i = c_i
    m.cod = cod
    DO CASE 
	 *ДД 
     CASE (LEFT(m.c_i,3)='ДД_' AND INLIST(m.cod,1017,1807)) AND !(BETWEEN(m.rslt,355,356) OR INLIST(m.rslt,317,318))
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для ДД взрослого населения (услуги 1017,1807, карта ДД_ ) не равно 355-356,317,318)')

     *CASE (LEFT(m.c_i,3)='ДД_' AND !(INLIST(m.cod,1017,1807) OR BETWEEN(m.cod,1925,1935))) AND ;
     *	!INLIST(m.rslt,353,357,358)
     CASE (LEFT(m.c_i,3)='ДД_' AND !(INLIST(m.cod,1017,1807) OR (BETWEEN(cod,1936,1948) OR BETWEEN(cod,1955,1967)))) AND ;
     	!INLIST(m.rslt,353,357,358)
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для ДД взрослого населения (услуги НЕ 1017,1807,1936-1948,1955-1967 карта ДД_ ) не равно 353,357,358)')

     *CASE (LEFT(m.c_i,3)='ДД_' AND BETWEEN(m.cod,1925,1935)) AND ;
     *	!(INLIST(m.rslt,317,318,353,357,358) OR BETWEEN(m.rslt,355,356))
     * tip=1
     CASE (LEFT(m.c_i,3)='ДД_' AND (BETWEEN(cod,1936,1948) OR BETWEEN(cod,1955,1967))) AND ;
     	!(INLIST(m.rslt,317,318,353,357,358) OR BETWEEN(m.rslt,355,356))
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для ДД взрослого населения (услуги 1936-1948,1955-1967, карта ДД_ ) не равно 317,318,353,355-356,357,358,)')
	 *ДД

     * ПМО взрослого населения
     *CASE (LEFT(m.c_i,6)='ПРОФВ_' AND BETWEEN(m.cod,1921,1924)) AND ;
     *	!(BETWEEN(m.rslt,343,344) OR INLIST(m.rslt,373,374))
     CASE (LEFT(m.c_i,6)='ПРОФВ_' AND (BETWEEN(cod,1949,1954) OR BETWEEN(cod,1968,1973))) AND ;
     	!(BETWEEN(m.rslt,343,344) OR INLIST(m.rslt,373,374))
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для проф. медосмотров взрослого населения (услуги 1949-1954/1968-1973) не равно 343-344, 373,374')
     * ПМО взрослого населения
     
     * ПМО несовершеннолетних 
     CASE (LEFT(m.c_i,6)='ПРОФД_' AND (BETWEEN(m.cod,101933,101945) OR m.cod=101951)) AND ;
     	!(BETWEEN(m.rslt,332,336) OR BETWEEN(m.rslt,361,364))
     *CASE (LEFT(m.c_i,6)='ПРОФД_' AND (BETWEEN(m.cod,101933,101945) OR m.cod=101951)) AND ;
     	!(BETWEEN(m.rslt,332,336))
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для проф. медосмотров несовершеннолетних (услуги 101933-101945,101951) не равно 332-336')
     	
     CASE (LEFT(m.c_i,6)='ПРОФД_' AND INLIST(m.cod,101028, 101030)) AND ;
     	!BETWEEN(m.rslt,332,336) && этих кодов нет в dspcodes 101028, 101030!
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для проф. медосмотров несовершеннолетних (услуги 101028, 101030, карта ПРОФД_) не равно 332-336')
     * ПМО несовершеннолетних 

     * ДУ
     && добавлено со счетов за декабрь 101003 и BETWEEN(m.rslt,369,372)
     CASE (LEFT(m.c_i,3)='ДУ_' AND INLIST(m.cod,101028,101030,101003)) AND !(BETWEEN(m.rslt,347,351) OR BETWEEN(m.rslt,369,372))
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для проф. медосмотров несовершеннолетних (услуги 101028, 101030, карта ДУ_ ) не равно 347-351')

     CASE (LEFT(m.c_i,3)='ДУ_' AND !INLIST(m.cod,101028,101030,101003)) AND BETWEEN(m.rslt,369,372)
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для ДУ равно 369,372 для услуг НЕ 101028,101030,101003 и карты ДУ_')

     * tip =3
     CASE (LEFT(m.c_i,3)='ДУ_' AND INLIST(m.cod,101952)) AND !(BETWEEN(m.rslt,347,351) OR BETWEEN(m.rslt,369,372))
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для ПМО несовершеннолетних (услуга 101952, карта ДУ_ ) не равно 347-351,369-372')

     *CASE (LEFT(m.c_i,3)='ДУ_' AND INLIST(m.cod,101028,101030)) AND !BETWEEN(m.rslt,347,351)
     * rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для проф. медосмотров несовершеннолетних (услуги 101028,101030, карта ДУ_ ) не равно 347-351')

     *CASE (LEFT(m.c_i,3)='ДУ_' AND !INLIST(m.cod,101952,101028,101030)) AND !BETWEEN(m.rslt,369,372)
     * rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для проф. медосмотров несовершеннолетних (услуги НЕ 101952,101028,101030 и карта ДУ_ ) не равно 369-372')
     * ДУ
     
     *ДС
     CASE (LEFT(m.c_i,3)='ДС_' AND INLIST(m.cod,101952)) AND !(BETWEEN(m.rslt,321,325) OR BETWEEN(m.rslt,365,368))
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для ДС (услуга 101952, карта ДC_ ) не равно 321-325, 365-368')

     && добавлено со счетов за декабрь 101003 и BETWEEN(m.rslt,365,368)
     CASE (LEFT(m.c_i,3)='ДС_' AND INLIST(m.cod,101028,101030,101003)) AND !(BETWEEN(m.rslt,321,325) OR BETWEEN(m.rslt,365,368))
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для проф. медосмотров несовершеннолетних (услуги 101028,101030, карта ДC_ ) не равно 321-325')

     CASE (LEFT(m.c_i,3)='ДС_' AND !INLIST(m.cod,101028,101030,101003)) AND BETWEEN(m.rslt,365,368)
      rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для ДС (услуги НЕ 101028,101030,101003 карта ДC_) равно 365,368')

     *CASE (LEFT(m.c_i,3)='ДС_' AND !INLIST(m.cod,101952,101028,101030)) AND !(BETWEEN(m.rslt,365,368))
     * rval = InsError('S', 'R2A', m.recid, '', ;
     	'Значение поля rslt для проф. медосмотров несовершеннолетних (услуги НЕ 101952,101028,101030, карта ДC_ ) не равно 365-368')
     *ДС
	 


    OTHERWISE 

    ENDCASE 
    
     m.otd = otd
     m.o_otd   = SUBSTR(otd,2,2)
     m.usl_ok  = IIF(SEEK(m.o_otd, 'profot'), profot.usl_ok, '0')
     DO CASE 
      CASE m.usl_ok='1' AND LEFT(STR(m.rslt,3),1) != '1'
      *CASE !(IsDstOtd(m.otd) OR IsPlkOtd(m.otd)) AND LEFT(STR(m.rslt,3),1) != '1'
       m.recid = recid
       rval = InsError('S', 'R2A', m.recid, '', ;
       	'Значение поля rslt (результат лечения) в условиях круглосуточного стационара не соответствует маске 1ХХ')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.usl_ok='2' AND LEFT(STR(m.rslt,3),1) != '2'
      *CASE IsDstOtd(m.otd) AND LEFT(STR(m.rslt,3),1) != '2'
       m.recid = recid
       rval = InsError('S', 'R2A', m.recid, '', ;
       	'Значение поля rslt (результат лечения) в условиях дневного стационара  не соответствует маске 2ХХ')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.usl_ok='3' AND LEFT(STR(m.rslt,3),1) != '3'
      *CASE IsPlkOtd(m.otd) AND LEFT(STR(m.rslt,3),1) != '3'
       m.recid = recid
       rval = InsError('S', 'R2A', m.recid, '', ;
       	'Значение поля rslt (результат лечения) в амбулаторных условиях не соответствует маске 3ХХ')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      OTHERWISE 
     ENDCASE 
    
   ENDIF 

   IF M.R3A == .T.
    m.prvs = prvs
    m.otd  = otd
    m.cod  = cod
    
    IF EMPTY(m.prvs)
     m.recid = recid
     rval = InsError('S', 'R3A', m.recid, '',;
     	'Пустое поле кода специальности (prvs) (кодификатор специальностей НСИ "SpV014XX", поле code)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ELSE  
     IF !SEEK(m.prvs, 'kspec')
      m.recid = recid
      rval = InsError('S', 'R3A', m.recid, '',;
     	'Регистрация кода специальности (prvs), не включенной в кодификатор специальностей НСИ "SpV014XX"')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ELSE 
	  m.usl_ok = IIF(SEEK(SUBSTR(otd,2,2), 'profot'), profot.usl_ok, '0')
	  
	  IF m.usl_ok='1' AND !INLIST(INT(m.cod/1000),200,297,300,397)
	   m.v_mp = IIF(m.usl_ok='1', 31, 0)
	  ELSE 
	   m.v_mp = IIF(SEEK(m.cod, 'usvmp'), usvmp.vmp, 0)
	  ENDIF 
	  
	  IF m.v_mp>0
	   m.prvs = prvs
	   DO CASE 
	    CASE INLIST(m.prvs,4,5,204,205,206,208,209,210,211,212,213,214,216,217,218,219,221) AND INLIST(m.v_mp,12,13,31)
	     IF !(SUBSTR(m.otd,2,2)='73' AND (IsUsl(m.cod) AND ((m.cod>=2000 AND m.cod<=101000) OR m.cod>102000)))
	      m.recid = recid
	      rval = InsError('S', 'R3A', m.recid, '',;
	   	 	'Недопустимый код специальности врача '+STR(m.prvs,4)+' для вида медицинской помощи '+STR(m.v_mp,3))
	      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	     ENDIF 
	    CASE INLIST(m.prvs,223,224,226,227,228,230,231,232,233,235,279,281,282,287,288,3200) AND INLIST(m.v_mp,12,13,31)
	     IF !(SUBSTR(m.otd,2,2)='73' AND (IsUsl(m.cod) AND ((m.cod>=2000 AND m.cod<=101000) OR m.cod>102000)))
	      m.recid = recid
	      rval = InsError('S', 'R3A', m.recid, '',;
	   	 	'Недопустимый код специальности врача '+STR(m.prvs,4)+' для вида медицинской помощи '+STR(m.v_mp,3))
	      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	     ENDIF 
	    OTHERWISE 
	   ENDCASE 
	  ENDIF 
	  
	  IF m.usl_ok='1' AND INLIST(SUBSTR(otd,4,3),'005','167')
	   IF m.prvs<>9
	    m.recid = recid
	    rval = InsError('S', 'R3A', m.recid, '',;
	   	 	'Недопустимый код специальности врача (д.б. 9) '+STR(m.prvs,4)+' в отделении реанимации и интенсивной терапии')
	    m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	   ENDIF 
	  ENDIF 
	  
	 ENDIF 
    ENDIF 
   ENDIF 

   
   IF M.NRA == .T.  && Версия до 15.01.2019!
    m.cod    = cod
    m.otd    = otd
    m.profil = profil
    m.lpu_ord = lpu_ord
    m.ord     = ord
    m.IsTpnR  = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod)), .T., .F.)
    m.Is02    = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='q', .t., .f.)
    m.IsR     = IIF(SEEK(m.cod, 'tarif') and tarif.tpn='r', .t., .f.)
    m.d_type  = d_type
    m.c_i     = c_i
    
    m.sn_pol  = sn_pol
    m.facotd  = SUBSTR(otd,2,2)
    m.profil  = profil
    m.lpu_ord = lpu_ord
    m.otd     = otd
    m.ds      = ds
   
    DO CASE 
     CASE IsPlkOtd(m.otd) AND !INLIST(m.cod,56029,156002)
      m.recid = recid
      m.ord   = ord
      m.lpu_ord = lpu_ord
      *IF !INLIST(m.ord,0,4,6,7,8,9)
      IF !INLIST(m.ord,0,4,6,7,8) && С 202001
       rval = InsError('S', 'NRA', m.recid, '', ;
       	'ord='+STR(m.ord,1)+' в АПО-отделении/кабинете (фасетный код='+m.facotd+') при допустимых ord=0,4,6,7,8')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 

     CASE IsDstOtd(m.otd)
      m.recid = recid
      m.ord   = ord
      m.lpu_ord = lpu_ord
      IF !INLIST(m.ord,0,1,5) && Версия от 28.04.2019
       rval = InsError('S', 'NRA', m.recid, '',;
       	'ord='+STR(m.ord,1)+' в отделении дневного стационара (фасетный код='+m.facotd+') при допустимых ord=0,1,5')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 

     OTHERWISE 
      m.recid = recid
      m.ord   = ord
      m.lpu_ord = lpu_ord
      *IF !INLIST(m.ord,1,2,3,5,6)
      IF !IsVmp(m.cod) AND !INLIST(m.ord,0,1,2,3,5,6) && Версия от 28.04.2019
       rval = InsError('S', 'NRA', m.recid, '',;
       	'ord='+STR(m.ord,1)+' при применении МЭС в отделении стационара (фасетный код='+m.facotd+') при допустимых ord=0,1,2,3,5,6')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      IF IsVmp(m.cod) AND !INLIST(m.ord,1,2,3,5) && Версия от 28.04.2019
       rval = InsError('S', 'NRA', m.recid, '',;
       	'ord='+STR(m.ord,1)+' при применении ВМП в отделении стационара (фасетный код='+m.facotd+') при допустимых ord=1,2,3,5')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
      *IF INLIST(m.cod,200408,200491) AND INLIST(m.ord,1,5) && Версия от 28.04.2019
      * rval = InsError('S', 'NRA', m.recid, '', ;
      * 	'ord='+STR(m.ord,1)+' при применении кодов с 200408 по 200491 в отделении стационара (фасетный код='+m.facotd+') при допустимых ord=1,5)')
      * m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      *ENDIF 
      *IF IsVmp(m.cod) AND INLIST(m.ord,2,3) AND LEFT(m.ds,1)!='I' && Версия от 28.04.2019
      IF IsVmp(m.cod) AND INLIST(m.ord,3) AND LEFT(m.ds,1)!='I' && Версия от 28.04.2019
       rval = InsError('S', 'NRA', m.recid, '',;
       	'ord='+STR(m.ord,1)+' при применении ВМП в отделении стационара (фасетный код='+m.facotd+') при диагнозе!=Ixxx.xx ('+m.ds+')')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 

      IF USED('sprved') && с 202001
       IF SEEK(m.mcod, 'sprved')
        DO CASE 
         CASE !IsUsl(m.cod) AND INLIST(m.ord,0,1,2,5)
         CASE ((IsUsl(m.cod) AND people.tip_p=3) OR INLIST(m.cod,56029,156002)) AND INLIST(m.ord,0,1,5,2)
         OTHERWISE 
          rval = InsError('S', 'NRA', m.recid, '',;
       	 	'Недопустимое ord='+STR(m.ord,1)+' в ведомственном МО (МО из справочника sprved)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDCASE 
       ENDIF 
      ENDIF 

    ENDCASE 
    
    IF ((INLIST(m.cod, 70150,70160,70170,170150,170151,170160,170161,170170,170171,61400,161400,161401) AND m.ord<>2) ;
    	AND OCCURS('#', m.c_i)<>3) AND (INLIST(m.cod,61400,161400,161401,61430,161430,161431) AND m.ord<>1)
     IF !(INLIST(m.lpuid,4511,2293,4586) AND m.ord=3)
      rval = InsError('S', 'NRA', m.recid, '',;
     	'ord<>2 при МЭС=70150,70160,70170,170150,170151,170160,170161,170170,170171,61400,161400,161401')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 

   ENDIF && IF M.NRA == .T. 

   IF M.G1A == .T.
    m.recid   = recid
    m.cod     = cod 
    m.c_i     = c_i
    m.sn_pol  = sn_pol
    m.ord     = ord
    m.lpu_ord = lpu_ord
    m.lpu_ord = IIF(m.lpu_ord>9999 AND FLOOR(m.lpu_ord/10000)=77, m.lpu_ord%10000, m.lpu_ord)
    m.ordmcod = IIF(SEEK(m.lpu_ord, 'sprlpu'), sprlpu.mcod, '')
    m.facotd  = SUBSTR(otd,2,2)
    m.profil  = profil
    m.date_ord = date_ord
    m.d_type   = d_type
    m.ds       = ds
    
    m.d_pos = IIF(USED('Gosp_d') AND SEEK(m.c_i, 'Gosp_d'), Gosp_d.d_pos, {})
    m.d_pos = IIF(EMPTY(m.d_pos), IIF(SEEK(m.sn_pol, 'people'), people.d_beg, {}), m.d_pos)
    m.tipp  = IIF(SEEK(m.sn_pol, 'people'), people.tipp, '')
    
    m.pr_gosp = IIF(USED('Gosp_d') AND SEEK(m.c_i, 'Gosp_d'), SUBSTR(Gosp_d.otd,4,3), SUBSTR(m.otd,4,3))
    m.pr_diag = IIF(USED('Gosp_d') AND SEEK(m.c_i, 'Gosp_d'), Gosp_d.ds, m.ds)
    
    IF !INLIST(m.pr_gosp,'018','060','017','029','021','122')
     IF USED('Gosp_d') AND SEEK(m.c_i, 'Gosp_d')
      m.vir       = STR(Gosp_d.cod,6) + Gosp_d.ds
      m.vir_empty = STR(Gosp_d.cod,6) + SPACE(6)
      IF SEEK(m.vir, 'ms_ds_prv')
       m.pr_gosp = ms_ds_prv.prv
      ELSE 
       IF SEEK(m.vir_empty, 'ms_ds_prv')
        m.pr_gosp = ms_ds_prv.prv
       ENDIF 
      ENDIF  
     ELSE 
      m.vir       = STR(m.cod,6) + m.ds
      m.vir_empty = STR(m.cod,6) + SPACE(6)
      IF SEEK(m.vir, 'ms_ds_prv')
       m.pr_gosp = ms_ds_prv.prv
      ELSE 
       IF SEEK(m.vir_empty, 'ms_ds_prv')
        m.pr_gosp = ms_ds_prv.prv
       ENDIF 
      ENDIF  
     ENDIF 
     *m.pr_gosp = IIF(!EMPTY(m.pr_gosp), m.pr_gosp, SUBSTR(m.otd,4,3))
    ENDIF 

    IF !INLIST(m.pr_gosp,'018','060','017','029','021','122')
     m.pcod = pcod
     m.prvs = IIF(SEEK(m.pcod, 'doctor'), doctor.prvs, 0)
     m.pr_gosp = PADL(m.prvs,3,'0')
    ENDIF 
    
    m.k_ved = IIF(SEEK(m.lpu_ord, 'sprlpu'), sprlpu.prn_kodved, 0)
    
    IF m.ord>0 AND !EMPTY(m.date_ord) AND SEEK(m.lpu_ord, 'f003') AND !EMPTY(f003.d_end)
     IF m.date_ord>f003.d_end
      rval = InsError('S', 'G5A', m.recid, '', ;
      	'Дата направления date_ord ('+DTOC(m.date_ord)+') выдано МО '+ALLTRIM(f003.s_name)+;
      		' после выхода его из системы ОМС '+DTOC(f003.d_end))
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
     IF m.date_ord<f003.d_beg
      rval = InsError('S', 'G5A', m.recid, '', ;
      	'Дата направления date_ord ('+DTOC(m.date_ord)+') выдано МО '+ALLTRIM(f003.s_name)+;
      		' до начала его работы в системе ОМС '+DTOC(f003.d_beg))
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
    IF !(IsPlkOtd(otd) OR IsDstOtd(otd)) && только стационар
     IF SEEK(m.mcod, 'sprved')
      IF m.ord>0 AND m.lpu_ord>0
       IF !SEEK(m.lpu_ord, 'sprlpu')
        IF m.lpu_ord<=9999
         rval = InsError('S', 'G5A', m.recid, '', ;
      	 	'Направления на госпитализацюи в ведомственный стационар (sprved) выдано неизвестным МО: '+STR(m.lpu_ord))
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       ELSE 

      * Здесь проверять МО прикрепления
      IF !EMPTY(m.date_ord)
       m.lpu_pr = 0
       m.d_att  = {}
	   DO CASE 
	    CASE m.tipp='В'
	     m.polis = ALLTRIM(sn_pol)
	     IF LEN(m.polis)=9
	      m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.lpu_tera, 0)
	      m.lpu_st   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.lpu_stom, 0)
	      m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.date_tera, {})
	      m.d_st     = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.date_stom, {})
	     ENDIF 

	    CASE INLIST(m.tipp,'П','Э','К')
	     m.polis   = LEFT(sn_pol,16)
	     m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.lpu_tera, 0)
	     m.lpu_st   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.lpu_stom, 0)
 	     m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.date_tera, {})
 	     m.d_st     = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.date_stom, {})

	    CASE m.tipp='С'
	     m.polis = ALLTRIM(sn_pol)
	     m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.lpu_tera, 0)
	     m.lpu_st   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.lpu_stom, 0)
	     m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.date_tera, {})
	     m.d_st     = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.date_stom, {})

	    OTHERWISE 
	   ENDCASE 
	    
	   IF (EMPTY(m.d_att) OR m.d_att<=m.date_ord) AND (EMPTY(m.d_st) OR m.d_st<=m.date_ord)
        
	   
	   ELSE &&  взять предыдущий номерник 

        m.lpu_pr = 0
        m.d_att  = {}
	    DO CASE 
	     CASE m.tipp='В'
	      m.polis = ALLTRIM(sn_pol)
	      IF LEN(m.polis)=9
	       m.lpu_pr   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_tera, 0)
	       m.lpu_st   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_stom, 0)
	       m.d_att    = IIF(SEEK(m.polis, 'vsn'), vsn.date_tera, {})
	       m.d_st     = IIF(SEEK(m.polis, 'vsn'), vsn.date_stom, {})
	      ENDIF 

	     CASE INLIST(m.tipp,'П','Э','К')
	      m.polis   = LEFT(sn_pol,16)
	      m.lpu_pr   = IIF(SEEK(m.polis, 'enp'), enp.lpu_tera, 0)
	      m.lpu_st   = IIF(SEEK(m.polis, 'enp'), enp.lpu_stom, 0)
 	      m.d_att    = IIF(SEEK(m.polis, 'enp'), vsn.date_tera, {})
 	      m.d_st     = IIF(SEEK(m.polis, 'enp'), vsn.date_stom, {})

	     CASE m.tipp='С'
	      m.polis = ALLTRIM(sn_pol)
	      m.lpu_pr   = IIF(SEEK(m.polis, 'kms'), kms.lpu_tera, 0)
	      m.lpu_st   = IIF(SEEK(m.polis, 'kms'), kms.lpu_stom, 0)
	      m.d_att    = IIF(SEEK(m.polis, 'kms'), vsn.date_tera, {})
	      m.d_st     = IIF(SEEK(m.polis, 'kms'), vsn.date_stom, {})
         OTHERWISE 
	    ENDCASE 

	    IF (!EMPTY(m.d_att) AND m.d_att<=m.date_ord) AND (!EMPTY(m.d_st) OR m.d_st<=m.date_ord)
	    ELSE 
	     * Здесь открыть номерник предыдущего месяца
	    ENDIF 
	   ENDIF 
	  ELSE  
       m.lpu_pr = 0
       m.d_att  = {}
      ENDIF && IF !EMPTY(m.date_ord)
      * Здесь проверять МО прикрепления

        *m.pr_m  = IIF(SEEK(m.sn_pol, 'people'), people.prmcod, '')
        *m.pr_id = IIF(!EMPTY(m.pr_m) AND SEEK(m.pr_m, 'sprlpu', 'mcod'), sprlpu.prn_kodved, 0)
        m.pr_id = IIF(m.lpu_pr>0 AND SEEK(m.lpu_pr, 'sprlpu', 'lpu_id'), sprlpu.prn_kodved, 0)
        *m.prmcod_id = IIF(!EMPTY(m.pr_m) AND SEEK(m.pr_m, 'sprlpu', 'mcod'), sprlpu.lpu_id, 0)

        DO CASE 
         *CASE m.k_ved = 0 && k_ved - ведомство по lpu_ord
         
         *CASE sprlpu.prn_kodved=28 && ДЗМ, ок!

         CASE m.k_ved=28 && направил ДЗМ, все ок!

         CASE sprved.prn_kodved=m.k_ved && свое ведомство, ок, проверяем прикрепление
          IF m.pr_id>0 AND m.pr_id<>m.k_ved
           rval = InsError('S', 'G5A', m.recid, '', ;
      	 	'Направления на госпитализацюи в ведомственный ('+STR(sprved.prn_kodved,3)+;
      	 		') стационар выдало МО ведомства '+STR(m.k_ved)+' пациенту, прикрепленному к иному ведомству:'+STR(m.pr_id))
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ELSE 

           *IF m.pr_id>0 AND m.prmcod_id<>m.lpu_ord
           * rval = InsError('S', 'G5A', m.recid, '', ;
      	   *	'Направления на госпитализацюи в ведомственный ('+STR(sprved.prn_kodved,3)+;
      	   *		') стационар выдало МО ведомства '+STR(m.lpu_ord)+' пациенту, прикрепленному к иному МО того же ведомства:'+STR(m.prmcod_id))
           * m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
           *ENDIF 

          * Здесь проверять МО прикрепления
          IF !EMPTY(m.date_ord)
           m.lpu_pr = 0
           m.d_att  = {}
		   DO CASE 
		    CASE m.tipp='В'
		     m.polis = ALLTRIM(sn_pol)
		     IF LEN(m.polis)=9
		      m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.lpu_tera, 0)
		      m.lpu_st   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.lpu_stom, 0)
		      m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.date_tera, {})
		      m.d_st     = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.date_stom, {})
		     ENDIF 

		    CASE INLIST(m.tipp,'П','Э','К')
		     m.polis   = LEFT(sn_pol,16)
		     m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.lpu_tera, 0)
		     m.lpu_st   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.lpu_stom, 0)
 	         m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.date_tera, {})
 	         m.d_st     = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.date_stom, {})

		    CASE m.tipp='С'
		     m.polis = ALLTRIM(sn_pol)
		     m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.lpu_tera, 0)
		     m.lpu_st   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.lpu_stom, 0)
		     m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.date_tera, {})
		     m.d_st     = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.date_stom, {})

		    OTHERWISE 
		   ENDCASE 
		    
		   IF (EMPTY(m.d_att) OR m.d_att<=m.date_ord) AND (EMPTY(m.d_st) OR m.d_st<=m.date_ord)
            *IF m.qcod<>'I3'
		    IF (!EMPTY(m.d_att) AND m.lpu_pr<>m.lpu_ord) AND (!EMPTY(m.d_st) AND m.lpu_st<>m.lpu_ord) && если не прикреплен ни к кому, то не бракуем!
             rval = InsError('S', 'G5A', m.recid, '', ;
       	      'МО направления ('+STR(m.lpu_ord,4)+') не соответствует МО прикрепления ('+STR(m.lpu_pr,4)+;
       	      ') на дату выдачи направления '+DTOC(m.date_ord)+', d_att='+DTOC(m.d_att))
		    ENDIF
		    *ENDIF 
		   
		   ELSE &&  взять предыдущий номерник 

            m.lpu_pr = 0
            m.d_att  = {}
		    DO CASE 
		     CASE m.tipp='В'
		      m.polis = ALLTRIM(sn_pol)
		      IF LEN(m.polis)=9
		       m.lpu_pr   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_tera, 0)
		       m.lpu_st   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_stom, 0)
		       m.d_att    = IIF(SEEK(m.polis, 'vsn'), vsn.date_tera, {})
		       m.d_st     = IIF(SEEK(m.polis, 'vsn'), vsn.date_stom, {})
		      ENDIF 

		     CASE INLIST(m.tipp,'П','Э','К')
		      m.polis   = LEFT(sn_pol,16)
		      m.lpu_pr   = IIF(SEEK(m.polis, 'enp'), enp.lpu_tera, 0)
		      m.lpu_st   = IIF(SEEK(m.polis, 'enp'), enp.lpu_stom, 0)
 	          m.d_att    = IIF(SEEK(m.polis, 'enp'), vsn.date_tera, {})
 	          m.d_st     = IIF(SEEK(m.polis, 'enp'), vsn.date_stom, {})

		     CASE m.tipp='С'
		      m.polis = ALLTRIM(sn_pol)
		      m.lpu_pr   = IIF(SEEK(m.polis, 'kms'), kms.lpu_tera, 0)
		      m.lpu_st   = IIF(SEEK(m.polis, 'kms'), kms.lpu_stom, 0)
		      m.d_att    = IIF(SEEK(m.polis, 'kms'), vsn.date_tera, {})
		      m.d_st     = IIF(SEEK(m.polis, 'kms'), vsn.date_stom, {})
	         OTHERWISE 
		    ENDCASE 

		    IF (!EMPTY(m.d_att) AND m.d_att<=m.date_ord) AND (!EMPTY(m.d_st) OR m.d_st<=m.date_ord)
		     IF m.lpu_pr<>m.lpu_ord
		      *IF m.qcod<>'I3'
              rval = InsError('S', 'G5A', m.recid, '', ;
       	       'МО направления ('+STR(m.lpu_ord,4)+') не соответствует МО прикрепления ('+STR(m.lpu_pr,4)+;
       	       ') на дату выдачи направления '+DTOC(m.date_ord)+', d_att='+DTOC(m.d_att))
       	      *ENDIF 
		     ENDIF
		    ELSE 
		     * Здесь открыть номерник предыдущего месяца
		    ENDIF 
		   ENDIF 
          ENDIF && IF !EMPTY(m.date_ord)
          * Здесь проверять МО прикрепления

          ENDIF 

         OTHERWISE 
          rval = InsError('S', 'G5A', m.recid, '', ;
      	 	'Направления на госпитализацюи в ведомственный ('+STR(sprved.prn_kodved,3)+;
      	 		') стационар (sprved) выдало МО иного ведомства: '+STR(m.k_ved))
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

       ENDCASE 
       ENDIF 
      ENDIF 
     ENDIF 
    ENDIF 
    
    DO CASE 
     CASE m.ord = 0 && не контролируется!

     CASE m.ord = 1
      m.lp_ved = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.prn_kodved, 0)
      IF INLIST(m.lp_ved,48,55,100)
       rval = InsError('S', 'G1A', m.recid, '', ;
       	'Недопустимое ord ('+STR(m.ord,1)+' для МО, входящего в перечень подведомственных федеральных органам власти ('+STR(m.k_ved,3)+')')
      ELSE 
       IF m.d_pos<{03.04.2020}
        IF !SEEK(m.lpu_ord, 'sprlpu') AND !SEEK(m.lpu_ord, 'f003')
         rval = InsError('S', 'G1A', m.recid, '', ;
       	 'Недопустимое lpu_ord ('+STR(m.lpu_ord,4)+') при ord='+STR(m.ord,1)+;
       	 	' (допустимые lpu_ord из перечня lpu_id системы ОМС г. Москвы либо F003 (до 03.04.2020) профиль '+m.pr_gosp)
        ENDIF 
       ELSE 
        DO CASE 
         CASE SEEK(m.lpu_ord, 'f003')
         CASE m.lpu_ord=4963 && МНПЦДК
         CASE SEEK(m.lpu_ord, 'wom')
         CASE m.lpu_ord=m.lpuid AND m.d_type='6'
         CASE INLIST(m.pr_gosp,'018','060','017','029','021','122') AND SEEK(m.lpu_ord, 'sprlpu')
         CASE m.pr_diag='C' OR BETWEEN(LEFT(m.pr_diag,3),'D00','D09')
         CASE INLIST(INT(m.cod/1000),49) OR BETWEEN(m.cod,97008,97010) OR BETWEEN(m.cod,197010,197011)
         CASE OCCURS('#', m.c_i)=3
         *CASE SEEK(m.lpu_ord, 'pilot') OR IIF(m.lpuid=5139, SEEK(m.lpu_ord, 'pilots'), .F.)
         CASE SEEK(m.lpu_ord, 'pilot') OR SEEK(m.lpu_ord, 'pilots')
          * Здесь проверить МО прикрепления
          IF !EMPTY(m.date_ord)
           m.lpu_pr = 0
           m.d_att  = {}
		   DO CASE 
		    CASE m.tipp='В'
		     m.polis = ALLTRIM(sn_pol)
		     IF LEN(m.polis)=9
		      m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.lpu_tera, 0)
		      m.lpu_st   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.lpu_stom, 0)
		      m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.date_tera, {})
		      m.d_st     = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.date_stom, {})
		     ENDIF 

		    CASE INLIST(m.tipp,'П','Э','К')
		     m.polis   = LEFT(sn_pol,16)
		     m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.lpu_tera, 0)
		     m.lpu_st   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.lpu_stom, 0)
 	         m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.date_tera, {})
 	         m.d_st     = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.date_stom, {})

		    CASE m.tipp='С'
		     m.polis = ALLTRIM(sn_pol)
		     m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.lpu_tera, 0)
		     m.lpu_st   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.lpu_stom, 0)
		     m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.date_tera, {})
		     m.d_st     = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.date_stom, {})

		    OTHERWISE 
		   ENDCASE 
		    
		   *IF (EMPTY(m.d_att) OR m.d_att<=m.date_ord) AND IIF(m.lpuid=5139, EMPTY(m.d_st) OR m.d_st<=m.date_ord, .T.)
		   IF (EMPTY(m.d_att) OR m.d_att<=m.date_ord) AND (EMPTY(m.d_st) OR m.d_st<=m.date_ord)
            
		    *IF (!EMPTY(m.d_att) AND m.lpu_pr<>m.lpu_ord) AND (IIF(m.lpuid=5139, (!EMPTY(m.d_st) AND m.lpu_st<>m.lpu_ord), .T.)) && если не прикреплен ни к кому, то не бракуем!
		    IF (!EMPTY(m.d_att) AND (m.lpu_pr>0 AND m.lpu_pr<>m.lpu_ord)) AND (!EMPTY(m.d_st) AND (m.lpu_st>0 AND m.lpu_st<>m.lpu_ord)) && если не прикреплен ни к кому, то не бракуем!
		     *IF m.qcod<>'I3'
             rval = InsError('S', 'G1A', m.recid, '', ;
       	      'МО направления ('+STR(m.lpu_ord,4)+') не соответствует МО прикрепления ('+STR(m.lpu_pr,4)+;
       	      ') на дату выдачи направления '+DTOC(m.date_ord)+', d_att='+DTOC(m.d_att))
       	     *ENDIF 
		    ENDIF
		   
		   ELSE &&  взять предыдущий номерник 

            m.lpu_pr = 0
            m.d_att  = {}
		    DO CASE 
		     CASE m.tipp='В'
		      m.polis = ALLTRIM(sn_pol)
		      IF LEN(m.polis)=9
		       m.lpu_pr   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_tera, 0)
		       m.lpu_st   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_stom, 0)
		       m.d_att    = IIF(SEEK(m.polis, 'vsn'), vsn.date_tera, {})
		       m.d_st     = IIF(SEEK(m.polis, 'vsn'), vsn.date_stom, {})
		      ENDIF 

		     CASE INLIST(m.tipp,'П','Э','К')
		      m.polis   = LEFT(sn_pol,16)
		      m.lpu_pr   = IIF(SEEK(m.polis, 'enp'), enp.lpu_tera, 0)
		      m.lpu_st   = IIF(SEEK(m.polis, 'enp'), enp.lpu_stom, 0)
 	          m.d_att    = IIF(SEEK(m.polis, 'enp'), vsn.date_tera, {})
 	          m.d_st     = IIF(SEEK(m.polis, 'enp'), vsn.date_stom, {})

		     CASE m.tipp='С'
		      m.polis = ALLTRIM(sn_pol)
		      m.lpu_pr   = IIF(SEEK(m.polis, 'kms'), kms.lpu_tera, 0)
		      m.lpu_st   = IIF(SEEK(m.polis, 'kms'), kms.lpu_stom, 0)
		      m.d_att    = IIF(SEEK(m.polis, 'kms'), vsn.date_tera, {})
		      m.d_st     = IIF(SEEK(m.polis, 'kms'), vsn.date_stom, {})
	         OTHERWISE 
		    ENDCASE 

		    *IF (!EMPTY(m.d_att) AND m.d_att<=m.date_ord) AND IIF(m.lpuid=5139, (!EMPTY(m.d_st) OR m.d_st<=m.date_ord), .T.)
		    IF (!EMPTY(m.d_att) AND m.d_att<=m.date_ord) AND (!EMPTY(m.d_st) OR m.d_st<=m.date_ord)
		     IF m.lpu_pr>0 AND m.lpu_pr<>m.lpu_ord
		      *IF m.qcod<>'I3'
              rval = InsError('S', 'G1A', m.recid, '', ;
       	       'МО направления ('+STR(m.lpu_ord,4)+') не соответствует МО прикрепления ('+STR(m.lpu_pr,4)+;
       	       ') на дату выдачи направления '+DTOC(m.date_ord)+', d_att='+DTOC(m.d_att))
       	      *ENDIF 
		     ENDIF
		    ELSE 
		     * Здесь открыть номерник предыдущего месяца
		    ENDIF 
		   ENDIF 
          ENDIF && IF !EMPTY(m.date_ord)
          * Здесь проверять МО прикрепления
         OTHERWISE 
         rval = InsError('S', 'G1A', m.recid, '', ;
       	  'Недопустимое lpu_ord ('+STR(m.lpu_ord,4)+') при ord='+STR(m.ord,1)+;
       	  	' (допустимые lpu_ord из перечня lpu_id системы ОМС г. Москвы либо F003  (до 03.04.2020), профиль '+m.pr_gosp)
        ENDCASE 
       ENDIF && IF m.d_pos<{03.04.2020}
      ENDIF && IF INLIST(m.lp_ved,48,55,100)
      
     CASE m.ord=2 
      DO CASE 
       CASE INLIST(m.lpu_ord, 4708, 502009)
       CASE INLIST(m.cod,61410,161410,161411) AND (SEEK(m.lpu_ord, 'sprnco') AND sprnco.trs=1)
       OTHERWISE 
       rval = InsError('S', 'G1A', m.recid, '', ;
       	'Недопустимое lpu_ord ('+STR(m.lpu_ord,4)+') при ord='+STR(m.ord,1)+' (допустимые lpu_ord 4708,502009')
      ENDCASE 
    
     CASE m.ord = 3 && не контролируется!

     CASE m.ord = 4
      IF !SEEK(m.lpu_ord, 'sprlpu') AND !SEEK(m.lpu_ord, 'f003')
       rval = InsError('S', 'G1A', m.recid, '', ;
       	'Недопустимое lpu_ord ('+STR(m.lpu_ord,4)+') при ord='+STR(m.ord,1)+' (допустимые lpu_ord из перечня lpu_id системы ОМС г. Москвы или 4708')
      ENDIF 

     CASE m.ord = 5
      m.lp_ved = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.prn_kodved, 0)
      IF !INLIST(m.lp_ved,48,55,100)
       rval = InsError('S', 'G1A', m.recid, '', ;
       	'Недопустимое ord ('+STR(m.ord,1)+') для МО, не входящего в перечень подведомственных федеральных органам власти ('+STR(m.k_ved,3)+')')
      ELSE 
       IF m.d_pos<{03.04.2020}
        IF !SEEK(m.lpu_ord, 'sprlpu') AND !SEEK(m.lpu_ord, 'f003')
         rval = InsError('S', 'G1A', m.recid, '', ;
       	 'Недопустимое lpu_ord ('+STR(m.lpu_ord,4)+') при ord='+STR(m.ord,1)+;
       	 	' (допустимые lpu_ord из перечня lpu_id системы ОМС г. Москвы либо F003 (до 03.04.2020) профиль '+m.pr_gosp)
        ENDIF 
       ELSE 
        DO CASE 
         CASE SEEK(m.lpu_ord, 'f003')
         CASE m.lpu_ord=4963 && МНПЦДК
         CASE SEEK(m.lpu_ord, 'wom')
         CASE m.lpu_ord=m.lpuid AND m.d_type='6'
         CASE INLIST(m.pr_gosp,'018','060','017','029','021','122') AND SEEK(m.lpu_ord, 'sprlpu')
         CASE m.pr_diag='C' OR BETWEEN(LEFT(m.pr_diag,3),'D00','D09')
         CASE INLIST(INT(m.cod/1000),49) OR BETWEEN(m.cod,97008,97010) OR BETWEEN(m.cod,197010,197011)
         CASE OCCURS('#', m.c_i)=3
         CASE SEEK(m.lpu_ord, 'pilot')
          * Здесь проверить МО прикрепления
          IF !EMPTY(m.date_ord) && AND m.date_ord>=GOMONTH(m.tdat1,-3)
           m.lpu_pr = 0
           m.d_att  = {}
		   DO CASE 
		    CASE m.tipp='В'
		     m.polis = ALLTRIM(sn_pol)
		     IF LEN(m.polis)=9
		      m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.lpu_tera, 0)
		      m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'vsn'), outs_n.date_tera, {})
		     ENDIF 

		    CASE INLIST(m.tipp,'П','Э','К')
		     m.polis   = LEFT(sn_pol,16)
		     m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.lpu_tera, 0)
 	         m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'enp'), outs_n.date_tera, {})

		    CASE m.tipp='С'
		     m.polis = ALLTRIM(sn_pol)
		     m.lpu_pr   = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.lpu_tera, 0)
		     m.d_att    = IIF(USED('outs_n') AND SEEK(m.polis,'outs_n', 'kms'), outs_n.date_tera, {})

		    OTHERWISE 
		   ENDCASE 
		    
		   IF EMPTY(m.d_att) OR m.d_att<=m.date_ord
           
		    IF !EMPTY(m.d_att) AND (m.lpu_pr>0 AND m.lpu_pr<>m.lpu_ord)
		     *IF m.qcod<>'I3'
             rval = InsError('S', 'G1A', m.recid, '', ;
       	      'МО направления ('+STR(m.lpu_ord,4)+') не соответствует МО прикрепления ('+STR(m.lpu_pr,4)+;
       	      ') на дату выдачи направления '+DTOC(m.date_ord))
       	     *ENDIF 
		    ENDIF
		   
		   ELSE &&  взять предыдущий номерник 

            m.lpu_pr = 0
            m.d_att  = {}
		    DO CASE 
		     CASE m.tipp='В'
		      m.polis = ALLTRIM(sn_pol)
		      IF LEN(m.polis)=9
		       m.lpu_pr   = IIF(SEEK(m.polis, 'vsn'), vsn.lpu_tera, 0)
		       m.d_att    = IIF(SEEK(m.polis, 'vsn'), vsn.date_tera, {})
		      ENDIF 

		     CASE INLIST(m.tipp,'П','Э','К')
		      m.polis   = LEFT(sn_pol,16)
		      m.lpu_pr   = IIF(SEEK(m.polis, 'enp'), enp.lpu_tera, 0)
 	          m.d_att    = IIF(SEEK(m.polis, 'enp'), vsn.date_tera, {})

		     CASE m.tipp='С'
		      m.polis = ALLTRIM(sn_pol)
		      m.lpu_pr   = IIF(SEEK(m.polis, 'kms'), kms.lpu_tera, 0)
		      m.d_att    = IIF(SEEK(m.polis, 'kms'), vsn.date_tera, {})
	         OTHERWISE 
		    ENDCASE 

		    *IF m.d_att<=m.date_ord AND m.date_ord>=GOMONTH(m.tdat1,-3)
		    IF !EMPTY(m.d_att) AND m.d_att<=m.date_ord && Значит все ОК!
		     IF (m.lpu_pr>0 AND m.lpu_pr<>m.lpu_ord)
		      *IF m.qcod<>'I3'
              rval = InsError('S', 'G1A', m.recid, '', ;
       	       'МО направления ('+STR(m.lpu_ord,4)+') не соответствует МО прикрепления ('+STR(m.lpu_pr,4)+;
       	       ') на дату выдачи направления '+DTOC(m.date_ord))
       	      *ENDIF 
		     ENDIF
		    ELSE 
		     * Здесь открыть номерник предыдущего месяца
		     * Здесь открыть номерник предыдущего месяца
		    ENDIF 
		   ENDIF 
          ENDIF 
          * Здесь проверять МО прикрепления
         OTHERWISE 
         rval = InsError('S', 'G1A', m.recid, '', ;
       	  'Недопустимое lpu_ord ('+STR(m.lpu_ord,4)+') при ord='+STR(m.ord,1)+;
       	  	' (допустимые lpu_ord из перечня lpu_id системы ОМС г. Москвы либо F003  (до 03.04.2020), профиль '+m.pr_gosp)
        ENDCASE 
       ENDIF 
      ENDIF 

     CASE m.ord = 6
      IF m.lpu_ord!= 5650 &&9999 && призывники
       rval = InsError('S', 'G1A', m.recid, '', ;
       	'Недопустимое lpu_ord ('+STR(m.lpu_ord,4)+') при ord='+STR(m.ord,1)+' (допустимое lpu_ord 5650')
      ENDIF 
     CASE m.ord = 7
      IF m.lpu_ord!=7665
       rval = InsError('S', 'G1A', m.recid, '', ;
       	'Недопустимое lpu_ord ('+STR(m.lpu_ord,4)+') при ord='+STR(m.ord,1)+' (допустимое lpu_ord 7665')
      ENDIF 
     CASE m.ord = 8
      IF m.lpu_ord!=8888
       rval = InsError('S', 'G1A', m.recid, '', ;
       	'Недопустимое lpu_ord ('+STR(m.lpu_ord,4)+') при ord='+STR(m.ord,1)+' (допустимое lpu_ord 8888')
      ENDIF 
     OTHERWISE 
       rval = InsError('S', 'G1A', m.recid, '', ;
       	'Недопустимое сочетание lpu_ord ('+STR(m.lpu_ord,4)+') и ord='+STR(m.ord,1)+' (допустимые комбинации см. в Правилах файлового обмена')
     ENDCASE 
    
   ENDIF 

   IF M.G2A == .T.
    m.cod    = cod 
    m.sn_pol = sn_pol
    
    m.ord      = ord
    m.date_ord = date_ord
    m.recid    = recid
    m.d_u      = d_u
    
    *IF INLIST(m.ord,1,4,5,6,8,9)
    IF INLIST(m.ord,1,4,5,6,8) && с января 2020
     IF EMPTY(m.date_ord)
      rval = InsError('S', 'G2A', m.recid, '',;
      	'Пустое значение date_ord при ord='+STR(m.ord,1)+' (проверяется для ord=1,4,5,6,8,9)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ELSE 
      IF m.date_ord>m.d_u
       rval = InsError('S', 'G2A', m.recid, '',;
      	'Значение поля date_ord больше даты оказания услуги при '+STR(m.ord,1)+' (проверяется для ord=1,4,5,6,8,9)')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.G3A == .T.
    m.ord     = ord
    m.cod     = cod 
    m.sn_pol  = sn_pol
    m.recid   = recid
    m.lpu_ord = lpu_ord
    m.n_u     = ALLTRIM(n_u)
    
    IF (!EMPTY(m.n_u) AND TYPE(m.n_u)<>'N') AND m.ord<>5
     rval = InsError('S', 'G3A', m.recid, '', ;
     	'Недопустимые символы в n_u при ord!=5')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    IF m.ord=2
     IF EMPTY(m.n_u)
      rval = InsError('S', 'G3A', m.recid, '', ;
     	'Пустой номер наряда на госпитализацию (n_u) при экстренной госпитализации (ord=2)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ELSE 
      *IF INLIST(m.lpu_ord,4708,774708,502009) AND (TYPE(m.n_u)<>'N' OR LEN(ALLTRIM(m.n_u))<>9)
      *IF INLIST(m.lpu_ord,4708,774708,502009) AND (TYPE(m.n_u)<>'N')
      * rval = InsError('S', 'G3A', m.recid, '', ;
      *	'Поле n_u (номер наряда на госпитализацию) имеет не числовое значение.')
      * m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      *ENDIF 
     ENDIF 
    ENDIF 

    IF m.ord=5 AND EMPTY(m.n_u) 
     rval = InsError('S', 'G3A', m.recid, '', ;
     	'Пустой номер наряда на госпитализацию (n_u) при госпитализации в федеральные МО (ord=5)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF FIELD('n_vmp')='N_VMP'
     m.n_vmp   = ALLTRIM(n_vmp)
     IF IsVmp(m.cod)
      IF EMPTY(m.n_vmp)
       rval = InsError('S', 'G3A', m.recid, '',;
      	'Отсутствие талона ВМП (не заполнено поле n_vmp) при оказании услуг 200-х и/или 300-х групп')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ELSE 
       IF !IsN_VMP(m.n_vmp)
        rval = InsError('S', 'G3A', m.recid, '',;
      	'Номер талона-направления на ВМП ('+m.n_vmp+') не сответствует шаблону 99.9999.99999.999')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ENDIF 
     ENDIF 
     IF m.cod=97041 AND !IsOkNVmpForEco(m.n_vmp)
      rval = InsError('S', 'G3A', m.recid, '',;
      	'Отсутствие талона ВМП (не заполнено поле n_vmp) при оказании ЭКО')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 

    ENDIF 
    
   ENDIF 

   IF M.G4A == .T.
    m.cod = cod 
    m.sn_pol = sn_pol
    IF (IsMes(m.cod) AND !BETWEEN(m.cod,97107,97999)) OR IsVmp(m.cod)
     m.recid = recid
     m.ord   = ord
     m.ds_0 = ALLTRIM(ds_0)
     IF INLIST(m.ord,1,2,6) AND EMPTY(m.ds_0)
      rval = InsError('S', 'G4A', m.recid,'',;
      	'Пустой направительный диагноз (ds_0) для госпитализированного пациента')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.UVA == .T.
    m.cod = cod
    m.d_u = d_u
    m.ldr = people.dr
    m.ddr = IIF(OCCURS('#',c_i)=3 AND m.cod>100000, ;
     CTOD(SUBSTR(SUBSTR(c_i, AT('#',c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(c_i, AT('#',c_i,2)+1,8),5,2)+'.'+SUBSTR(SUBSTR(c_i, AT('#',c_i,2)+1,8),1,4)), ;
     people.dr)
    IF !EMPTY(m.ddr)
     nmonthes = ((m.d_u-m.ddr)/365.25)*12 && Переделал с m.dat2 на m.d_u 13.03.2017 по просьбе УралСиба
*     nmonthes = ((m.dat2-m.ddr)/365.25)*12 && Переделал с m.dat2 на m.d_u 13.03.2017 по просьбе УралСиба
*     nmonthes = ((m.ldr-m.ddr)/365.25)*12
     IF SEEK(cod, 'codwdr') AND (!BETWEEN(nmonthes, IIF(BETWEEN(cod,1821,1825), 0, CodWDr.min_ms), CodWDr.max_ms) AND ;
       (!INLIST(d_type,'1','5','6','e') AND ;
       (!SEEK(ds, 'nocodr', 'ds1') AND !SEEK(ds, 'nocodr', 'ds2') AND !SEEK(ds, 'nocodr', 'ds3'))))
      m.recid = recid
      rval = InsError('S', 'UVA', m.recid, '',;
      	'Возраст пациента '+ALLTRIM(STR(m.nmonthes))+' не попадает в допустимый интервал справочника codwdr: c '+;
      		ALLTRIM(STR(CodWDr.min_ms))+' по '+ALLTRIM(STR(CodWDr.max_ms)))
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.DVA == .T. && если нет в mesmkb то не проверяем!!!
    m.cod = cod 
    m.ds   = ds
    m.ds_0 = ds_0
    m.ds_2 = ds_2
    m.ds_3 = ds_3
    m.d_type = d_type
    
    IF INLIST(INT(m.cod/1000),43,143) AND SEEK(m.cod, 'codwdr')
     IF LEFT(m.ds,1)=LEFT(codwdr.z_ds,1) AND m.ds<>codwdr.z_ds
      m.recid = recid
      rval = InsError('S', 'DVA', m.recid, '',;
      	'Недопустимое сочетание услуги '+STR(m.cod,6)+' с диагнозом (допустим '+codwdr.z_ds+' по справочнику)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
    IF INLIST(m.cod,28165,128165) AND !INLIST(m.ds, 'U071','Z03.8', 'Z11.5')
      m.recid = recid
      rval = InsError('S', 'DVA', m.recid, '',;
      	'Недопустимое сочетание услуги '+STR(m.cod,6)+' с диагнозом '+m.ds+'(c 202012!)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF INLIST(m.cod, 97158, 81094) AND m.ds='C97'
     DO CASE 
      CASE INLIST(m.ds_0, 'C18.0', 'C18.1', 'C18.2', 'C18.3', 'C18.4', 'C18.5', 'C18.6', 'C18.7', 'C18.8', 'C18.9', 'C19', 'C20')
      CASE INLIST(m.ds_0, 'C21', 'C21.0', 'C21.1', 'C21.2', 'C21.8', 'C34.0', 'C34.1', 'C34.2', 'C34.3', 'C34.8', 'C34.9', 'C43.0')
      CASE INLIST(m.ds_0, 'C43.1', 'C43.2', 'C43.3', 'C43.4', 'C43.5', 'C43.6', 'C43.7', 'C43.8', 'C43.9', 'C50.0', 'C50.1', 'C50.2')
      CASE INLIST(m.ds_0, 'C50.3', 'C50.4', 'C50.5', 'C50.6', 'C50.8', 'C50.9', 'C61', 'C64')

      CASE INLIST(m.ds_2, 'C18.0', 'C18.1', 'C18.2', 'C18.3', 'C18.4', 'C18.5', 'C18.6', 'C18.7', 'C18.8', 'C18.9', 'C19', 'C20')
      CASE INLIST(m.ds_2, 'C21', 'C21.0', 'C21.1', 'C21.2', 'C21.8', 'C34.0', 'C34.1', 'C34.2', 'C34.3', 'C34.8', 'C34.9', 'C43.0')
      CASE INLIST(m.ds_2, 'C43.1', 'C43.2', 'C43.3', 'C43.4', 'C43.5', 'C43.6', 'C43.7', 'C43.8', 'C43.9', 'C50.0', 'C50.1', 'C50.2')
      CASE INLIST(m.ds_2, 'C50.3', 'C50.4', 'C50.5', 'C50.6', 'C50.8', 'C50.9', 'C61', 'C64')

      CASE INLIST(m.ds_3, 'C18.0', 'C18.1', 'C18.2', 'C18.3', 'C18.4', 'C18.5', 'C18.6', 'C18.7', 'C18.8', 'C18.9', 'C19', 'C20')
      CASE INLIST(m.ds_3, 'C21', 'C21.0', 'C21.1', 'C21.2', 'C21.8', 'C34.0', 'C34.1', 'C34.2', 'C34.3', 'C34.8', 'C34.9', 'C43.0')
      CASE INLIST(m.ds_3, 'C43.1', 'C43.2', 'C43.3', 'C43.4', 'C43.5', 'C43.6', 'C43.7', 'C43.8', 'C43.9', 'C50.0', 'C50.1', 'C50.2')
      CASE INLIST(m.ds_3, 'C50.3', 'C50.4', 'C50.5', 'C50.6', 'C50.8', 'C50.9', 'C61', 'C64')
      
      OTHERWISE 
       m.recid = recid
       rval = InsError('S', 'DVA', m.recid, '',;
      	'Услуги 97158 / 81094 с DS=С97 без заполнения одного из DS_2/DS_0 или DS_3 кодом из перечня C18.0...')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDCASE 
    ENDIF 
    
    IF (IsMes(m.cod) OR IsVMP(m.cod) OR INLIST(m.cod, 97158, 81094)) AND !INLIST(SUBSTR(PADL(Cod,6,'0'),2,2), '83', '84')
    
     IF SEEK(m.cod, 'MesMkb', 'cod')
    
      m.IsOtdSkp = IIF(SUBSTR(otd,2,2)='09', .T., .F.)

      m.perem = IIF(!ISDIGIT(SUBSTR(Ds,5,1)), STR(Cod,6)+' '+LEFT(Ds,3)+'   ', STR(Cod,6)+' '+Ds)
      m.p_1 = IIF(!ISDIGIT(SUBSTR(Ds,5,1)), STR(Cod,6)+' '+LEFT(Ds,3)+'   ', STR(Cod,6)+' '+LEFT(Ds,5)+' ')

      IF (!SEEK(IIF(!m.IsOtdSkp, m.perem, LEFT(m.perem,5)), IIF(m.IsOtdSkp, 'ReesKp', 'MesMkb'), 'ds_ms')) AND ;
     	(!SEEK(IIF(!m.IsOtdSkp, m.p_1, LEFT(m.p_1,5)), IIF(m.IsOtdSkp, 'ReesKp', 'MesMkb'), 'ds_ms'))
      
       IF !(INLIST(cod,97158,81094) OR INLIST(m.cod,61410,161410,161411,61400,161400,161401) OR BETWEEN(m.cod,70150,70170) OR BETWEEN(m.cod,170150,170171))
		IF !INLIST(INT(m.cod/1000),200,297)
         IF !INLIST(m.d_type,'1','5','6')
          m.recid = recid
          rval = InsError('S', 'DVA', m.recid, '',;
      	   IIF(m.IsOtdSkp, 'Сочетание МЭС-Диагноз, применненное в отделении СКП, не найдено в справочнике ReesKp (поиск по выражению '+m.perem+')', ;
      		'Сочетание МЭС-Диагноз не найдено в справочнике MesMkb (поиск по выражению '+m.perem+')'))
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 
        ELSE 
         m.recid = recid
         rval = InsError('S', 'DVA', m.recid, '',;
      	 IIF(m.IsOtdSkp, 'Сочетание МЭС-Диагноз, применненное в отделении СКП, не найдено в справочнике ReesKp (поиск по выражению '+m.perem+')', ;
      	  'Сочетание МЭС-Диагноз не найдено в справочнике MesMkb (поиск по выражению '+m.perem+')'))
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 

       ELSE 
        m.recid = recid
        rval = InsError('S', 'DVA', m.recid, '',;
         IIF(m.IsOtdSkp, 'Сочетание МЭС-Диагноз, применненное в отделении СКП, не найдено в справочнике ReesKp (поиск по выражению '+m.perem+')', ;
      		'Сочетание МЭС-Диагноз не найдено в справочнике MesMkb (поиск по выражению '+m.perem+')'))
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      
      ENDIF 
     
     ENDIF  && IF SEEK(m.cod, 'MesMkb', 'cod')
    
    ENDIF
   ENDIF && IF M.DVA == .T.

   && Проверка по "y" && Только такие услуги в этом отделении
   IF M.UOA == .T.
    m.cod = cod
    m.otd = otd
    m.d_type = d_type
    m.usl_ok  = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), profot.usl_ok, '0')

    
    * Эта часть алгоритма отнесена только для иногородних пациентов!
    *SET ORDER TO notd IN CodOtd
    *m.IsCheck = IIF(SEEK(SUBSTR(otd,4,6), 'CodOtd', 'notd'), .T., .F.)
    
    *IF VAL(m.gcPeriod)>201909
    * m.IsCheck = IIF(m.cod = 97001 AND INLIST(SUBSTR(otd,2,2),'80','81') AND SUBSTR(otd,4,3)='136', .F., m.IsCheck)
    *ENDIF 

    *IF IsCheck AND d_type!='2'
    * IsOk = .f.
    * DO WHILE SUBSTR(otd,4,6) = CodOtd.otd
    *  IF cod = CodOtd.Cod
    *   IsOk = .t.
    *   EXIT
    *  ENDIF
    *  SKIP IN CodOtd
    * ENDDO 
    
    * IF IsOk = .f.
    *  m.recid = recid
    *  rval = InsError('S', 'UOA', m.recid)
    *  m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    * ENDIF 
    *ENDIF 
    * Эта часть алгоритма отнесена только для иногородних пациентов!
    
    IF INLIST(INT(m.cod/1000),29,129,59,159) AND SUBSTR(m.otd,4,3)<>'067'
     m.recid = recid
     rval = InsError('S', 'UOA', m.recid, '', ;
     	'Оказание услуги из раздела 29,129,59,159 в отделении не 067 (4-6 разряд фасетного кода)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    IF (SUBSTR(m.otd,2,2)='85' OR SUBSTR(m.otd,4,3)='067') AND !INLIST(INT(m.cod/1000),29,129,59,159)
     m.recid = recid
     rval = InsError('S', 'UOA', m.recid, '', ;
     	'Оказание услуги из не раздела 29,129,59,159 в отделении 85 (2-3 разряд фасетного кода)')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF (m.usl_ok='1' AND INLIST(SUBSTR(m.otd,4,3),'005','167')) AND ;
    	!(INLIST(INT(m.cod/1000),83,183,49,149,51,151,52,252,53,153,54,154,55,155) OR INLIST(m.cod,56029,156003))
     IF !(m.d_type='5' AND IsVMP(m.cod))
      m.recid = recid
      rval = InsError('S', 'UOA', m.recid, '', ;
     	'Оказание услуги из не раздела 83,183,149,149,51,52,53,54,55,151,152,153,154,155 и не 56029,156003 в отделении 005/167 (4-6 разряд фасетного кода)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
    * С 202010
    IF INLIST(INT(m.cod/1000),83,183) AND !(m.usl_ok='1' AND INLIST(SUBSTR(m.otd,4,3),'005','167'))
      m.recid = recid
      rval = InsError('S', 'UOA', m.recid, '', ;
     	'Оказание услуги из раздела 83,183 в отделении 005/167 (4-6 разряд фасетного кода)')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    
   ENDIF 
   
   IF !(INLIST(m.cod,25050,25203,25243,25271,25268,26003,26087,26158,26180,26229,26243,26275,27005,27009,27010,27015,27016,27019,27020,27024,27027) OR ;
   	  INLIST(m.cod,27028,27030,27032,27050,27061,27071,28019,28021,28024,28050,28067,28077,28093,28097,28119,28126,28144,28147,28164,28188,28190,28208) OR ;
   	  INLIST(m.cod,125050,125203,125243,125271,125268,126003,126087,126158,126180,126229,126243,126275,127005,127009,127010,127015,127016,127019,127020,127024,127027) OR ;
   	  INLIST(m.cod,127028,127030,127032,127050,127061,127071,128019,128021,128024,128050,128067,128077,128093,128097,128119,128126,128144,128147,128164,128188,128190,128208))
      
   IF M.NOA == .T.
    m.perem = sn_pol+str(cod,6)+dtos(d_u)
    *IF SEEK(m.perem, 'e_day') AND d_type!='2'
    IF SEEK(m.perem, 'e_day')

     m.ocntr = e_day.cntr
     m.ncntr = m.ocntr + k_u
     IF m.ncntr<=e_day.k_norm
      REPLACE e_day.cntr WITH m.ncntr IN e_day
     ELSE 
      REPLACE e_day.cntr WITH m.ncntr IN e_day && !!!
      *IF IIF(!INLIST(INT(cod/1000),49,149), !INLIST(d_type,'2','8'), .T.) AND OCCURS('#', m.c_i)<3
      IF OCCURS('#', m.c_i)<3
       m.recid = recid
       rval = InsError('S', 'NOA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF  && !!!
     ENDIF 

    ENDIF 
   ENDIF 
   
   IF M.NMA == .T.
    m.perem = sn_pol+str(cod,6)
    *IF SEEK(m.perem, 'e_month') AND d_type!='2'
    IF SEEK(m.perem, 'e_month')
     m.ocntr = e_month.cntr
     m.ncntr = m.ocntr + k_u
     IF m.ncntr<=e_month.k_norm
      REPLACE e_month.cntr WITH m.ncntr IN e_month
     ELSE 
      REPLACE e_month.cntr WITH m.ncntr IN e_month
      *IF !INLIST(d_type,'2','8') AND OCCURS('#', m.c_i)<3 && !!!
      IF OCCURS('#', m.c_i)<3
       m.recid = recid
       rval = InsError('S', 'NMA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
   ENDIF 

   ENDIF 
   
   IF M.D4A == .T.
    m.sn_pol = sn_pol
    m.cod = cod
    m.dsptip = IIF(SEEK(m.cod,'dspcodes'), dspcodes.tip, 0)
    m.dsptip = IIF(!INLIST(m.cod, 25204, 35401), m.dsptip, 0)
    
    m.d_u  = d_u
    m.w    = IIF(SEEK(m.sn_pol, 'people'), people.w, 0)
    m.dr   = IIF(SEEK(m.sn_pol, 'people'), people.dr, {})
    m.adj  = CTOD(STRTRAN(DTOC(m.dr), STR(YEAR(m.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.adj  = 0
    m.vozr = (YEAR(m.d_u) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)
    
    m.recid = recid

    IF m.dsptip=2
     DO CASE 
      CASE (m.cod = 1949 AND (m.w<>2 OR !INLIST(m.vozr,19,21,23,25,27,29,31,33)))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1950 AND (m.w<>2 OR !INLIST(m.vozr,18,20,22,24,26,28,30,32,34)))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1951 AND (m.w<>2 OR !INLIST(m.vozr,35,37,39)))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1952 AND (m.w<>2 OR !INLIST(m.vozr,36,38)))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1953 AND ;
      	(m.w<>2 OR (!INLIST(m.vozr,40,42,44,46,48,50,52,54,56,58,60,62,64,66,68) AND ;
      		!INLIST(m.vozr,70,72,74,76,78,80,82,84,86,88,90,92,94,96,98))))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1954 AND ;
      	(m.w<>2 OR (!INLIST(m.vozr,41,43,45,47,49,51,53,55,57,59,61,63,65,67,69,71) AND ;
      		!INLIST(m.vozr,73,75,77,79,81,83,85,87,89,91,93,95,97,99))))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1968 AND (m.w<>1 OR !INLIST(m.vozr,19,21,23,25,27,29,31,33)))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1969 AND (m.w<>1 OR !INLIST(m.vozr,18,20,22,24,26,28,30,32,34)))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1970 AND (m.w<>1 OR !INLIST(m.vozr,35,37,39)))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1971 AND (m.w<>1 OR !INLIST(m.vozr,36,38)))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1972 AND ;
      	(m.w<>1 OR (!INLIST(m.vozr,40,42,44,46,48,50,52,54,56,58,60,62,64,66,68) AND ;
      		!INLIST(m.vozr,70,72,74,76,78,80,82,84,86,88,90,92,94,96,98))))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1973 AND ;
      	(m.w<>1 OR (!INLIST(m.vozr,41,43,45,47,49,51,53,55,57,59,61,63,65,67,69,71) AND ;
      		!INLIST(m.vozr,73,75,77,79,81,83,85,87,89,91,93,95,97,99))))
       m.cmnt = 'Профилактическая услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'DKA', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)


     ENDCASE 
    ENDIF 

    IF m.dsptip=1
     DO CASE 
      CASE (m.cod = 1925 AND (m.w<>2 OR !INLIST(m.vozr,41,43,47,49,53,55,59,61,77,79,81,83,85,87,89,91,93,95,97,99)))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE (m.cod = 1926 AND (m.w<>2 OR !INLIST(m.vozr,65,67,69,71,73,75,76,78,80,82,84,86,88,90,92,94,96,98)))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1927 AND (m.w<>2 OR !INLIST(m.vozr,21,27,33,40,44,46,50,52,56,58,62,64,66,68,70,72,74))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1928 AND (m.w<>2 OR !INLIST(m.vozr,18,24,30,36,39))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1929 AND (m.w<>2 OR !INLIST(m.vozr,42,45,48,51,54,57,60,63))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1930 AND (m.w<>1 OR !INLIST(m.vozr,18,21,24,27,30,33,36,39))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1931 AND (m.w<>1 OR !INLIST(m.vozr,41,43,47,49,51,53,57,59,61,63,77,79,81,83,85,87,89,91,93,95,97,99))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1932 AND (m.w<>1 OR !INLIST(m.vozr,65,67,69,71,73,75,76,78,80,82,84,86,88,90,92,94,96,98))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1933 AND (m.w<>1 OR !INLIST(m.vozr,40,42,44,46,48,52,54,56,58,62,66,68,70,72,74))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1934 AND (m.w<>1 OR !INLIST(m.vozr,50,55,60,64))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1935 AND (m.w<>1 OR !INLIST(m.vozr,45))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1936 AND (m.w<>2 OR !INLIST(m.vozr,18,24,30))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      * с этого места новые коды
      CASE m.cod = 1937 AND (m.w<>2 OR !INLIST(m.vozr,21,27,33))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1938 AND (m.w<>2 OR !INLIST(m.vozr,36))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1939 AND (m.w<>2 OR !INLIST(m.vozr,39))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1940 AND (m.w<>2 OR !INLIST(m.vozr,40,44,46,50,52,56,58,62,64))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1941 AND (m.w<>2 OR !INLIST(m.vozr,41,43,47,49,53,55,59,61))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1942 AND (m.w<>2 OR !INLIST(m.vozr,42,48,54,60))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1943 AND (m.w<>2 OR !INLIST(m.vozr,45))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1944 AND (m.w<>2 OR !INLIST(m.vozr,51,57,63))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1945 AND (m.w<>2 OR !INLIST(m.vozr,65,67,69,71,73,75))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1946 AND (m.w<>2 OR !INLIST(m.vozr,66,68,70,72,74))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1947 AND (m.w<>2 OR !INLIST(m.vozr,76,78,80,82,84,86,88,90,92,94,96,98))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1948 AND (m.w<>2 OR !INLIST(m.vozr,77,79,81,83,85,87,89,91,93,95,97,99))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1955 AND (m.w<>1 OR !INLIST(m.vozr,18,24,30))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1956 AND (m.w<>1 OR !INLIST(m.vozr,21,27,33))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1957 AND (m.w<>1 OR !INLIST(m.vozr,36))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1958 AND (m.w<>1 OR !INLIST(m.vozr,39))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1959 AND (m.w<>1 OR !INLIST(m.vozr,40,42,44,46,48,52,54,56,58,62))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1960 AND (m.w<>1 OR !INLIST(m.vozr,41,43,47,49,51,53,57,59,61,63))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1961 AND (m.w<>1 OR !INLIST(m.vozr,45))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1962 AND (m.w<>1 OR !INLIST(m.vozr,50,60,64))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1963 AND (m.w<>1 OR !INLIST(m.vozr,55))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1964 AND (m.w<>1 OR !INLIST(m.vozr,65,67,69,71,73,75))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1965 AND (m.w<>1 OR !INLIST(m.vozr,66,68,70,72,74))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1966 AND (m.w<>1 OR !INLIST(m.vozr,76,78,80,82,84,86,88,90,92,94,96,98))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

      CASE m.cod = 1967 AND (m.w<>1 OR !INLIST(m.vozr,77, 79,81,83,85,87,89,91,93,95,97,99))
       m.cmnt = 'Диспансерная услуга '+PADL(m.cod,6,'0')+' не соответствует полу '+STR(m.w,1)+' или  возрасту ('+STR(m.vozr,2)+' лет)'
       rval = InsError('S', 'D4A', m.recid, '', m.cmnt)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

     ENDCASE 
    ENDIF 
   ENDIF 

   IF M.D2A == .T. AND USED('dspp')
    m.cod = cod
    m.dsptip = IIF(SEEK(m.cod,'dspcodes'), dspcodes.tip, 0)
    m.dsptip = IIF(!INLIST(m.cod, 25204, 35401), m.dsptip, 0)
    
    m.d_u = d_u
    m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.vozr = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)
    
    IF m.dsptip = 1
     m.perem = LEFT(sn_pol,17)+STR(2,1)
     IF SEEK(m.perem, 'dspp')
      DO WHILE LEFT(dspp.sn_pol,17)+STR(dspp.tip,1) = m.perem
       IF YEAR(m.d_u) = YEAR(dspp.d_u) AND EMPTY(dspp.er)
        m.recid = recid
        m.cmnt = 'Застрахованному '+ALLTRIM(sn_pol)+' в том же году ('+DTOC(dspp.d_u)+') была оказана услуга '+PADL(dspp.cod,6,'0')+;
      	' (Профосмотр взрослого населения) в МО: '+dspp.mcod
        rval = InsError('S', 'D6A', m.recid, '1', m.cmnt)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      SKIP IN dspp
      ENDDO 
     ENDIF  
    ENDIF 
    
    IF m.dsptip = 2
     m.perem = LEFT(sn_pol,17)+STR(1,1)
     IF SEEK(m.perem, 'dspp')
      DO WHILE LEFT(dspp.sn_pol,17)+STR(dspp.tip,1) = m.perem
       IF YEAR(m.d_u) = YEAR(dspp.d_u) AND EMPTY(dspp.er)
        m.recid = recid
        m.cmnt = 'Застрахованному '+ALLTRIM(sn_pol)+' в том же году ('+DTOC(dspp.d_u)+') была оказана услуга '+PADL(dspp.cod,6,'0')+;
      	' (Профосмотр взрослого населения) в МО: '+dspp.mcod
        rval = InsError('S', 'D6A', m.recid, '1', m.cmnt)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      SKIP IN dspp
      ENDDO 
     ENDIF  
    ENDIF 
    
    IF m.dsptip = 3
     m.perem = LEFT(sn_pol,17)+STR(4,1)
     IF SEEK(m.perem, 'dspp')
      DO WHILE LEFT(dspp.sn_pol,17)+STR(dspp.tip,1) = m.perem
       IF YEAR(m.d_u) = YEAR(dspp.d_u)  AND EMPTY(dspp.er)
        m.recid = recid
        m.cmnt = 'Застрахованному '+ALLTRIM(sn_pol)+' в том же году ('+DTOC(dspp.d_u)+') была оказана услуга '+PADL(dspp.cod,6,'0')+;
      	' (Профосмотр взрослого населения) в МО: '+dspp.mcod
        rval = InsError('S', 'D6A', m.recid, '1', m.cmnt)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      SKIP IN dspp
      ENDDO 
     ENDIF  
    ENDIF 

    IF m.dsptip > 0 && AND !INLIST(m.dsptip,2,4)
    
    DO CASE 
     CASE m.dsptip = 1  && Диспансеризция взрослых, первый этап, tip=1, муж: 1936-1948, жен: 1955-1967
      m.lastt = IIF(m.vozr<40, dspcodes.last, 12)
     CASE m.dsptip = 2 && ПМО взрослых, tip=2, 1949-1973
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
     IF SEEK(m.perem, 'dspp')
      DO WHILE LEFT(dspp.sn_pol,17)+PADL(dspp.tip,1,'0') = m.perem
       IF IIF(m.lastt>=12, YEAR(m.d_u)-YEAR(dspp.d_u)<m.lastt/12, (m.d_u - dspp.d_u)/30<m.lastt) AND EMPTY(dspp.er) ;
       	AND IIF(INLIST(m.dsptip,2,4), YEAR(m.d_u)=YEAR(dspp.d_u), .T.)
         m.recid = recid
         m.cmnt = 'Застрахованному '+ALLTRIM(sn_pol)+' ранее ('+DTOC(dspp.d_u)+') оказана услуга '+PADL(dspp.cod,6,'0')+;
      	 ' из той же категории, возраст: '+STR(m.vozr,2)+' лет'
         rval = InsError('S', IIF(INLIST(m.dsptip,2,4), 'DKA', 'D2A'), m.recid, '1', m.cmnt)
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       SKIP IN dspp
      ENDDO 
     ENDIF  
    ENDIF 
    
   ENDIF && IF m.dsptip>0
   
   ENDIF && IF M.D2A == .T.

   IF M.DKA == .T. AND USED('dspp')
    m.cod = cod
    m.dsptip = IIF(SEEK(m.cod,'dspcodes'), dspcodes.tip, 0)

    IF !(INLIST(m.dsptip,2,4) OR INLIST(m.cod,101927,101928))
    ELSE 
    
    m.d_u = d_u
    m.adj = CTOD(STRTRAN(DTOC(people.dr), STR(YEAR(people.dr),4), STR(YEAR(m.d_u),4)))-m.d_u
    m.vozr = (YEAR(m.d_u) - YEAR(people.dr)) - IIF(m.adj>0, 1, 0)
    
    m.k_key = LEFT(m.sn_pol,17) + PADL(m.cod,6,"0")
    
    IF SEEK(m.k_key, 'dspp', 'un_tag')
     m.dd_u = dspp.d_u
     IF INLIST(m.cod, 101937, 101945)
      IF m.d_u - m.dd_u < 365 AND m.dd_u >= CTOD('01.01.'+STR(tYear,4)) AND EMPTY(dspp.er)
       m.recid = recid
       rval = InsError('S', 'DKA', m.recid, '',;
       	'Комплексная услуга профнаправления '+STR(m.cod,6)+' оказывалась ранее, -'+DTOC(dspp.d_u))
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ELSE 
      IF dspp.d_u >= CTOD('01.01.'+STR(tYear,4)) AND EMPTY(dspp.er)
       m.recid = recid
       rval = InsError('S', 'DKA', m.recid, '',;
      	'Комплексная услуга профнаправления '+STR(m.cod,6)+' оказывалась ранее, -'+DTOC(dspp.d_u))
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
    
    ENDIF 
   
   ENDIF && IF M.DKA == .T.
   
   IF M.NUA == .T.
    SET ORDER TO ncod IN sovmno
    IF SEEK(cod, 'sovmno') && Алгоритм NU - несовместимые услуги 
     DO WHILE sovmno.cod == cod
      IF SEEK(sn_pol+STR(sovmno.cod_1,6)+DTOS(d_u), 'talon_exp')
       IF (EMPTY(UPPER(sovmno.Stac)) OR (!IsStac AND UPPER(sovmno.Stac)='P') OR ;
        (IsStac AND UPPER(sovmno.Stac)='S')) && AND (d_type != '2' OR talon_exp.d_type != '2')
        m.recid = recid
        IF !SEEK(talon_exp.recid, 'sError') OR (SEEK(talon_exp.recid, 'sError') AND sError.c_err!='NUA')
         *IF !INLIST(d_type,'2','a') OR !INLIST(talon_exp.d_type,'2','a') && вторая часть условия добавлена с 202010
         IF !INLIST(d_type,'a') OR !INLIST(talon_exp.d_type,'a') && вторая часть условия добавлена с 202010
          rval = InsError('S', 'NUA', m.recid, '',;
         	'Услуга несовместима с услугой '+PADL(talon_exp.cod,6,'0')+'оказанной '+DTOC(talon_exp.d_u))
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 
        ENDIF 
       ENDIF 
      ENDIF 
      SKIP +1 IN sovmno 
     ENDDO 
    ENDIF
   ENDIF  

   IF M.NSA == .T.
    SET ORDER TO scod IN sovmno
    IF SEEK(cod, 'sovmno') && Алгоритм NS - несовместимые услуги 
     IsSovmUsl = .F.
     DO WHILE sovmno.cod == cod
      IF SEEK(sn_pol+STR(sovmno.cod_1,6)+DTOS(d_u), 'talon_exp')
       IsSovmUsl = .T.
       EXIT 
      ENDIF 
      SKIP +1 IN sovmno 
     ENDDO 
     IF !IsSovmUsl
      *IF (EMPTY(UPPER(sovmno.Stac)) OR (!IsStac AND UPPER(sovmno.Stac)='P') OR ;
      *   (IsStac AND UPPER(sovmno.Stac)='S')) AND !INLIST(d_type,'2','a')
      IF (EMPTY(UPPER(sovmno.Stac)) OR (!IsStac AND UPPER(sovmno.Stac)='P') OR ;
         (IsStac AND UPPER(sovmno.Stac)='S')) AND !INLIST(d_type,'a')
       m.recid = recid
       rval = InsError('S', 'NSA', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.NSA == .T. AND USED('ns36')
    IF INLIST(cod,36022,136022,36023,136023,36024,136024)
     IsSovmUsl = .F.
     GO TOP IN ns36 
     DO WHILE !EOF('ns36')
      IF SEEK(sn_pol+STR(ns36.cod,6), 'talon_exp')
       IsSovmUsl = .T.
       EXIT 
      ENDIF 
      SKIP +1 IN ns36
     ENDDO 
     IF !IsSovmUsl
      m.recid = recid
      rval = InsError('S', 'NSA', m.recid, '',;
      	'Отсутствует совместо выполняемый МЭС (справочник ns36) с услугой '+STR(m.cod,6))
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 

   IF M.HNA == .T.
    m.cod    = cod
    m.sn_pol = sn_pol
    m.d_u    = d_u
    IF INLIST(m.cod,15001,115001) AND SEEK(m.sn_pol, 'polic_h') AND m.d_u - polic_h.d_u < 365 AND ;
    	polic_h.d_u >= CTOD('01.01.'+STR(tYear,4))
     m.pr_mcod = polic_h.mcod
     m.pr_d_u  = polic_h.d_u
     m.recid = recid
     rval = InsError('S', 'HNA', m.recid, '',;
     	'Пациент ранее обращался в ЦЗ МО '+m.pr_mcod+' '+DTOC(m.pr_d_u))
*     InsErrorSV(m.mcod, 'S', 'HNA', m.recid)
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    RELEASE cod, sn_pol, d_u
   ENDIF 

   IF M.IPA == .T. 
    m.cod    = cod
    IF IsGSP(m.cod) AND IIF(SEEK(m.cod, 'codprv', 'cod'), .T., .F.)
     m.pcod = pcod
     m.prvs = IIF(SEEK(m.pcod, 'doctor'), doctor.prvs, 0)
     m.vir = STR(m.cod,6) + STR(m.prvs,3)
     IF !SEEK(m.vir, 'codprv')
      m.recid = recid
      rval = InsError('S', 'IPA', m.recid)
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.EGA == .T.
    m.cod = cod 
    m.otd = otd
    
    IF INLIST(INT(m.cod/1000),96,196) AND m.mcod<>'0371001'
     m.recid = recid
     rval = InsError('S', 'EGA', m.recid, '',;
     	'Оказание скоропомощных услуг в МО')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    *IF INLIST(SUBSTR(m.otd,2,2),'80','81') AND ;
    	!(INLIST(INT(m.cod/1000),97,197,297,397) OR m.cod=1781 OR ;
    		(INLIST(SUBSTR(m.otd,4,3),'018','060','012') AND INLIST(m.cod,36022,36023,36024,136022,136023,136024)))
    IF INLIST(SUBSTR(m.otd,2,2),'80','81') AND ;
    	!(INLIST(INT(m.cod/1000),97,197,297,397) OR m.cod=1781 OR ;
    		(INLIST(SUBSTR(m.otd,4,3),'060','012') AND INLIST(m.cod,36022,36023,36024)))
     IF m.mcod = '0343999' AND m.cod=200356 && Письмо МГФОМС
     ELSE 
      m.recid = recid
      rval = InsError('S', 'EGA', m.recid, '',;
     	'Оказание в дневном стационаре (фасетный код отделения '+SUBSTR(m.otd,2,2)+') услуги '+PADL(m.cod,6,'0'))
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
    
    DO CASE 
     CASE INLIST(SUBSTR(m.otd,2,2),'00','01','08','85','90','91','92','93') && АПП
      IF SEEK(m.cod, 'codwdr') AND codwdr.stac='s'
       DO CASE 
        CASE INLIST(m.cod,29006,29007)
        CASE SUBSTR(m.otd,2,2)='85' AND INLIST(FLOOR(m.cod/1000),29,129,59,159)
        CASE SUBSTR(m.otd,2,2)='90' AND ;
        	!INLIST(FLOOR(m.cod/1000),49,149,51,52,53,54,55,151,152,153,154,155) AND ;
        	!INLIST(m.cod,36022,36023,36024,136022,136023,136024)
        *CASE SUBSTR(m.otd,2,2)='90' AND ;
        	!INLIST(FLOOR(m.cod/1000),49,149,51,52,53,54,55,151,152,153,154,155) AND ;
        	!INLIST(m.cod,36022,36023,36024)
        CASE SUBSTR(m.otd,2,2)='01' AND m.cod=1780
        *CASE FLOOR(m.cod/1000)=146 AND INLIST(m.lpuid,1912,1940,2049,1874,1909)
        CASE FLOOR(m.cod/1000)=146 AND INLIST(m.lpuid,1874,1909)
        OTHERWISE 
         m.recid = recid
         rval = InsError('S', 'EGA', m.recid, '',;
     	 	'Оказание в амб.-пол. отделении (фасетный код отделения '+SUBSTR(m.otd,2,2)+') услуги '+PADL(m.cod,6,'0'))
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDCASE 
       
      
      ENDIF 
      
     CASE INLIST(SUBSTR(m.otd,2,2),'70','73') && Приемные отделения стационара
      IF BETWEEN(m.cod,1001,1730) OR BETWEEN(m.cod,1801,1830) OR BETWEEN(m.cod,101001,101773)
         m.recid = recid
         rval = InsError('S', 'EGA', m.recid, '',;
     	 	'Оказание в приемно отделении стационара услуги '+PADL(m.cod,6,'0')+' по амбулаторному приему врачей-специалистов')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     CASE INLIST(SUBSTR(m.otd,2,2),'80','81') && ДСТ
     
     OTHERWISE && круглосуточный стационар
      IF IsUsl(m.cod) AND !INLIST(SUBSTR(m.otd,2,2),'70','73')
       DO CASE 
        CASE INLIST(FLOOR(m.cod/1000),51,52,53,54,55,151,152,153,154,155) && симультанные услуги
        CASE INLIST(FLOOR(m.cod/1000),49,149) AND m.cod!=49020
        CASE INLIST(m.cod, 36022, 36023, 36024, 136022, 136023, 136024)
        *CASE INLIST(m.cod, 36022, 36023, 36024)
        CASE INLIST(m.cod, 1781, 101781, 56029, 156003)
        CASE INLIST(FLOOR(m.cod/1000), 99, 29, 129)
        CASE INLIST(FLOOR(m.cod/1000),138) AND m.lpuid=1874
        
        OTHERWISE 
         m.recid = recid
         rval = InsError('S', 'EGA', m.recid, '',;
     	 	'Оказание в профильном отделении стационара (фасетный код отделения '+SUBSTR(m.otd,2,2)+') услуги '+PADL(m.cod,6,'0'))
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDCASE 
      ENDIF 
    ENDCASE 

   ENDIF 

*   IF M.POA == .T. && Диспансеризация в непрофильном ЛПУ
*    m.cod    = cod
**    IF INLIST(m.cod, 101927, 101928) AND !SEEK(m.mcod, 'lpu_m')
*    IF INLIST(m.cod, 101927, 101928)
*     m.recid = recid
*     rval = InsError('S', 'POA', m.recid)
**     InsErrorSV(m.mcod, 'S', 'POA', m.recid)
*     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
*    ENDIF 
*    RELEASE cod, sn_pol, d_u
*   ENDIF 

   IF M.VDA == .T. AND m.tdat1>={01.05.2014} && Просроченный сертификат
    m.pcod = pcod
    m.prvs  = prvs
    m.d_ser = {}
    m.d_u   = d_u
    IF SEEK(m.pcod, 'doctor')
     m.d_ser  = doctor.d_ser
     m.d_ser2 = IIF(FIELD('d_ser2', 'doctor')=UPPER('d_ser2'), IIF(!EMPTY(doctor.d_ser2), doctor.d_ser2, {01.01.0001}), {01.01.0001})
    ENDIF 
    IF !EMPTY(m.d_ser)
     IF (m.d_u-m.d_ser > 365.25*5 AND m.d_ser+365.25*5<{15.03.2020}) AND (m.d_u-m.d_ser2>365.25*5 AND m.d_ser2+365.25*5<{15.03.2020})
     *IF (m.d_u-m.d_ser > 365.25*5) AND (m.d_u-m.d_ser2>365.25*5)
      m.recid = recid
      rval = InsError('S', 'VDA', m.recid, '',;
      	'Сертификат специалиста выдан более пяти лет назад ('+DTOC(m.d_ser)+')')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ENDIF 
    ENDIF 
   ENDIF 
   
   IF M.TPA
    m.c_i    = c_i
    m.sn_pol = sn_pol
    m.d_u    = d_u

    IF OCCURS('#', m.c_i)=3
     m.dr = CTOD(SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),7,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),5,2)+'.'+SUBSTR(SUBSTR(m.c_i,AT('#',m.c_i,2)+1,8),1,4))
    ELSE 
     m.dr = IIF(SEEK(m.sn_pol, 'people'), people.dr, {})
    ENDIF 

    m.adj  = CTOD(STRTRAN(DTOC(m.dr), STR(YEAR(m.dr),4), STR(YEAR(people.d_beg),4)))-people.d_beg
    *m.vozr   = (YEAR(m.d_u) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)
    m.vozr   = (YEAR(people.d_beg) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)
    m.i_otd  = SUBSTR(otd,4,3)
    m.k_u    = k_u
    
    m.tip     = tip
    m.d_type = d_type
    
    DO CASE 
    CASE INLIST(m.d_type,'1','2','5','6') OR m.tip='5'
    CASE m.k_u<=1 AND (SEEK(m.ds, 'nocodr', 'ds1') OR SEEK(m.ds, 'nocodr', 'ds2') OR SEEK(m.ds, 'nocodr', 'ds3'))
    
    OTHERWISE 
    
    IF INLIST(m.i_otd,'017','018','019','020','021','068','086') AND m.vozr>=18
     m.recid = recid
     rval = InsError('S', 'TPA', m.recid, '',;
     	'Пациент старше 18 лет при детском профиле МП (017,018,019,020,021,068,086), с 02.2020')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF INLIST(m.i_otd,'055') AND m.vozr>1
     m.recid = recid
     rval = InsError('S', 'TPA', m.recid, '',;
     	'Пациент страше 1 года при профиле МП=055, с 02.2020')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    IF INLIST(m.i_otd,'014') AND m.vozr<65
     m.recid = recid
     rval = InsError('S', 'TPA', m.recid, '',;
     	'Пациент моложе 65 лет при профиле МП=014, с 02.2020')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 

    IF INLIST(m.i_otd,'029','060','108','122') AND m.vozr<18
     m.recid = recid
     rval = InsError('S', 'TPA', m.recid, '',;
     	'Пациент моложе 18 лет при взрослом профиле МП (029,060,108,122), с 02.2020')
     m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
    ENDIF 
    
    ENDCASE 

   ENDIF 
   
   IF M.CPA==.T.
    m.sn_pol  = sn_pol
    m.cod     = cod 
    m.recid   = recid
    m.recid_s = recid_lpu

    IF INLIST(m.cod,61430,161430,161431) && Только для этих МЭСов проверяем ЛС
     IF !USED('cv_ls')
      rval    = InsError('S', 'CPA', m.recid, '', 'Отсутствует файл CV_LS')
     ELSE 
      IF !SEEK(m.recid_s, 'cv_ls')
       rval    = InsError('S', 'CPA', m.recid,'','Не найдены сведения о ЛС в файле CV_LS')
      ELSE 
       IF !INLIST(cv_ls.sid,'DD0024424','DD0024425','DD0000079','DD0015518','DD0034264','DD0033837','DD0019782','DD0004470','DD0024469','DD0024470')
        rval    = InsError('S', 'CPA', m.recid, '', 'Примено некорректное ЛС при лечение COVID: '+cv_ls.sid)
       ELSE 
        IF EMPTY(cv_ls.date_inj)
         rval    = InsError('S', 'CQA', m.recid, '', 'Не заполнено поле date_inj файла CV_LS')
        ELSE 
         m.date_inj = cv_ls.date_inj
         IF SEEK(m.sn_pol, 'people') AND !BETWEEN(m.date_inj, people.d_beg,people.d_end)
          rval    = InsError('S', 'CQA', m.recid, '', 'Дата date_inj вне диапазона госпитализаци: '+DTOC(people.d_beg)+'-'+DTOC(people.d_end))
         ENDIF 
        ENDIF 
        IF !INLIST(cv_ls.tip_opl,1,2,3,4,5)
         rval    = InsError('S', 'CRA', m.recid, '', 'Некорректное значение поля tip_opl файл CV_LS: ' + STR(cv_ls.tip_opl,1))
        ENDIF 
        IF cv_ls.ot_d=0
         rval    = InsError('S', 'CYA', m.recid, '', 'Некорректное значение поля ot_d файл CV_LS: ' + TRANSFORM(ot_d,'999999.9999'))
        ENDIF 
        m.n_ru = cv_ls.n_ru
        IF !SEEK(m.n_ru, 'mfc', 'n_ru')
         rval    = InsError('S', 'CTA', m.recid, '', ;
         	'Значение поля n_ru файлa CV_LS: '+ALLTRIM(m.n_ru)+' не найдено в справочнике medicament_mfc')
        ENDIF 
       ENDIF 
      ENDIF 
     ENDIF 
    ENDIF 
   ENDIF 
   
   *MESSAGEBOX(STR(recid,6),0+64,'x')
   IF M.O0A = .T. && M.O0A - отключалка для всей онкологии!
    m.cod       = cod
    m.ds        = ds
    m.ds_2      = ds_2
    m.sn_pol    = sn_pol
    m.recid_s   = recid_lpu
    m.recid_sl  = ''
    m.recid_usl = ''
    m.o_otd     = SUBSTR(otd,2,2)
    m.usl_ok    = IIF(SEEK(m.o_otd, 'profot'), profot.usl_ok, '0')
    m.is_gsp    = IIF(m.usl_ok='1', .T., .F.)
	m.usl_tip   = 0
	m.ds_onk    = ds_onk
	m.d_beg     = people.d_beg
	m.d_end     = people.d_end
    *m.reab    = reab
    
    *m.IsOnkDs = IIF(m.ds_onk=1 OR LEFT(m.ds,1)='C' OR ;
  		(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')), .T., .F.)
    *m.IsOnkDs = IIF((BETWEEN(LEFT(m.ds,3), 'C00', 'C80') OR m.ds='C97') OR ;
    	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) , .T., .F.)
    *m.IsOnkDs = IIF(m.ds_onk=1 OR LEFT(m.ds,1)='C' OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) OR ;
  		BETWEEN(LEFT(m.ds,3),'D00','D09') , .T., .F.)
    
    * со счетов за апрель
    *m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR ;
  	(m.ds='D70' AND (BETWEEN(LEFT(m.ds_2,3), 'C00', 'C80') OR m.ds_2='C97')) OR ;
  		BETWEEN(LEFT(m.ds,3),'D00','D09') , .T., .F.)
    m.IsOnkDs = IIF(LEFT(m.ds,1)='C' OR BETWEEN(LEFT(m.ds,3),'D00','D09') , .T., .F.)
    * со счетов за апрель

    *IF m.IsOnkDs AND m.is_gsp && IsGsp(m.cod)
    IF m.IsOnkDs
     IF !USED('onk_sl')
      m.recid = recid
      rval    = InsError('S', 'O6A', m.recid)
     ELSE 
      IF !SEEK(m.recid_s, 'onk_sl')
       m.recid = recid
       rval    = InsError('S', 'O6A', m.recid)
      ELSE 
       m.recid_sl = onk_sl.recid
       IF sn_pol<>onk_sl.sn_pol OR c_i<>onk_sl.c_i OR cod<>onk_sl.cod
        m.recid = recid
        rval    = InsError('S', 'O6A', m.recid, '',;
        	'Отсутствует запись в onk_sl по ключу recid_lpu+sn_pol+c_i+cod')
       ENDIF 
       IF !SEEK(onk_sl.ds1_t, 'onreas')
        m.recid = recid
        rval    = InsError('S', 'O6A', m.recid)
       ENDIF 
       
       IF IsVMP(m.cod)
        IF !INLIST(onk_sl.ds1_t,0,1,2,6)
         m.recid = recid
         rval    = InsError('S', 'O6A', m.recid)
        ENDIF 
       ENDIF 
       
       IF !IsVMP(m.cod)
        IF INLIST(onk_sl.ds1_t,0,1,2,3,4)
         IF !SEEK(onk_sl.stad, 'onstad')
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ELSE && 5,6
         IF !EMPTY(onk_sl.stad)
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ENDIF 
       ELSE && IF IsVMP(m.cod)
        IF INLIST(onk_sl.ds1_t,0,1,2)
         IF !SEEK(onk_sl.stad, 'onstad')
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ELSE && 5,6
         IF !EMPTY(onk_sl.stad)
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 
       
       IF SEEK(onk_sl.stad, 'onstad')
        m.c_len = LEN(ALLTRIM(onstad.ds))
        IF !EMPTY(onstad.ds)
         IF LEFT(m.ds, m.c_len) != LEFT(onstad.ds, m.c_len)
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ELSE 
         IF SEEK(m.ds, 'onstad', 'ds') OR SEEK(LEFT(m.ds,5)+SPACE(1), 'onstad', 'ds') OR SEEK(LEFT(m.ds,3)+SPACE(3), 'onstad', 'ds')
         *IF SEEK(LEFT(m.ds, m.c_len), 'onstad', 'ds')
          m.recid = recid
          rval    = InsError('S', 'O1A', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 

       IF onk_sl.ds1_t=0 AND (m.tdat1-people.dr)/365.25>=18 AND !SEEK(onk_sl.onk_t, 'ontum')
        m.recid = recid
        rval    = InsError('S', 'OVA', m.recid)
       ELSE 
        IF (onk_sl.ds1_t!=0 OR (m.tdat1-people.dr)/365.25<18) AND !EMPTY(onk_sl.onk_t)
         m.recid = recid
         rval    = InsError('S', 'OVA', m.recid)
        ENDIF 
       ENDIF 
       
       IF SEEK(onk_sl.onk_t, 'ontum')
        m.c_len = LEN(ALLTRIM(ontum.ds))
        IF !EMPTY(ontum.ds)
         IF LEFT(m.ds, m.c_len) != LEFT(ontum.ds, m.c_len)
          m.recid = recid
          rval    = InsError('S', 'OVA', m.recid)
         ENDIF 
        ELSE 
         *IF SEEK(m.ds, 'ontum', 'ds')
         IF SEEK(m.ds, 'ontum', 'ds') OR SEEK(LEFT(m.ds,5)+SPACE(1), 'ontum', 'ds') OR SEEK(LEFT(m.ds,3)+SPACE(3), 'ontum', 'ds')
          m.recid = recid
          rval    = InsError('S', 'OVA', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 
       
       IF onk_sl.ds1_t=0 AND (m.tdat1-people.dr)/365.25>=18 AND !SEEK(onk_sl.onk_n, 'onnod')
        m.recid = recid
        rval    = InsError('S', 'OWA', m.recid)
       ELSE 
        IF (onk_sl.ds1_t!=0 OR (m.tdat1-people.dr)/365.25<18) AND !EMPTY(onk_sl.onk_n)
         m.recid = recid
         rval    = InsError('S', 'OWA', m.recid)
        ENDIF 
       ENDIF 
       
       IF SEEK(onk_sl.onk_n, 'onnod')
        m.c_len = LEN(ALLTRIM(onnod.ds))
        IF !EMPTY(onnod.ds)
         IF LEFT(m.ds, m.c_len) != LEFT(onnod.ds, m.c_len)
          m.recid = recid
          rval    = InsError('S', 'OWA', m.recid)
         ENDIF 
        ELSE 
         *IF SEEK(LEFT(m.ds, m.c_len), 'onnod', 'ds')
         IF SEEK(m.ds, 'onnod', 'ds') OR SEEK(LEFT(m.ds,5)+SPACE(1), 'onnod', 'ds') OR SEEK(LEFT(m.ds,3)+SPACE(3), 'onnod', 'ds')
          m.recid = recid
          rval    = InsError('S', 'OWA', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 
       
       IF onk_sl.ds1_t=0 AND (m.tdat1-people.dr)/365.25>=18 AND !SEEK(onk_sl.onk_m, 'onmet')
        m.recid = recid
        rval    = InsError('S', 'OXA', m.recid)
       ELSE 
        *IF !EMPTY(onk_sl.onk_m)
        IF (onk_sl.ds1_t!=0 OR (m.tdat1-people.dr)/365.25<18) AND !EMPTY(onk_sl.onk_m)
         m.recid = recid
         rval    = InsError('S', 'OXA', m.recid)
        ENDIF 
       ENDIF 
       
       IF SEEK(onk_sl.onk_m, 'onmet')
        m.c_len = LEN(ALLTRIM(onmet.ds))
        IF !EMPTY(onmet.ds)
         IF LEFT(m.ds, m.c_len) != LEFT(onmet.ds, m.c_len)
          m.recid = recid
          rval    = InsError('S', 'OXA', m.recid)
         ENDIF 
        ELSE 
         *IF SEEK(m.ds, 'onmet', 'ds')
         IF SEEK(m.ds, 'onmet', 'ds') OR SEEK(LEFT(m.ds,5)+SPACE(1), 'onmet', 'ds') OR SEEK(LEFT(m.ds,3)+SPACE(3), 'onmet', 'ds')
          m.recid = recid
          rval    = InsError('S', 'OXA', m.recid)
         ENDIF 
        ENDIF 
       ENDIF 

       IF INLIST(onk_sl.ds1_t,1,2)
        IF !INLIST(onk_sl.mtstz,0,1)
         m.recid = recid
         rval    = InsError('S', 'OJA', m.recid)
        ENDIF 
       ELSE 
        IF onk_sl.mtstz!=0
         m.recid = recid
         rval    = InsError('S', 'OJA', m.recid)
        ENDIF 
       ENDIF 
     
       IF USED('onk_usl')
        IF SEEK(m.recid_sl, 'onk_usl')

         IF INLIST(onk_usl.usl_tip,3,4)
          * Значение 0 допустимо, следовательно проверка бессмысленна!
         ELSE 
          IF !EMPTY(onk_sl.sod)
           m.recid = recid
           rval    = InsError('S', 'ONA', m.recid)
          ENDIF 
         ENDIF 

         IF INLIST(onk_usl.usl_tip,3,4)
          * Значение 0 допустимо, следовательно проверка бессмысленна!
         ELSE 
          IF !EMPTY(onk_sl.k_fr)
           m.recid = recid
           rval    = InsError('S', 'OTA', m.recid)
          ENDIF 
         ENDIF 

         IF INLIST(onk_usl.usl_tip,2,4)
          IF EMPTY(onk_sl.wei) OR EMPTY(onk_sl.hei) OR EMPTY(onk_sl.bsa)
           m.recid = recid
           rval    = InsError('S', 'OYA', m.recid)
          ENDIF 
          IF onk_sl.wei>0 AND onk_sl.wei<1.5
           m.recid = recid
           rval    = InsError('S', 'OYA', m.recid, '',;
           	'Поле wei заполнено '+STR(onk_sl.wei,5)+', но более 1.5')
          ENDIF 
          IF onk_sl.hei>0 AND !BETWEEN(onk_sl.hei,40,260)
           m.recid = recid
           rval    = InsError('S', 'OYA', m.recid, '',;
           	'Поле hei заполнено '+STR(onk_sl.hei,5)+', но не в диапазаоне 40-260')
          ENDIF 
         ELSE 
          IF !EMPTY(onk_sl.wei) OR !EMPTY(onk_sl.hei) OR !EMPTY(onk_sl.bsa)
           *m.recid = recid
           *rval    = InsError('S', 'OYA', m.recid)
          ENDIF 
         ENDIF 

        ENDIF 
       ENDIF 

      ENDIF && IF !SEEK(m.recid_s, 'onk_sl')
     ENDIF && IF !USED('onk_sl')
     
     IF USED('onk_sl') AND USED('onk_diag') && проверка файла onk_diag
      IF !EMPTY(m.recid_sl) && m.recid_sl = onk_sl.recid
       IF SEEK(m.recid_sl, 'onk_diag')
        IF (!EMPTY(onk_diag.diag_date) OR !EMPTY(onk_diag.diag_code) OR !EMPTY(onk_diag.rec_rslt)) AND ;
        	!INLIST(onk_diag.diag_tip,1,2)
         m.recid = recid
         rval    = InsError('S', 'OLA', m.recid)
        ENDIF 
        IF (EMPTY(onk_diag.diag_date) AND EMPTY(onk_diag.diag_code) AND EMPTY(onk_diag.rec_rslt)) AND ;
        	onk_diag.diag_tip!=0
         m.recid = recid
         rval    = InsError('S', 'OLA', m.recid)
        ENDIF 
        
        IF onk_diag.diag_tip=1 AND !SEEK(onk_diag.diag_code, 'onmrf')
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=1 AND onk_diag.diag_code=0
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=1 AND !SEEK(onk_diag.diag_code, 'onmrds')
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=2 AND !SEEK(onk_diag.diag_code, 'onigh')
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=2 AND !SEEK(onk_diag.diag_code, 'onigds')
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=2 AND onk_diag.diag_code=0
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF !EMPTY(onk_diag.diag_rslt) AND onk_diag.diag_code=0
         m.recid = recid
         rval    = InsError('S', 'ODA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=1 AND !SEEK(onk_diag.diag_rslt, 'onmrfr')
         m.recid = recid
         rval    = InsError('S', 'OEA', m.recid)
        ENDIF 

        IF onk_diag.diag_tip=2 AND !SEEK(onk_diag.diag_rslt, 'onigrt')
         m.recid = recid
         rval    = InsError('S', 'OEA', m.recid)
        ENDIF 
        
        IF onk_diag.rec_rslt=1 AND onk_diag.diag_code=0
         m.recid = recid
         rval    = InsError('S', 'OEA', m.recid)
        ENDIF 

        IF EMPTY(onk_diag.diag_date) AND ;
        	(!EMPTY(onk_diag.diag_code) OR !EMPTY(onk_diag.diag_tip) OR !EMPTY(onk_diag.rec_rslt))
         m.recid = recid
         rval    = InsError('S', 'OQA', m.recid)
        ELSE 
         IF !EMPTY(onk_diag.diag_date) AND onk_diag.diag_date>d_u
          rval    = InsError('S', 'OQA', m.recid, '', ;
          	'Дата diag_date (onk_diag) '+DTOC(onk_diag.diag_date)+'позже даты связанной с ней услуги (talon) ('+DTOC(d_u)+')')
         ENDIF 
        ENDIF 

        IF onk_diag.rec_rslt!=1 AND !EMPTY(onk_diag.diag_rslt)
         m.recid = recid
         rval    = InsError('S', 'OKA', m.recid)
        ENDIF 
        IF onk_diag.rec_rslt!=0 AND EMPTY(onk_diag.diag_rslt)
         m.recid = recid
         rval    = InsError('S', 'OKA', m.recid)
        ENDIF 
        
        
       ENDIF 
      ENDIF 
     ENDIF 
	 
	 * IF IsGsp(m.cod) OR IsDst(m.cod) && Отключено 21.11.2019 после проверки ВТБ
	 IF USED('onk_sl')
	 IF INLIST(m.usl_ok,'1','2') && AND INLIST(onk_sl.ds1_t,1,2) && Включено 22.11.2019 после проверки ВТБ
	 *IF INLIST(m.usl_ok,'1','2','3') && AND INLIST(onk_sl.ds1_t,1,2) && Включено 29.03.2020 после проверки ВТБ

     IF !USED('onk_usl')
      m.recid = recid
      IF IsVMP(m.cod)
       rval    = InsError('S', 'O8A', m.recid)
      ENDIF 
      IF !IsVMP(m.cod) AND INLIST(onk_sl.ds1_t,0,1,2)
       rval    = InsError('S', 'O8A', m.recid)
      ENDIF 
     ELSE 
      IF EMPTY(m.recid_sl) OR (!EMPTY(m.recid_sl) AND !SEEK(m.recid_sl, 'onk_usl'))
       m.recid = recid
       *rval    = InsError('S', 'O8A', m.recid)
       IF IsVMP(m.cod)
        rval    = InsError('S', 'O8A', m.recid)
       ENDIF 
       IF !IsVMP(m.cod) AND INLIST(onk_sl.ds1_t,0,1,2)
        rval    = InsError('S', 'O8A', m.recid)
       ENDIF 
      ELSE 
       m.recid_usl = onk_usl.recid
       IF !SEEK(onk_usl.usl_tip, 'onlech')
        m.recid = recid
        rval    = InsError('S', 'O8A', m.recid)
       ELSE 
	    m.usl_tip = onk_usl.usl_tip
       ENDIF 
       
       IF onk_usl.usl_tip!=1 AND !EMPTY(onk_usl.hir_tip)
        m.recid = recid
        rval    = InsError('S', 'O9A', m.recid)
       ENDIF 
       IF onk_usl.usl_tip=1 AND EMPTY(onk_usl.hir_tip)
        m.recid = recid
        rval    = InsError('S', 'O9A', m.recid)
       ENDIF 

       IF onk_usl.usl_tip=1 && проверку по ho - должна быть операция
        IF !IsVMP(m.cod)
         IF !USED('ho')
          m.recid = recid
          rval = InsError('S', 'O8A', m.recid, '',;
        	'МЭС применен без оперативного пособия (отсутствует файл ho)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ELSE 
          m.cod = cod
          m.c_i    = c_i
          m.sn_pol = sn_pol
          m.vir = m.sn_pol + m.c_i + PADL(m.cod,6,'0')
          IF !SEEK(m.vir, 'ho')
           m.recid = recid
           rval = InsError('S', 'O8A', m.recid, '',;
        	'МЭС применен без оперативного пособия ;
        	(отсутствует соответствующая запись в файле ho)')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ENDIF 
         ENDIF 
        ENDIF 
       ENDIF 

       IF !EMPTY(onk_usl.hir_tip) AND !SEEK(onk_usl.hir_tip, 'onhir')
        m.recid = recid
        rval    = InsError('S', 'O9A', m.recid)
       ENDIF 
       
       IF onk_usl.usl_tip!=2 AND !EMPTY(onk_usl.lek_tip_l)
        m.recid = recid
        rval    = InsError('S', 'OAA', m.recid)
       ENDIF 
       IF onk_usl.usl_tip=2 AND EMPTY(onk_usl.lek_tip_l)
        m.recid = recid
        rval    = InsError('S', 'OAA', m.recid)
       ENDIF 
       IF !EMPTY(onk_usl.lek_tip_l) AND !SEEK(onk_usl.lek_tip_l, 'onlekl')
        m.recid = recid
        rval    = InsError('S', 'OAA', m.recid)
       ENDIF 

       IF onk_usl.usl_tip!=2 AND !EMPTY(onk_usl.lek_tip_v)
        m.recid = recid
        rval    = InsError('S', 'OBA', m.recid)
       ENDIF 
       IF onk_usl.usl_tip=2 AND EMPTY(onk_usl.lek_tip_v)
        m.recid = recid
        rval    = InsError('S', 'OBA', m.recid)
       ENDIF 
       IF !EMPTY(onk_usl.lek_tip_v) AND !SEEK(onk_usl.lek_tip_v, 'onlekv')
        m.recid = recid
        rval    = InsError('S', 'OBA', m.recid)
       ENDIF 
       
       IF INLIST(onk_usl.usl_tip,3,4)
        IF EMPTY(onk_usl.luch_tip)
         m.recid = recid
         rval    = InsError('S', 'OCA', m.recid)
        ELSE 
         IF !SEEK(onk_usl.luch_tip, 'onluch')
          m.recid = recid
          rval    = InsError('S', 'OCA', m.recid)
         ENDIF 
        ENDIF 
       ELSE 
        IF !EMPTY(onk_usl.luch_tip)
         m.recid = recid
         rval    = InsError('S', 'OCA', m.recid)
        ENDIF 
       ENDIF 

       IF INLIST(onk_usl.usl_tip,2,4) AND !INLIST(onk_usl.pptr,0,1)
        m.recid = recid
        rval    = InsError('S', 'OZA', m.recid)
       ENDIF 
       IF !INLIST(onk_usl.usl_tip,2,4) AND onk_usl.pptr!=0
        m.recid = recid
        rval    = InsError('S', 'OZA', m.recid)
       ENDIF 

      ENDIF && IF !SEEK(m.recid_s, 'onk_usl')
     ENDIF && IF !USED('onk_usl')
     
     ELSE 

     IF !USED('onk_usl')
     ELSE 
      IF EMPTY(m.recid_sl) OR (!EMPTY(m.recid_sl) AND !SEEK(m.recid_sl, 'onk_usl'))
      ELSE 

       m.recid_usl = onk_usl.recid

       IF onk_usl.usl_tip=1 && проверку по ho - должна быть операция
        IF !IsVMP(m.cod)
         IF !USED('ho')
          m.recid = recid
          rval = InsError('S', 'O8A', m.recid, '',;
        	' onk_usl.usl_tip=1 без оперативного пособия (отсутствует файл ho)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ELSE 
          m.cod = cod
          m.c_i    = c_i
          m.sn_pol = sn_pol
          m.vir = m.sn_pol + m.c_i + PADL(m.cod,6,'0')
          IF !SEEK(m.vir, 'ho')
           m.recid = recid
           rval = InsError('S', 'O8A', m.recid, '',;
        	' onk_usl.usl_tip=1 без оперативного пособия (отсутствует соответствующая запись в файле ho)')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ENDIF 
         ENDIF 
        ENDIF 
       ENDIF 
      
       IF INLIST(onk_usl.usl_tip,2,4) AND !INLIST(onk_usl.pptr,0,1)
        m.recid = recid
        rval    = InsError('S', 'OZA', m.recid)
       ENDIF 
       IF !INLIST(onk_usl.usl_tip,2,4) AND onk_usl.pptr!=0
        m.recid = recid
        rval    = InsError('S', 'OZA', m.recid)
       ENDIF 

      ENDIF && IF !SEEK(m.recid_s, 'onk_usl')
     ENDIF && IF !USED('onk_usl')

     ENDIF && INLIST(m.usl_ok,1,2) Включено 22.11.2019 после проверки ВТБ
     ENDIF && IF USED('onk_sl')

     IF INLIST(m.usl_tip,2,4)
      IF !USED('onk_ls')
       m.recid = recid
       rval    = InsError('S', 'OSA', m.recid, '',;
       	'Файл onk_ls отсутсвует!')
      ELSE 
       IF EMPTY(m.recid_usl) OR (!EMPTY(m.recid_usl) AND !SEEK(m.recid_usl, 'onk_ls'))
        m.recid = recid
        rval    = InsError('S', 'OSA', m.recid, '',;
        	'Отсутствует соответствующая запись в файле onl_ls!')
       ELSE 
        ** Здесь алгоритмы X1 - X8 - проверка файла onk_ls
        DO WHILE onk_ls.recid_usl = m.recid_usl
         IF FIELD('tip_opl', 'onk_ls')='TIP_OPL'
          *IF onk_ls.tip_opl=1
          m.ds        = ds
          m.dd_sid    = onk_ls.sid
          m.gd_sid    = IIF(SEEK(m.dd_sid, 'medx'), ALLTRIM(medx.gd_sid), '')
          m.max_dose  = IIF(SEEK(m.dd_sid, 'medx'),medx.max_dose, 0)
          m.is_tarion = IIF(!EMPTY(m.gd_sid) AND SEEK(m.gd_sid, 'tarion', 'cod'), .T., .F.)
          
          m.dr = people.dr
          m.adj  = CTOD(STRTRAN(DTOC(m.dr), STR(YEAR(m.dr),4), STR(YEAR(people.d_beg),4)))-people.d_beg
          *m.vozr   = (YEAR(m.d_u) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)
          m.vozr   = (YEAR(people.d_beg) - YEAR(m.dr)) - IIF(m.adj>0, 1, 0)

          *IF INLIST(onk_ls.cod,97158,81094) AND !(m.vozr<18 OR BETWEEN(m.ds,'C81.0','C96.9'))
          IF !(m.vozr<18 OR BETWEEN(m.ds,'C81.0','C96.9'))
           IF EMPTY(onk_ls.regnum)
            m.recid = recid
            rval    = InsError('S', 'OSA', m.recid)
           ELSE 
            IF SEEK(onk_ls.code_sh, 'onlpsh')
             IF !SEEK(onk_ls.code_sh+LEFT(onk_ls.regnum,6), 'onlpsh', 'unik')
              m.recid = recid
              rval    = InsError('S', 'OSA', m.recid)
             ENDIF 
            ENDIF 
           ENDIF 
          ENDIF 

          IF onk_ls.tip_opl=1 AND m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND EMPTY(onk_ls.n_par)
           m.recid = recid
           rval    = InsError('S', 'X1A', m.recid)
          ENDIF 

          IF onk_ls.tip_opl=1 AND m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND EMPTY(onk_ls.r_up)
           m.recid = recid
           rval    = InsError('S', 'X2A', m.recid, '',;
           	'Пустое r_up')
          ENDIF 

          IF onk_ls.tip_opl=1 AND m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND !EMPTY(onk_ls.r_up) ;
          	AND !SEEK(LEFT(onk_ls.r_up,10), 'medpack')
           m.recid = recid
           rval    = InsError('S', 'X2A', m.recid, '',;
           	'r_up не найден в medpack: '+onk_ls.r_up)
          ENDIF 

          *IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND onk_ls.tip_opl=0
          IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND onk_ls.tip_opl=0
           m.recid = recid
           rval    = InsError('S', 'X3A', m.recid)
          ENDIF 

          *IF m.is_tarion AND INLIST(onk_ls.cod,97158,81094) AND !INLIST(onk_ls.tip_opl,0,1,2,3,4,5)
          IF m.is_tarion AND !INLIST(onk_ls.tip_opl,0,1,2,3,4,5)
           m.recid = recid
           rval    = InsError('S', 'X3A', m.recid)
          ENDIF 

          *IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND EMPTY(onk_ls.n_ru)
          IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND EMPTY(onk_ls.n_ru)
           m.recid = recid
           rval    = InsError('S', 'X4A', m.recid)
          ENDIF 
          
          IF USED('mfc')
           *IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND EMPTY(onk_ls.n_ru) AND ;
           	(SEEK(onk_ls.sid, 'mfc') AND mfc.n_ru!=onk_ls.n_ru)
           IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND EMPTY(onk_ls.n_ru) AND ;
           	(SEEK(onk_ls.sid, 'mfc') AND mfc.n_ru!=onk_ls.n_ru)
            m.recid = recid
            rval    = InsError('S', 'X4A', m.recid)
           ENDIF 
          ENDIF 

          IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND onk_ls.ot_d<=0
           m.recid = recid
           rval    = InsError('S', 'X5A', m.recid, '',;
           	'Поле ot_d не заполнено при заполненном regnum')
          ENDIF 

          IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND onk_ls.ot_d>m.max_dose
           m.recid = recid
           rval    = InsError('S', 'X5A', m.recid, '',;
           	'Превышение разовой дозы: '+TRANSFORM(onk_ls.ot_d,'99999.9999')+;
           	' (максимально допустимая '+TRANSFORM(m.max_dose,'99999.9999'))
          ENDIF 

          IF onk_ls.tip_opl=1 AND m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND EMPTY(onk_ls.dt_q)
           m.recid = recid
           rval    = InsError('S', 'X6A', m.recid)
          ENDIF 

          *IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND (EMPTY(onk_ls.dt_d) OR ;
          	onk_ls.dt_d != onk_ls.ot_d * onk_ls.dt_q)
          IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND (EMPTY(onk_ls.dt_d) OR ;
          	onk_ls.dt_d != onk_ls.ot_d * onk_ls.dt_q)
           m.recid = recid
           rval    = InsError('S', 'X7A', m.recid)
          ENDIF 

          *IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND EMPTY(onk_ls.sid)
          IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND EMPTY(onk_ls.sid)
           m.recid = recid
           rval    = InsError('S', 'X8A', m.recid)
          ENDIF 
		  
          *IF m.is_tarion AND INLIST(onk_ls.cod,97158,81094) AND (!EMPTY(onk_ls.regnum) AND EMPTY(onk_ls.date_inj))
          *IF m.is_tarion AND (!EMPTY(onk_ls.regnum) AND EMPTY(onk_ls.date_inj))
          *IF (!EMPTY(onk_ls.regnum) AND EMPTY(onk_ls.date_inj))
          * m.recid = recid
          * rval    = InsError('S', 'OPA', m.recid)
          *ENDIF 
		  
          *IF m.is_tarion
           m.recid = recid
           DO CASE 
             CASE EMPTY(onk_ls.regnum) AND !EMPTY(onk_ls.date_inj)
              rval    = InsError('S', 'OPA', m.recid, '',;
              	'Заполенное поле date_inj при пустом поле regnum')

             CASE !EMPTY(onk_ls.regnum) AND EMPTY(onk_ls.date_inj)
              rval    = InsError('S', 'OPA', m.recid, '',;
              	'Пустое поле date_inj при заполненном поле regnum')
              	
             CASE (IsMes(m.cod) OR IsVMP(m.cod)) && AND !BETWEEN(onk_ls.date_inj, m.d_beg, m.d_end)
              m.c_i = c_i
              m.d_1 = m.d_beg
              m.d_2 = m.d_beg
     		  IF USED('hosp')
               m.d_1 = IIF(SEEK(m.c_i, 'hosp'), hosp.d_pos-1, m.d_1)
               m.d_2 = IIF(SEEK(m.c_i, 'hosp'), hosp.d_vip, m.d_2)
              ENDIF 
              IF !BETWEEN(onk_ls.date_inj, m.d_1, m.d_2)
               rval    = InsError('S', 'OPA', m.recid, '',;
              	'Дата date_inj выходит из интервала '+DTOC(m.d_1)+'-'+DTOC(m.d_2))
              ENDIF 
              
              OTHERWISE 
             
           ENDCASE 
          *ENDIF 
		  
          *IF m.is_tarion
           m.recid = recid
           DO CASE  
           	CASE INLIST(onk_usl.usl_tip,2,4) AND !EMPTY(onk_ls.code_sh) AND !SEEK(onk_ls.code_sh, 'ondopk') 
           	 IF !((m.vozr<18 OR BETWEEN(m.ds,'C81.0','C96.9')) AND LOWER(onk_ls.code_sh)='нет')
              rval    = InsError('S', 'ORA', m.recid, '',;
              	'Значение поля code_sh '+onk_ls.code_sh+' не найдено в справочнике ondopkxx.dbf')
           	 ENDIF 
           	 
           	 CASE onk_usl.usl_tip=2
           	  IF LEFT(onk_ls.code_sh,2)<>'sh' AND LOWER(onk_ls.code_sh)<>'нет'
               rval    = InsError('S', 'ORA', m.recid, '',;
              	'Значение поля code_sh '+onk_ls.code_sh+' не найдено в справочнике ondopkxx.dbf')
           	  ENDIF 
           	 
           	 CASE onk_usl.usl_tip=4
           	  IF !INLIST(LEFT(onk_ls.code_sh,2),'sh','mt') AND LOWER(onk_ls.code_sh)<>'нет'
               rval    = InsError('S', 'ORA', m.recid, '',;
              	'Значение поля code_sh '+onk_ls.code_sh+' не найдено в справочнике ondopkxx.dbf')
           	  ENDIF 
           
             *CASE INLIST(onk_usl.usl_tip,2,4) AND !EMPTY(onk_ls.code_sh) AND ;
            * 	!(SEEK(onk_ls.code_sh, 'ondopk') OR LOWER(onk_ls.code_sh)='нет')
            *  rval    = InsError('S', 'ORA', m.recid, '',;
            *  	'Значение поля code_sh '+onk_ls.code_sh+' не найдено в справочнике ondopkxx.dbf')
              OTHERWISE 
             
           ENDCASE 
          *ENDIF 
		  
		  IF USED('onopls') AND X9A=.T.
           *IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND INLIST(onk_ls.cod,97158,81094) AND !EMPTY(onk_ls.sid) AND onk_ls.s_all>0
           IF m.is_tarion AND !EMPTY(onk_ls.regnum) AND !EMPTY(onk_ls.sid) AND onk_ls.s_all>0
		    m.is_target = IIF(SEEK(onk_ls.sid, 'medx'), medx.is_target, 0)
		    IF m.is_target = 1 
		     m.vir = PADL(m.cod,6,'0')+LEFT(onk_ls.regnum,6)+LEFT(m.ds,3)
		     IF SEEK(LEFT(m.ds,3),'onopls','ds') AND SEEK(LEFT(onk_ls.regnum,6), 'onopls', 'regnum')
		      IF !SEEK(m.vir, 'onopls')
               m.recid = recid
               rval    = InsError('S', 'X9A', m.recid)
              ENDIF 
             ENDIF 
            ENDIF 
           ENDIF 
          ENDIF 

         ENDIF 
         *ENDIF 

         SKIP IN onk_ls
        ENDDO 
        ** Здесь алгоритмы X1 - X8 - проверка файла onk_ls
       ENDIF 
       
      ENDIF && IF !USED('onk_ls')
     ENDIF && IF INLIST(m.usl_tip,2,4)
     
     *ENDIF && IsGsp or IsDst  && Отключено 21.11.2019 после проверки ВТБ

    ELSE  
     * сюда всетавить проверку наличия записи, если она быть не должна!
    ENDIF && IF m.IsOnkDs

    IF m.IsOnkDs OR m.ds_onk=1    
     IF !USED('onk_cons')
      m.recid = recid
      rval    = InsError('S', 'O7A', m.recid, '',;
      	'Отсутствует файл onk_cons!')
      m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
     ELSE 
      IF SEEK(m.recid_s, 'onk_cons')
       m.pr_cons = onk_cons.pr_cons
       m.dt_cons = onk_cons.dt_cons

       IF sn_pol<>onk_cons.sn_pol OR c_i<>onk_cons.c_i OR cod<>onk_cons.cod
        m.recid = recid
        rval    = InsError('S', 'O7A', m.recid, '',;
        	'Файл onk_cons не связывается по ключу sn_pol+c_i+cod')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 

       IF !SEEK(m.pr_cons, 'oncons')
        m.recid = recid
        rval    = InsError('S', 'O7A', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
       IF INLIST(m.pr_cons,1,2,3) AND EMPTY(m.dt_cons)
        m.recid = recid
        rval    = InsError('S', 'OOA', m.recid)
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ENDIF 
      ELSE 
       m.recid = recid
       rval    = InsError('S', 'O7A', m.recid)
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ELSE 
     IF USED('onk_cons')
      IF SEEK(m.recid_s, 'onk_cons')
       m.recid = recid
       rval    = InsError('S', 'O7A', m.recid, '',;
       	'Значение поля PR_CONS не равно "пусто" в иных случаях')
       m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
      ENDIF 
     ENDIF 
    ENDIF 

   ENDIF && IF M.O0A = .T.


   IF M.PPA == .T. AND m.qcod='I3' && перенес из ss_flk! вернул в ss_flk!
    
    *m.ks_plan = IIF(SEEK(m.lpuid, 'nsif'), nsif.ks, 0)
    
    IF USED('gr_plan') AND USED('nsif')

    m.otd = otd 
    m.cod = cod
    m.usl_ok = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), profot.usl_ok, ' ')
    m.tip = tip
    m.ds  = ds
    m.ord = ord
    m.k_u = k_u
    m.c_i = c_i

    m.sn_pol = sn_pol
    m.d_u    = d_u
    m.dr     = IIF(SEEK(m.sn_pol, 'people'), people.dr, {})
    m.vozr   = (YEAR(m.d_u-m.k_u) - YEAR(m.dr))
    m.s_all = s_all+s_lek
    
    m.recid = recid
    
    IF SEEK(m.recid, 'sError')
    ELSE 

    DO CASE 
     CASE m.usl_ok = '1' && стационар
      * алгоритм реализован в oneflk
      
     CASE m.usl_ok = '2' && дневной стационар
       *MESSAGEBOX(TRANSFORM(m.on_ds,'999'), 0+64, m.mcod)
      m.gr = IIF(SEEK(m.cod,'gr_plan'), gr_plan.gr_plan, '')
      *IF !INLIST(m.gr, 'gem', 'eco')
      IF !INLIST(m.gr, 'eco')
       m.ods_fact = nsif.ds_fact
       m.nds_fact = m.ods_fact + m.s_all
       *REPLACE ds_fact WITH m.nds_fact IN nsif
       
       *IF !SEEK(m.sn_pol, 'p_dst')
       * INSERT INTO p_dst FROM MEMVAR 
       * m.on_ds = nsif.n_ds
       * m.nn_ds = m.on_ds + 1
       *ELSE 
       * m.nn_ds = nsif.n_ds
       *ENDIF 
       DO CASE 
        CASE !SEEK(m.lpuid, 'nsif')
         rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'МО не найдено в справочнике nsif (дневной стационар)')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        
        CASE m.nds_fact>nsif.ds+nsif.ds_gem
        *CASE m.nds_fact>nsif.ds
         IF !m.IsPilot
          rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'Превышен лимит ds (дневной стационар, деньги)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 

        OTHERWISE 
         REPLACE ds_fact WITH m.nds_fact IN nsif && experiment!
         * OK!
       ENDCASE 
      ENDIF 
      
      IF m.gr = 'gem'
       m.ogem_fact = nsif.gem_fact
       m.ngem_fact = m.ogem_fact + m.s_all
       IF !SEEK(m.recid, 'sError')
        REPLACE gem_fact WITH m.ngem_fact IN nsif
       ENDIF 
       
       *DO CASE 
       * CASE !SEEK(m.lpuid, 'nsif')
       *  m.recid = recid
       *  rval    = InsError('S', 'PPA', m.recid, '',;
       *	 	'МО не найдено в справочнике nsif (дневной стационар gem)')
       *  m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       * 
       * CASE m.ngem_fact>nsif.ds_gem
       *  m.recid = recid
       *  rval    = InsError('S', 'PPA', m.recid, '',;
       *	 	'Превышен лимит ds (дневной стационар gem)')
       *  m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       * OTHERWISE 
         * OK!
         
       *ENDCASE 
      ENDIF 

      IF m.gr = 'eco'
       m.oeco_fact = nsif.eco_fact
       m.neco_fact = m.oeco_fact + m.s_all
       *REPLACE eco_fact WITH m.neco_fact IN nsif
       
       DO CASE 
        CASE !SEEK(m.lpuid, 'nsif')
         rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'МО не найдено в справочнике nsif (дневной стационар eco)')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        
        CASE m.neco_fact>nsif.ds_eco
         IF !m.IsPilot
          rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'Превышен лимит ds (дневной стационар eco)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 
        OTHERWISE 
         REPLACE eco_fact WITH m.neco_fact IN nsif
         * OK!
       ENDCASE 
      ENDIF 

     CASE m.usl_ok = '3' && АПП

      IF (IsUsl(m.cod) AND SEEK(m.c_i, 'c_pp')) OR (IsUsl(m.cod) AND (USED('hosp') AND SEEK(m.c_i, 'hosp')))

       m.oks_fact = nsif.ks_fact
       m.nks_fact = m.oks_fact + m.s_all
       REPLACE ks_fact WITH m.nks_fact IN nsif
       IF (IsUsl(m.cod) AND SEEK(m.c_i, 'c_pp'))
        rval    = InsError('S', 'PPA', m.recid, '',;
       	 'Превышен лимит ks (12)!')
        m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
       ELSE 
        IF m.nks_fact>nsif.ks
         rval    = InsError('S', 'PPA', m.recid, '',;
       	  'Превышен лимит ks (13)!')
         m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        ENDIF 
       ENDIF 

      ELSE 
      
       IF !(IsPilot OR m.IsPilots)

       m.gr = IIF(SEEK(m.cod,'gr_plan'), gr_plan.gr_plan, '')
       IF !INLIST(m.gr, 'kt')
        IF INLIST(INT(m.cod/1000),59,159) AND nsif.app<=0
         m.oks_fact = nsif.ks_fact
         m.nks_fact = m.oks_fact + m.s_all
      
         IF !(m.nks_fact<=nsif.ks)
          rval    = InsError('S', 'PPA', m.recid, '', 'Превышен лимит ks (59-ые коды)!')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         ENDIF 
       
        ELSE 
         m.oapp_fact = nsif.app_fact
         m.napp_fact = m.oapp_fact + m.s_all
         REPLACE app_fact WITH m.napp_fact IN nsif
       
         DO CASE 
          CASE !SEEK(m.lpuid, 'nsif')
           rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'МО не найдено в справочнике nsif (АПП)')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        
          CASE m.napp_fact>nsif.app
           IF !m.IsPilot
            rval    = InsError('S', 'PPA', m.recid, '',;
       	 	 'Превышен лимит app (АПП)')
            m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
           ENDIF 
          OTHERWISE 
          * OK!
         ENDCASE 
        ENDIF 
       ELSE 
        m.optkt_fact = nsif.ptkt_fact
        m.nptkt_fact = m.optkt_fact + m.s_all
        *REPLACE app_ptkt WITH m.nptkt_fact IN nsif
        REPLACE ptkt_fact WITH m.nptkt_fact IN nsif
       
        DO CASE 
         CASE !SEEK(m.lpuid, 'nsif')
          rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'МО не найдено в справочнике nsif (ПЭТ/КТ)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
        
         CASE m.nptkt_fact>nsif.app_ptkt
          IF !m.IsPilot
           rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'Превышен лимит app (ПЭТ/КТ)')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          ENDIF 

         OTHERWISE 
          * OK!
         
        ENDCASE 
        
        ENDIF 

       ENDIF 
      ENDIF 

     OTHERWISE
    ENDCASE 
    
    ENDIF && IF SEEK(m.recid, 'sError')
    ENDIF && IF USED('prvr2') AND USED('nsio')
   ENDIF && IF M.PPA == .T.

   IF M.PPA==.T. AND m.qcod='S7' && перенес из ss_flk! вернул в ss_flk!
    
    IF USED('gr_plan') AND USED('nsif')
     m.otd = otd 
     m.cod = cod
     m.usl_ok = IIF(SEEK(SUBSTR(m.otd,2,2), 'profot'), profot.usl_ok, ' ')
     m.tip = tip
     m.ds  = ds
     m.ord = ord
     m.k_u = k_u
     m.c_i = c_i

     m.sn_pol = sn_pol
     m.d_u    = d_u
     m.dr     = IIF(SEEK(m.sn_pol, 'people'), people.dr, {})
     m.vozr   = (YEAR(m.d_u-m.k_u) - YEAR(m.dr))
     m.s_all = s_all+s_lek
    
     m.recid = recid
    
     IF SEEK(m.recid, 'sError')
     ELSE 
      =SEEK(m.lpuid, 'nsif')
      DO CASE 
       CASE IsMes(m.cod) OR INLIST(FLOOR(m.cod/1000),200,300,397) OR INLIST(m.cod,56029,156003)&& КС
        m.on_ks = nsif.n_ks
        m.nn_ks = m.on_ks
        m.oks_fact = nsif.ks_fact
        m.nks_fact = m.oks_fact + m.s_all

        IF !SEEK(m.c_i, 'n_ks')
         m.nn_ks = m.on_ks + 1
        ENDIF 
        m.nsif = 2

        DO CASE 
         CASE !SEEK(m.lpuid, 'nsif')
          rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'МО не найдено в справочнике nsif (дневной стационар)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
         CASE m.nn_ks>nsif.n_ks_plan
          rval    = InsError('S', 'PPA', m.recid, '',;
      		'Превышен лимит ks (круглосуточный стационар, случаи)')
          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

         OTHERWISE 
          * OK!
          IF !SEEK(m.c_i, 'n_ks')
           INSERT INTO n_ks FROM MEMVAR 
          ENDIF 
          REPLACE ks_fact WITH m.nks_fact, n_ks WITH m.nn_ks IN nsif
         * OK!
        ENDCASE 
    
       CASE INLIST(FLOOR(m.cod/1000), 97,197,297) && ДС
        m.gr = IIF(SEEK(m.cod,'gr_plan'), gr_plan.gr_plan, '')
        IF !INLIST(m.gr, 'eco')
         m.on_ds = nsif.n_ds
         m.nn_ds = m.on_ds
         m.ods_fact = nsif.ds_fact
         m.nds_fact = m.ods_fact + m.s_all

         IF !SEEK(m.c_i, 'n_ds')
          m.nn_ds = m.on_ds + 1
         ENDIF 
         m.nsif = 2

         DO CASE 
          CASE !SEEK(m.lpuid, 'nsif')
           rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'МО не найдено в справочнике nsif (дневной стационар)')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

          CASE m.nn_ds>nsif.n_ds_plan
           *IF !m.IsPilot
            rval    = InsError('S', 'PPA', m.recid, '',;
      	 	 'Превышен лимит ds (дневной стационар, случаи)')
            m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
           *ENDIF 

          OTHERWISE 
           * OK!
           IF !SEEK(m.c_i, 'n_ds')
            INSERT INTO n_ds FROM MEMVAR 
           ENDIF 
           REPLACE ds_fact WITH m.nds_fact, n_ds WITH m.nn_ds IN nsif

           IF m.gr = 'gem'
	        m.on_gem = nsif.n_gem
	        m.nn_gem = m.on_gem
	        m.ogem_fact = nsif.gem_fact
	        m.ngem_fact = m.ogem_fact + m.s_all
            IF !SEEK(m.c_i, 'n_gem')
             INSERT INTO n_gem FROM MEMVAR 
             m.nn_gem = m.on_gem +1 
            ENDIF 
            REPLACE gem_fact WITH m.ngem_fact, n_gem WITH m.nn_gem IN nsif
            m.nsif = 4
          ENDIF 
          * OK!
         ENDCASE 
        ENDIF 

        IF m.gr = 'eco'
         m.on_eco = nsif.n_eco
         m.nn_eco = m.on_eco
         m.oeco_fact = nsif.eco_fact
         m.neco_fact = m.oeco_fact + m.s_all

         IF !SEEK(m.c_i, 'n_eco')
          m.nn_eco = m.on_eco + 1
         ENDIF 
         m.nsif = 2

         DO CASE 
          CASE !SEEK(m.lpuid, 'nsif')
           rval    = InsError('S', 'PPA', m.recid, '',;
       	 	'МО не найдено в справочнике nsif (дневной стационар)')
           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
          CASE m.nn_eco>nsif.n_eco_plan
           *IF !m.IsPilot
            rval    = InsError('S', 'PPA', m.recid, '',;
      	 	 'Превышен лимит ЭКО (случаи)')
            m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
           *ENDIF 

          OTHERWISE 
           * OK!
           IF !SEEK(m.c_i, 'n_eco')
            INSERT INTO n_eco FROM MEMVAR 
           ENDIF 
           REPLACE eco_fact WITH m.neco_fact, n_eco WITH m.nn_eco IN nsif
           * OK!
         ENDCASE 

        ENDIF 

	    OTHERWISE && АПП
	     DO CASE 
	      *CASE  (IsUsl(m.cod) AND SEEK(m.c_i, 'c_pp')) OR (IsUsl(m.cod) AND (USED('hosp') AND SEEK(m.c_i, 'hosp'))) ;
	      *	OR (IsUsl(m.cod) AND m.usl_ok='1') && INLIST(SUBSTR(m.otd,2,2),'70','73')

	      * m.on_ks = nsif.n_ks
	      * m.nn_ks = m.on_ks
	      * m.oks_fact = nsif.ks_fact
	      * m.nks_fact = m.oks_fact + m.s_all

	      * IF !SEEK(m.c_i, 'n_ks')
	      *  m.nn_ks = m.on_ks + 1
	      * ENDIF 
	      * m.nsif = 2

	      * DO CASE 
	      *  CASE !SEEK(m.lpuid, 'nsif')
	      *   rval    = InsError('S', 'PPA', m.recid, '',;
	      *  	'МО не найдено в справочнике nsif (круглосуточный стационар стационар)')
	      *   m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	      *  CASE m.nn_ks>nsif.n_ks_plan
	      *   rval    = InsError('S', 'PPA', m.recid, '',;
	      *	'Превышен лимит ks (круглосуточный стационар, случаи)')
	      *   m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

	      *  OTHERWISE 
	      *   * OK!
	      *   IF !SEEK(m.c_i, 'n_ks')
	      *    INSERT INTO n_ks FROM MEMVAR 
	      *   ENDIF 
	      *   REPLACE ks_fact WITH m.nks_fact, n_ks WITH m.nn_ks IN nsif
	      *  * OK!
	      * ENDCASE 

	      * *m.oks_fact = nsif.ks_fact
	      * *m.nks_fact = m.oks_fact + m.s_all
	      * *REPLACE ks_fact WITH m.nks_fact IN nsif

	      CASE  IsUsl(m.cod) AND m.usl_ok='2' 
           m.on_ds = nsif.n_ds
           m.nn_ds = m.on_ds
           m.ods_fact = nsif.ds_fact
           m.nds_fact = m.ods_fact + m.s_all

           IF !SEEK(m.c_i, 'n_ds')
            m.nn_ds = m.on_ds + 1
           ENDIF 
           m.nsif = 2

           DO CASE 
            CASE !SEEK(m.lpuid, 'nsif')
             rval    = InsError('S', 'PPA', m.recid, '',;
       	 	  'МО не найдено в справочнике nsif (дневной стационар)')
             m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)

            CASE m.nn_ds>nsif.n_ds_plan
             *IF !m.IsPilot
              rval    = InsError('S', 'PPA', m.recid, '',;
      	 	   'Превышен лимит ds (дневной стационар, случаи)')
              m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
             *ENDIF 

            OTHERWISE 
             IF !SEEK(m.c_i, 'n_ds')
              INSERT INTO n_ds FROM MEMVAR 
             ENDIF 
             REPLACE ds_fact WITH m.nds_fact, n_ds WITH m.nn_ds IN nsif
           ENDCASE 

	      OTHERWISE 
	       *IF !(IsPilot OR m.IsPilots)
	       m.gr = IIF(SEEK(m.cod,'gr_plan'), gr_plan.gr_plan, '')
	       IF !INLIST(m.gr, 'kt')
	        IF INLIST(INT(m.cod/1000),59,159) AND nsif.app<=0
	         m.oks_fact = nsif.ks_fact
	         m.nks_fact = m.oks_fact + m.s_all
	        ELSE 
	         m.oapp_fact = nsif.app_fact
	         m.napp_fact = m.oapp_fact + m.s_all
	         REPLACE app_fact WITH m.napp_fact IN nsif
	        ENDIF 
	       ELSE 
	        m.optkt_fact = nsif.ptkt_fact
	        m.nptkt_fact = m.optkt_fact + m.s_all
	        m.on_kt = nsif.n_kt
	        m.nn_kt = m.on_kt + m.k_u
	       
	        DO CASE 
	         CASE !SEEK(m.lpuid, 'nsif')
	          rval    = InsError('S', 'PPA', m.recid, '',;
	       	 	'МО не найдено в справочнике nsif (дневной стационар)')
	          m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	         *CASE m.nptkt_fact>nsif.app_ptkt
	         * IF !m.IsPilot
	         *  rval    = InsError('S', 'PPA', m.recid, '',;
	       	* 	'Превышен лимит app (ПЭТ/КТ)')
	         *  m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	         * ENDIF 
	         CASE m.nn_kt>nsif.n_kt_plan
	          *IF !m.IsPilot
	           rval    = InsError('S', 'PPA', m.recid, '',;
	       	 	'Превышен лимит app (ПЭТ/КТ, случаи)')
	           m.s_flk = m.s_flk + IIF(rval==.T., s_all, 0)
	          *ENDIF 
	         OTHERWISE 
	          REPLACE n_kt WITH m.nn_kt IN nsif
	          * OK!
	        ENDCASE 
	        ENDIF 
	       *ENDIF 
	      ENDCASE 

	   ENDCASE 
    
     ENDIF && IF SEEK(m.recid, 'sError')
    ENDIF && IF USED('prvr2') AND USED('nsio')
   ENDIF && IF M.PPA == .T.

RETURN 

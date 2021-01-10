PROCEDURE Parse79
 IF MESSAGEBOX('’Œ“»“≈ «¿√–”«»“‹ XML-—◊≈“?',4+32,'')=7
  RETURN .F.
 ENDIF 
 
 m.lIsSilent = .F.

 SET DEFAULT TO (pSoap)
 csprfile = ''
 zipfile = GETFILE('zip')
 IF EMPTY(zipfile)
  IF  !m.lIsSilent
   MESSAGEBOX(CHR(13)+CHR(10)+'¬€ Õ»◊≈√Œ Õ≈ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN .F.
 ENDIF 
 
 ffile = fso.GetFile(zipfile)
 IF ffile.size >= 2
  fhandl = ffile.OpenAsTextStream
  lcHead = fhandl.Read(2)
  fhandl.Close
 ELSE 
  lcHead = ''
 ENDIF 

 IF lcHead == 'PK' && ›ÚÓ zip-Ù‡ÈÎ!
  IF !UnzipOpen(zipfile)
   IF  !m.lIsSilent
    MESSAGEBOX('›“Œ Õ≈ ZIP-¿–’»¬!',0+64,'')
   ENDIF 
   RETURN .F.
  ENDIF 
  UnzipClose()
 ELSE 
  IF  !m.lIsSilent
   MESSAGEBOX('›“Œ Õ≈ ZIP-¿–’»¬!',0+64,'')
  ENDIF 
  RETURN .F.
 ENDIF 
 
 CREATE CURSOR curcur ("name" c(100))
 
 UnzipOpen(zipfile)
 m.FilesInZip = UnzipFileCount()
 IF m.FilesInZip<=0
  IF  !m.lIsSilent
   MESSAGEBOX('¬ ZIP-¿–’»¬≈ Õ» ŒƒÕŒ√Œ ‘¿…À¿!',0+64,STR(m.FilesInZip))
  ENDIF 
 ENDIF 
 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 

 DIMENSION arr(1,13)

 IF !UnzipGotoTopFile()
  UnzipClose()
  RETURN .F.
 ENDIF 
 UnzipAFileInfo('arr')
 m.name = ALLTRIM(arr(1,1))
 IF fso.FileExists(pSoap+'\'+m.name)
  fso.DeleteFile(pSoap+'\'+m.name)
 ENDIF 
 INSERT INTO curcur FROM MEMVAR 
 UnzipFile(pSoap)
 
 DO WHILE UnZipGoToNextFile()
  UnzipAFileInfo('arr')
  m.name = arr(1,1)
  IF fso.FileExists(pSoap+'\'+m.name)
   fso.DeleteFile(pSoap+'\'+m.name)
  ENDIF 
  INSERT INTO curcur FROM MEMVAR 
  UnzipFile(pSoap)
 ENDDO 
 
 UnzipClose()
 
 SELECT curcur
 SCAN
  m.f_name = name
  DO CASE 
   CASE m.f_name = 'HM' && —˜ÂÚ
    m.hm_file = ALLTRIM(m.f_name)
   CASE m.f_name = 'LM' && –Â„ËÒÚ
    m.lm_file = ALLTRIM(m.f_name)
  ENDCASE 
 ENDSCAN 
 m.hm_dbf = UPPER(STRTRAN(LOWER(m.hm_file), 'xml', 'dbf'))
 m.lm_dbf = UPPER(STRTRAN(LOWER(m.lm_file), 'xml', 'dbf'))

 * Õ‡˜ËÌ‡ÂÏ Ô‡ÒËÚ¸ 'LM' (–Â„ËÒÚ)
 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 WAIT "«¿√–”« ¿ XML..." WINDOW NOWAIT 
 IF !oxml.load(pSoap+'\'+lm_file)
  RELEASE oXml
  IF  !m.lIsSilent
   MESSAGEBOX('Õ≈ ”ƒ¿ÀŒ—‹ «¿√–”«»“‹ m.lm_file ‘¿…À!',0+64,'oxml.load()')
  ENDIF 
  RETURN 
 ENDIF 
 WAIT CLEAR 
 
 m.n_recs = oxml.selectNodes('PERS_LIST/PERS').length
 IF m.n_recs=0
  RELEASE oXml
  MESSAGEBOX('¬ Œ“¬≈“≈ Õ» ŒƒÕŒ… «¿œ»—»!',0+64,'')
  RETURN 
 ENDIF 

 *MESSAGEBOX('¬ Œ“¬≈“≈ '+STR(m.n_recs)+' «¿œ»—≈…!',0+64,'')

 CREATE CURSOR people (recid i autoinc, mcod c(7), prmcod c(7), prmcods c(7), period c(7), d_beg d, d_end d, s_all n(11,2), ;
 	tip_p c(1), sn_pol c(25), tipp c(1), enp c(16), qq c(2), id_pac c(10), fam c(25), im c(20), ot c(20), dr c(8), w n(1),;
 	d_type c(1), sv c(3), recid_lpu c(7), ispr l, fil_id n(6))
 SELECT people 
 INDEX on id_pac TAG id_pac
 SET ORDER TO id_pac
 
 WAIT "XML->DBF..."  WINDOW NOWAIT 

 FOR m.n_rec = 0 TO m.n_recs-1
  m.orec = oxml.selectNodes('PERS_LIST/PERS').item(m.n_rec)
  
  m.fam    = ""
  m.im     = ""
  m.ot     = ""
  m.dr     = ""
  m.w      = 0

  m.id_pac = ALLTRIM(orec.selectNodes('ID_PAC').item(0).text)
  m.fam    = ALLTRIM(orec.selectNodes('FAM').item(0).text)
  m.im     = ALLTRIM(orec.selectNodes('IM').item(0).text)
  m.ot     = ALLTRIM(orec.selectNodes('OT').item(0).text)
  m.w      = ALLTRIM(orec.selectNodes('W').item(0).text)
  m.w      = INT(VAL(m.w))
  m.dr     = ALLTRIM(orec.selectNodes('DR').item(0).text)
  m.dr     = STRTRAN(m.dr,'-','')

  INSERT INTO people FROM MEMVAR 

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ELSE 
     WAIT "XML->DBF..."  WINDOW NOWAIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDFOR 

 * Õ‡˜ËÌ‡ÂÏ Ô‡ÒËÚ¸ 'HM' (—˜ÂÚ)
 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 WAIT "«¿√–”« ¿ XML..." WINDOW NOWAIT 
 IF !oxml.load(pSoap+'\'+hm_file)
  RELEASE oXml
  IF  !m.lIsSilent
   MESSAGEBOX('Õ≈ ”ƒ¿ÀŒ—‹ «¿√–”«»“‹ m.hm_file ‘¿…À!',0+64,'oxml.load()')
  ENDIF 
  RETURN 
 ENDIF 
 WAIT CLEAR 
 
 m.oschet = oxml.selectNodes('ZL_LIST/SCHET').item(0)
 m.lpu_id = INT(VAL(oschet.selectNodes('CODE_MO').item(0).text))
 m.mcod = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.mcod, '')
 m.period = ALLTRIM(oschet.selectNodes('YEAR').item(0).text)+ALLTRIM(oschet.selectNodes('MONTH').item(0).text)

 m.n_zaps = oxml.selectNodes('ZL_LIST/ZAP').length
 IF m.n_zaps=0
  RELEASE oXml
  MESSAGEBOX('¬ Œ“¬≈“≈ Õ» ŒƒÕŒ… «¿œ»—»!',0+64,'')
  RETURN 
 ENDIF 

 CREATE CURSOR talon (recid i autoinc, mcod c(7), period c(6), sn_pol c(25), c_i c(30), ds c(6), ds_0 c(6), pcod c(10), otd c(8),  ;
 	cod n(6), tip c(1), d_u d, k_u n(3), d_type c(1), s_all n(11,2), profil c(3), rslt n(4), prvs n(4), ishod n(4), usl_ok c(2), ;
 	vidpom c(2), vnov_m n(4), ds_1 c(6), ds_2 c(6), ds_3 c(6), p_cel n(5,2), lpu_ord n(4), date_ord d,;
 	idcase c(12), sl_id c(12), idserv c(2))
 WAIT "XML->DBF..."  WINDOW NOWAIT 
 
 *m.s_all = 0 
 FOR m.n_zap = 0 TO m.n_zaps-1
  m.o_zap = oxml.selectNodes('ZL_LIST/ZAP').item(m.n_zap)
  m.nzap = INT(VAL(o_zap.selectNodes('NZAP').item(0).text))
  
  m.o_pacient = o_zap.selectNodes('PACIENT').item(0)
  m.id_pac    = ALLTRIM(m.o_pacient.selectNodes('ID_PAC').item(0).text) && char
  m.tipp      = ALLTRIM(m.o_pacient.selectNodes('VPOLIS').item(0).text) && char
  m.s_pol     = ALLTRIM(m.o_pacient.selectNodes('SPOLIS').item(0).text) && char
  m.n_pol     = ALLTRIM(m.o_pacient.selectNodes('NPOLIS').item(0).text) && char
  m.sn_pol    = m.s_pol + m.n_pol
  m.IsEnp     = IsEnp(m.sn_pol)
  IF SEEK(m.id_pac, 'people') AND EMPTY(people.sn_pol)
   REPLACE sn_pol WITH m.sn_pol, enp WITH IIF(m.IsEnp, m.sn_pol, ''), tipp WITH m.tipp, mcod WITH m.mcod, ;
   	period WITH m.period IN people 
  ENDIF 
  
  m.n_z_sl = o_zap.selectNodes('Z_SL').length
  FOR m.z_sl = 0 TO m.n_z_sl-1
   m.o_z_sl = o_zap.selectNodes('Z_SL').item(m.z_sl)
   
   m.idcase   = ALLTRIM(o_z_sl.selectNodes('IDCASE').item(0).text)
   m.usl_ok   = ALLTRIM(o_z_sl.selectNodes('USL_OK').item(0).text)
   m.vidpom   = ALLTRIM(o_z_sl.selectNodes('VIDPOM').item(0).text)
   m.for_pom  = ALLTRIM(o_z_sl.selectNodes('FOR_POM').item(0).text)
   m.lpu_ord  = INT(VAL(ALLTRIM(o_z_sl.selectNodes('NPR_MO').item(0).text)))
   m.date_ord = ALLTRIM(o_z_sl.selectNodes('NPR_DATE').item(0).text)
   m.yy       = SUBSTR(m.date_ord, 1, 4)
   m.mm       = SUBSTR(m.date_ord, 5, 2)
   m.dd       = SUBSTR(m.date_ord, 7, 2)
   m.date_ord = CTOD(m.dd+'.'+m.mm+'.'+m.yy)
   m.vnov_m   = INT(VAL(ALLTRIM(o_z_sl.selectNodes('VNOV_M').item(0).text)))
   m.rslt     = INT(VAL(ALLTRIM(o_z_sl.selectNodes('RSLT').item(0).text)))
   m.ishod    = INT(VAL(ALLTRIM(o_z_sl.selectNodes('ISHOD').item(0).text)))
   m.idsp     = ALLTRIM(o_z_sl.selectNodes('IDSP').item(0).text)
   m.sumv     = VAL(ALLTRIM(o_z_sl.selectNodes('SUMV').item(0).text))
   IF SEEK(m.id_pac, 'people')
    m.o_sumv = people.s_all
    m.n_sumv = m.o_sumv + m.sumv
    REPLACE s_all WITH m.n_sumv IN people 
   ENDIF 
   
   m.n_sl = o_z_sl.selectNodes('SL').length
   FOR m.sl = 0 TO m.n_sl-1
    m.o_sl = o_z_sl.selectNodes('SL').item(m.sl)
    
    m.sl_id = ALLTRIM(o_sl.selectNodes('SL_ID').item(0).text)
    m.det   = INT(VAL(ALLTRIM(o_sl.selectNodes('DET').item(0).text))) && ÔËÁÌ‡Í ‰ÂÚÒÍÓ„Ó ÔÓÙËÎˇ 1/0
    IF o_sl.selectNodes('P_CEL').length>0
    m.p_cel = VAL(ALLTRIM(o_sl.selectNodes('P_CEL').item(0).text)) && ˆÂÎ¸ ÔÓÒÂ˘ÂÌËˇ
    ENDIF 
    m.c_i   = ALLTRIM(o_sl.selectNodes('NHISTORY').item(0).text)
    m.ds_0  = ALLTRIM(o_sl.selectNodes('DS0').item(0).text)
    m.ds_1  = ALLTRIM(o_sl.selectNodes('DS1').item(0).text)
    m.ds_2  = ALLTRIM(o_sl.selectNodes('DS2').item(0).text)
    m.ds_3  = ALLTRIM(o_sl.selectNodes('DS3').item(0).text)
    *m.pcod = ALLTRIM(o_sl.selectNodes('IDDOKT').item(0).text)
    
    m.n_usl = o_sl.selectNodes('USL').length
    FOR m.usl = 0 TO m.n_usl-1
     m.o_usl = o_sl.selectNodes('USL').item(m.usl)
     
     m.idserv = ALLTRIM(o_usl.selectNodes('IDSERV').item(0).text)
     *m.lpu_id = INT(VAL(ALLTRIM(o_usl.selectNodes('LPU_1').item(0).text)))
     m.otd    = ALLTRIM(o_usl.selectNodes('PODR').item(0).text)
     m.profil = ALLTRIM(o_usl.selectNodes('PROFIL').item(0).text)
     m.d_u    = STRTRAN(ALLTRIM(o_usl.selectNodes('DATE_OUT').item(0).text),'-','')
     m.yy     = SUBSTR(m.d_u,1,4)
     m.mm     = SUBSTR(m.d_u,5,2)
     m.dd     = SUBSTR(m.d_u,7,2)
     m.d_u    = CTOD(m.dd+'.'+m.mm+'.'+m.yy)
     m.ds     = ALLTRIM(o_usl.selectNodes('DS').item(0).text)
     m.cod    = INT(VAL(ALLTRIM(o_usl.selectNodes('CODE_USL').item(0).text)))
     m.k_u    = INT(VAL(ALLTRIM(o_usl.selectNodes('KOL_USL').item(0).text)))
     m.s_all  = VAL(ALLTRIM(o_usl.selectNodes('SUMV_USL').item(0).text))
     m.prvs   = INT(VAL(ALLTRIM(o_usl.selectNodes('PRVS').item(0).text)))
     m.pcod   = ALLTRIM(o_usl.selectNodes('CODE_MD').item(0).text)
     
     INSERT INTO talon FROM MEMVAR 

   ENDFOR 
  ENDFOR 
 ENDFOR 
  

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('¬€ ’Œ“»“≈ œ–≈–¬¿“‹ Œ¡–¿¡Œ“ ”?',4+32,'') == 6
     EXIT 
    ELSE 
     WAIT "XML->DBF..."  WINDOW NOWAIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDFOR 
 
 SELECT talon 
 COPY TO &pSoap\&hm_dbf
 USE 
 
 SELECT people
 COPY TO &pSoap\&lm_dbf
 USE 

 USE IN sprlpu
 
 WAIT CLEAR 
 *MESSAGEBOX(TRANSFORM(m.s_all,'999999.99'),0+64,'')
 
 MESSAGEBOX('OK!',0+64,'')

RETURN  
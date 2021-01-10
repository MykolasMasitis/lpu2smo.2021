* –‡·Ó˜ËÈ ÏÓ‰ÛÎ¸, ËÒÔÓÎ¸ÁÛ˛˘ËÈ DOM
PROCEDURE medpack_dom
 IF MESSAGEBOX(CHR(13)+CHR(10)+' ŒÕ¬≈–“»–Œ¬¿“‹ medicament_man_pack.xml?'+CHR(13)+CHR(10),;
 	4+32,'medicament_man_pack.xml')=7
  RETURN 
 ENDIF 
 
 pUpdDir = fso.GetParentFolderName(pbin)+'\UPDATE'
 IF !fso.FolderExists(pUpdDir)
  fso.CreateFolder(pUpdDir)
 ENDIF 

 SET DEFAULT TO (pUpdDir)
 csprfile = ''
 csprfile=GETFILE('xml')
 IF EMPTY(csprfile)
  MESSAGEBOX(CHR(13)+CHR(10)+'¬€ Õ»◊≈√Œ Õ≈ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 WAIT "«¿√–”« ¿ XML..." WINDOW NOWAIT 
 IF !oxml.load('&csprfile')
  RELEASE oXml
  MESSAGEBOX('Õ≈ ”ƒ¿ÀŒ—‹ «¿√–”«»“‹ '+csprfile+' ‘¿…À!',0+64,'oxml.load()')
  RETURN 
 ENDIF 
 WAIT CLEAR 

 m.n_recs = oxml.selectNodes('RESULTS/ROW').length
 IF m.n_recs=0
  RELEASE oXml
  MESSAGEBOX('¬ Œ“¬≈“≈ Õ» ŒƒÕŒ… «¿œ»—»!',0+64,'')
  RETURN 
 ENDIF 
 
 *CREATE CURSOR curss(ID C(32), PMP_ID C(10), R_UP C(10), NAME C(100), PRIM_TYPE C(50),;
 	PRIM_QTY_VALUE n(11,2), PRIM_QTY_UNIT c(10), PRIM_QTY_OKEI n(11,2), PRIM_MASS_VALUE n(11,2),;
	PRIM_MASS_UNIT C(10), PRIM_MASS_OKEI n(11,2), PRIM_VOL_VALUE n(11,2), PRIM_VOL_UNIT C(10), ;
	PRIM_VOL_OKEI n(11,2), SEC_TYPE C(50), SEC_PRIM_QTY N(11,2), TERT_TYPE C(50), TERT_SEC_QTY n(11,2),;
	IN_BULK L, VERSION_ID C(50))
 CREATE CURSOR curss(R_UP C(10), NAME C(100), PMP_ID C(10), TYPE C(50), ;
 	QTY_VALUE n(11,2), QTY_UNIT c(10), MASS_VALUE n(6,2), MASS_UNIT C(10), VOL_VALUE n(6,2), VOL_UNIT C(10))
 SELECT curss
 INDEX on r_up TAG r_up
 
 FOR m.n_rec = 0 TO m.n_recs-1
  m.orec = oxml.selectNodes('RESULTS/ROW').item(m.n_rec).selectNodes('COLUMN')
  
  *M.ID             = orec.item(0).text
  M.PMP_ID          = orec.item(1).text && PMP_MEDICAMENT_MAN_ID
  M.R_UP            = orec.item(2).text && CODE
  M.NAME            = orec.item(3).text
  M.TYPE       = orec.item(4).text && PRIM_TYPE
  M.QTY_VALUE  = VAL(ALLTRIM(STRTRAN(orec.item(5).text,',','.'))) && PRIM_QTY_VALUE
  M.QTY_UNIT   = orec.item(6).text && PRIM_QTY_UNIT
  *M.QTY_OKEI   = VAL(ALLTRIM(orec.item(7).text)) && PRIM_QTY_OKEI
  M.MASS_VALUE = VAL(ALLTRIM(STRTRAN(orec.item(8).text,',','.'))) && PRIM_MASS_VALUE
  M.MASS_UNIT  = orec.item(9).text && PRIM_MASS_UNIT
  *M.MASS_OKEI  = VAL(ALLTRIM(orec.item(10).text)) && PRIM_MASS_OKEI
  M.VOL_VALUE  = VAL(ALLTRIM(STRTRAN(orec.item(11).text,',','.'))) && PRIM_VOL_VALUE
  M.VOL_UNIT   = orec.item(12).text && PRIM_VOL_UNIT
  *M.VOL_OKEI   = VAL(ALLTRIM(orec.item(13).text)) && PRIM_VOL_OKEI
  *M.SEC_TYPE        = orec.item(14).text
  *M.SEC_PRIM_QTY    = VAL(ALLTRIM(orec.item(15).text))
  *M.TERT_TYPE       = orec.item(16).text
  *M.TERT_SEC_QTY    = VAL(ALLTRIM(orec.item(17).text))
  *M.IN_BULK         = IIF(orec.item(18).text='1', .T., .F.)
  *M.VERSION_ID      = orec.item(19).text

  *M.ID             = orec.selectNodes('COLUMN').item(0).text
  *M.PMP_MEDICAMENT_MAN_ID = orec.selectNodes('COLUMN').item(1).text
  *M.CODE            = orec.selectNodes('COLUMN').item(2).text
  *M.NAME            = orec.selectNodes('COLUMN').item(3).text
  *M.PRIM_TYPE       = orec.selectNodes('COLUMN').item(4).text
  *M.PRIM_QTY_VALUE  = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(5).text))
  *M.PRIM_QTY_UNIT   = orec.selectNodes('COLUMN').item(6).text
  *M.PRIM_QTY_OKEI   = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(7).text))
  *M.PRIM_MASS_VALUE = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(8).text))
  *M.PRIM_MASS_UNIT  = orec.selectNodes('COLUMN').item(9).text
  *M.PRIM_MASS_OKEI  = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(10).text))
  *M.PRIM_VOL_VALUE  = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(11).text))
  *M.PRIM_VOL_UNIT   = orec.selectNodes('COLUMN').item(12).text
  *M.PRIM_VOL_OKEI   = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(13).text))
  *M.SEC_TYPE        = orec.selectNodes('COLUMN').item(14).text
  *M.SEC_PRIM_QTY    = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(15).text))
  *M.TERT_TYPE       = orec.selectNodes('COLUMN').item(16).text
  *M.TERT_SEC_QTY    = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(17).text))
  *M.IN_BULK         = orec.selectNodes('COLUMN').item(18).text
  *M.VERSION_ID      = orec.selectNodes('COLUMN').item(19).text

  INSERT INTO curss FROM MEMVAR 
  
  IF m.n_rec/100 = INT(m.n_rec/100)
   WAIT STR(m.n_rec,6)+'\'+STR(m.n_recs,6) WINDOW NOWAIT 
  ENDIF 

 ENDFOR  
 
 SELECT curss 
 SET SAFETY OFF
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medpack.dbf')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medpack.dbf')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medpack.cdx')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medpack.cdx')
 ENDIF 

 COPY TO &pBase/&gcPeriod/nsi/medpack WITH cdx 
 SET SAFETY ON
 USE 
 RELEASE oXml
 
 MESSAGEBOX('‘¿…À —‘Œ–Ã»–Œ¬¿Õ!',0+64,'medpack')
		

RETURN 
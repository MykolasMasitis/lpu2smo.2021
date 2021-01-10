PROCEDURE medicament_dom
 IF MESSAGEBOX(CHR(13)+CHR(10)+' ŒÕ¬≈–“»–Œ¬¿“‹ medicament.xml?'+CHR(13)+CHR(10),4+32,'DOM (XML->DBF)')=7
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
 
 IF OpenFile(pBase+'\'+gcPeriod+'\nsi\tarion', 'tarion', 'shar', 'cod')>0
  IF USED('tarion')
   USE IN tarion
  ENDIF 
  RETURN 
 ENDIF 
 
 oXML  = CREATEOBJECT("MsXml2.DOMDocument")
 WAIT "«¿√–”« ¿ XML..." WINDOW NOWAIT 
 IF !oxml.load('&csprfile')
  RELEASE oXml
  MESSAGEBOX('Õ≈ ”ƒ¿ÀŒ—‹ «¿√–”«»“‹ '+csprfile+' ‘¿…À!',0+64,'oxml.load()')
  USE IN tarion
  RETURN 
 ENDIF 
 WAIT CLEAR 

 m.n_recs = oxml.selectNodes('RESULTS/ROW').length
 IF m.n_recs=0
  USE IN tarion
  RELEASE oXml
  MESSAGEBOX('¬ Œ“¬≈“≈ Õ» ŒƒÕŒ… «¿œ»—»!',0+64,'')
  RETURN 
 ENDIF 
 
 *CREATE CURSOR curss (ID C(32), CODE C(9), NAME C(100), GD_SID C(8), GD_NAME C(100), GD_UNIT C(100),;
 	GD_DOSAGE C(100), IS_STANDARD C(100), INFO_MISSING C(100), IS_VITAL C(100), DN_SID C(8), DN_NAME C(100),;
 	DF_SID C(5), DF_NAME C(100), GN_SID C(8), GN_NAME C(100), VERSION_ID C(100), IS_OMS_TARIF C(100), ;
 	MAX_SINGLE_DOSE C(100), IS_TARGET C(100), IS_ADJUVANT C(100), IS_HORMONAL C(100), IS_AUXILIARY C(100), ;
 	CODE_N020 C(100), STRENGTH_TEXT C(100), IS_STRENGTH_PERDOSE C(100), STRENGTH_MASS_VALUE C(100), ;
 	STRENGTH_MASS_UNIT C(100), STRENGTH_MASS_OKEI C(100), STRENGTH_VOL_VALUE C(100), STRENGTH_VOL_UNIT C(100),;
 	STRENGTH_VOL_OKEI C(100), ESKLP_SMNN_CODE C(100))
 *CREATE CURSOR curss (DD_SID C(10), DD_NAME C(100), IS_OMS L, ;
 	MASS_VALUE n(11,2), MASS_UNIT C(10), VOL_VALUE n(11,2), VOL_UNIT C(10), ;
 	GD_SID C(8), GD_NAME C(100), DN_SID C(8), DN_NAME C(100), DF_SID C(5), DF_NAME C(100), GN_SID C(8), GN_NAME C(100))
 CREATE CURSOR curss (DD_SID C(10), DD_NAME C(100), IS_OMS L, ;
 	MASS_VALUE n(6,2), MASS_UNIT C(10), VOL_VALUE n(6,2), VOL_UNIT C(10), GD_SID C(8), GD_NAME C(100), IS_TARGET n(1))
 INDEX on dd_sid TAG dd_sid
 
 FOR m.n_rec = 0 TO m.n_recs-1
  m.orec = oxml.selectNodes('RESULTS/ROW').item(m.n_rec)

  m.is_oms  = IIF(orec.selectNodes('COLUMN').item(17).text='1', .T., .F.)
  IF !m.is_oms
   *LOOP 
  ENDIF 
  m.gd_sid  = ALLTRIM(orec.selectNodes('COLUMN').item(3).text)
  IF !SEEK(m.gd_sid, 'tarion')
   *LOOP 
  ENDIF 
  
  *m.id      = orec.selectNodes('COLUMN').item(0).text
  m.dd_sid  = orec.selectNodes('COLUMN').item(1).text
  m.dd_name = orec.selectNodes('COLUMN').item(2).text
  m.gd_name = orec.selectNodes('COLUMN').item(4).text
  *m.dn_sid  = orec.selectNodes('COLUMN').item(10).text
  *m.dn_name = orec.selectNodes('COLUMN').item(11).text
  *m.df_sid  = orec.selectNodes('COLUMN').item(12).text
  *m.df_name = orec.selectNodes('COLUMN').item(13).text
  *m.gn_sid  = orec.selectNodes('COLUMN').item(14).text
  *m.gn_name = orec.selectNodes('COLUMN').item(15).text
  m.Is_Target = INT(VAL(ALLTRIM(orec.selectNodes('COLUMN').item(20).text)))

  M.MASS_VALUE = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(26).text))
  M.MASS_UNIT  = orec.selectNodes('COLUMN').item(27).text
  M.VOL_VALUE  = VAL(ALLTRIM(orec.selectNodes('COLUMN').item(29).text))
  M.VOL_UNIT   = orec.selectNodes('COLUMN').item(30).text

  INSERT INTO curss FROM MEMVAR 

 ENDFOR  
 
 SELECT curss 
 SET SAFETY OFF
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medicament.dbf')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medicament.dbf')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medicament.cdx')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medicament.cdx')
 ENDIF 

 COPY TO &pBase/&gcPeriod/nsi/medicament WITH cdx 
 SET SAFETY ON
 USE 
 RELEASE oXml

 USE IN tarion
 
 MESSAGEBOX('‘¿…À —‘Œ–Ã»–Œ¬¿Õ!',0+64,'medicamental')

RETURN 
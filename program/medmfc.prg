PROCEDURE MedMFC
 IF MESSAGEBOX(CHR(13)+CHR(10)+' ŒÕ¬≈–“»–Œ¬¿“‹ medicament_mfc.xml?'+CHR(13)+CHR(10),4+32,'XML->DBF')=7
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
 
 *IF OpenFile(pBase+'\'+gcPeriod+'\nsi\tarion', 'tarion', 'shar', 'cod')>0
 * IF USED('tarion')
 *  USE IN tarion
 * ENDIF 
 * RETURN 
 *ENDIF 
 
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
 
*	<ROW>
*		<COLUMN NAME="ID"><![CDATA[8AA94475AD410282E053C0A8C23399CE]]></COLUMN>
*		<COLUMN NAME="PMP_MEDICAMENT_ID"><![CDATA[DD0000408]]></COLUMN>
*		<COLUMN NAME="CODE"><![CDATA[MD00004946]]></COLUMN>
*		<COLUMN NAME="NAME"><![CDATA[¿Ï·ÓÍÒÓÎ Ú‡·Î. 30 Ï„, ŒÚ:  ‡ÌÓÌÙ‡Ï‡ ÔÓ‰‡Í¯Ì «¿Œ (–ÓÒÒËˇ), À—-001707 ÓÚ 14.09.2010 (Ó·Ì. 21.11.2016)]]></COLUMN>
*		<COLUMN NAME="MANUFACTURER_NAME"><![CDATA[ ‡ÌÓÌÙ‡Ï‡ ÔÓ‰‡Í¯Ì «¿Œ]]></COLUMN>
*		<COLUMN NAME="MANUFACTURER_COUNTRY"><![CDATA[–ÓÒÒËˇ]]></COLUMN>
*		<COLUMN NAME="CERTIFICATE_NUMBER"><![CDATA[À—-001707]]></COLUMN>
*		<COLUMN NAME="CERTIFICATE_ISSUED"><![CDATA[14.09.2010 00:00:00]]></COLUMN>
*		<COLUMN NAME="CERTIFICATE_END"><![CDATA[]]></COLUMN>
*		<COLUMN NAME="CERTIFICATE_OWNER_NAME"><![CDATA[ ‡ÌÓÌÙ‡Ï‡ ÔÓ‰‡Í¯Ì «¿Œ]]></COLUMN>
*		<COLUMN NAME="CERTIFICATE_OWNER_COUNTRY"><![CDATA[–ÓÒÒËˇ]]></COLUMN>
*		<COLUMN NAME="VERSION_ID"><![CDATA[024.010619]]></COLUMN>
*	<ROW>

 CREATE CURSOR curss (ID C(32), DD_ID C(10), MD_ID C(10), NAME C(250), MFC_NAME C(250), MFC_COUNTRY C(25),;
 	N_RU C(20), D_ISSUED D, D_END D, OWN_NAME C(250), OWN_COUNTRY C(25), VERSION_ID C(10))
 INDEX on dd_id TAG dd_id
 INDEX on n_ru TAG n_ru
 
 FOR m.n_rec = 0 TO m.n_recs-1
  m.orec = oxml.selectNodes('RESULTS/ROW').item(m.n_rec)
  
  m.id          = orec.selectNodes('COLUMN').item(0).text
  m.dd_id       = orec.selectNodes('COLUMN').item(1).text
  m.md_id       = orec.selectNodes('COLUMN').item(2).text
  m.name        = orec.selectNodes('COLUMN').item(3).text
  m.mfc_name    = orec.selectNodes('COLUMN').item(4).text
  m.mfc_cntry   = orec.selectNodes('COLUMN').item(5).text
  m.n_ru        = orec.selectNodes('COLUMN').item(6).text
  m.c_issued    = CTOD(LEFT(orec.selectNodes('COLUMN').item(7).text,10))
  m.c_end       = CTOD(LEFT(orec.selectNodes('COLUMN').item(8).text,10))
  m.own_name    = orec.selectNodes('COLUMN').item(9).text
  m.own_cntry   = orec.selectNodes('COLUMN').item(10).text
  m.version_id  = orec.selectNodes('COLUMN').item(11).text

  INSERT INTO curss FROM MEMVAR 

 ENDFOR  
 
 SELECT curss 
 *SET SAFETY OFF
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medicament_mfc.dbf')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medicament_mfc.dbf')
 ENDIF 
 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\medicament_mfc.cdx')
  fso.DeleteFile(pBase+'\'+m.gcPeriod+'\nsi\medicament_mfc.cdx')
 ENDIF 

 COPY TO &pBase/&gcPeriod/nsi/medicament_mfc WITH cdx 
 *SET SAFETY ON
 USE 
 RELEASE oXml
 
 MESSAGEBOX('‘¿…À —‘Œ–Ã»–Œ¬¿Õ!',0+64,'medicamental_mfc')
	

RETURN 
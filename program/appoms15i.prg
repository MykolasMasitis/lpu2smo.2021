PROCEDURE AppOMS15i
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÇÀÃÐÓÇÈÒÜ ÔÎÐÌÓ ÎÌÑ-15/3è?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 m.mmyy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),2)
 m.repname = 'oms_15_3i_'+m.qcod+'.xls'
 IF !fso.FileExists(pbase+'\'+gcperiod+'\'+m.repname)
  MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÎÁÍÀÐÓÆÅÍ ÔÀÉË '+m.repname+CHR(13)+CHR(10)+;
   'Â ÄÈÐÅÊÒÎÐÈÈ '+pbase+'\'+gcperiod+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË AISOMS.DBF'+CHR(13)+CHR(10)+;
   'Â ÄÈÐÅÊÒÎÐÈÈ '+pbase+'\'+gcperiod+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 *CREATE CURSOR oms15 (lpuid n(4), paz_outs n(6), paz n(6), ch01m n(6), ch01f n(6), ch14m n(6), ch14f n(6), ch514m n(6), ch514f n(6), ;
 	ch1517m n(6), ch1517f n(6), m1824 n(6), f1824 n(6), m2534 n(6), f2534 n(6), m3544 n(6), f3544 n(6), ;
 	m4559 n(6), f4559 n(6), m6068 n(6), f5564 n(6), m69 n(6), f65 n(6))

 CREATE CURSOR oms15 (lpuid n(4), paz_outs n(6), paz n(6), ch0001 n(6), ch0104 n(6), ch0517 n(6),;
    m1859 n(6), f1854 n(6), m6064 n(6), f5564 n(6), m6599 n(6), f6599 n(6))
 INDEX on lpuid TAG lpuid 

 CREATE CURSOR oms15st (lpuid n(4), paz_outs n(6), paz n(6), ch0001 n(6), ch0104 n(6), ch0517 n(6),;
    m1859 n(6), f1854 n(6), m6064 n(6), f5564 n(6), m6599 n(6), f6599 n(6))
 INDEX on lpuid TAG lpuid 
 
 oWMI = GETOBJECT('winmgmts://')

 cQuery = "select * from win32_process where name='excel.exe'"

 oResult = oWMI.ExecQuery(cQuery)
 IF oResult.Count>0
  MESSAGEBOX('ÎÁÍÀÐÓÆÅÍÎ'+STR(oResult.Count,2)+' ÇÀÏÓÙÅÍÍÛÕ ÏÐÎÖÅÑÑÀ EXCEL',0+64,'')
  FOR EACH oProcess IN oResult
   oProcess.Terminate(1)
  NEXT

  oResult = oWMI.ExecQuery(cQuery)
  IF oResult.Count>0
   MESSAGEBOX('ÎÁÍÀÐÓÆÅÍÎ'+STR(oResult.Count,2)+' ÇÀÏÓÙÅÍÍÛÕ ÏÐÎÖÅÑÑÀ EXCEL',0+64,'')
  ELSE 
   MESSAGEBOX('ÂÑÅ ÏÐÎÖÅÑÑÛ EXCEL ÇÀÊÐÛÒÛ!',0+64,'')
  ENDIF 
 ENDIF 

 WAIT "ÇÀÏÓÑÊ EXCEL..." WINDOW NOWAIT 
 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 m.IsVisible = .f. 
 m.IsQuit    = .t.

 LOCAL oEx as Exception

 m.err = .f. 
 TRY 
  oDoc = oExcel.WorkBooks.Open(pbase+'\'+gcperiod+'\'+m.repname,.T.)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  USE IN oms15
  MESSAGEBOX('ÍÅ ÓÄÀËÎÑÜ ÎÒÊÐÛÒÜ ÔÀÉË!'+CHR(13)+CHR(10)+oEx.Message, 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  RETURN .F.
 ENDIF 
 
 oExcel.Sheets(1).Select
 
 m.pnLpu = 0
 nCells=200
 FOR nCell=1 TO nCells
  m.nLpu  = oExcel.Cells(nCell,1).Value && ïîðÿäêîâûé íîìåð
  IF EMPTY(m.nLpu)
   LOOP 
  ENDIF 
  DO CASE 
   CASE VARTYPE(m.nLpu)='C'
    m.nLpu = INT(VAL(m.nLpu))
   CASE VARTYPE(m.nLpu)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  
  IF m.nLpu=0
   LOOP 
  ENDIF 
  
  IF m.nLpu-m.pnLpu != 1
   EXIT 
  ENDIF 
  m.nlpuid = oExcel.Cells(nCell,2).Value && lpu_id
  m.lpuid = m.nlpuid
  DO CASE 
   CASE VARTYPE(m.nlpuid)='C'
    m.nlpuid = INT(VAL(m.nlpuid))
   CASE VARTYPE(m.nlpuid)='N'
   OTHERWISE 
    EXIT 
  ENDCASE 
  
  m.paz = 0
  m.s_sum    = oExcel.Cells(nCell,6).Value && ñóììà ïî ñòðîêå, êîíòðîëüíàÿ öèôðà
  
  m.ch0001   = oExcel.Cells(nCell,7).Value
  m.ch0104   = oExcel.Cells(nCell,9).Value
  m.ch0517   = oExcel.Cells(nCell,11).Value
  m.m1859    = oExcel.Cells(nCell,13).Value
  m.f1854    = oExcel.Cells(nCell,15).Value
  m.m6064    = oExcel.Cells(nCell,17).Value
  m.f5564    = oExcel.Cells(nCell,19).Value
  m.m6599    = oExcel.Cells(nCell,21).Value
  m.f6599    = oExcel.Cells(nCell,23).Value
  
  DO CASE 
   CASE VARTYPE(m.s_sum)='C'
    m.s_sum = INT(VAL(m.s_sum))
   CASE VARTYPE(m.s_sum)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.ch0001)='C'
    m.ch0001 = INT(VAL(m.ch0001))
   CASE VARTYPE(m.ch0001)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch0104)='C'
    m.ch0104 = INT(VAL(m.ch0104))
   CASE VARTYPE(m.ch0104)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch0517)='C'
    m.ch0517 = INT(VAL(m.ch0517))
   CASE VARTYPE(m.ch0517)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m1859)='C'
    m.m1859 = INT(VAL(m.m1859))
   CASE VARTYPE(m.m1859)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f1854)='C'
    m.f1854 = INT(VAL(m.f1854))
   CASE VARTYPE(m.f1854)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m6064)='C'
    m.m6064 = INT(VAL(m.m6064))
   CASE VARTYPE(m.m6064)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f5564)='C'
    m.f5564 = INT(VAL(m.f5564))
   CASE VARTYPE(m.f5564)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m6599)='C'
    m.m6599 = INT(VAL(m.m6599))
   CASE VARTYPE(m.m6599)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f6599)='C'
    m.f6599 = INT(VAL(m.f6599))
   CASE VARTYPE(m.f6599)='N'
   OTHERWISE 
    LOOP 
  ENDCASE

  m.t_sum = m.ch0001 + m.ch0104 + m.ch0517 + m.m1859 + m.f1854 + m.m6064 + m.f5564 + m.m6599 + m.f6599 
  IF m.s_sum <>  m.t_sum 
   MESSAGEBOX('ÑÓÌÌÀ '+TRANSFORM(m.s_sum)+' ÍÅ ÐÀÂÍÀ '+TRANSFORM(m.t_sum),0+64, STR(m.nlpuid,4))
  ENDIF 
  
  m.paz = m.paz + ;
  	m.ch0001 + m.ch0104 + m.ch0517 + m.m1859 + m.f1854 + m.m6064 + m.f5564 + m.m6599 + m.f6599

  INSERT INTO oms15 FROM MEMVAR 

  IF m.nLpu>=1
   m.pnLpu = m.nLpu
  ENDIF 
  
 NEXT 
 
 SELECT oms15
 SET ORDER TO lpuid
 COPY TO &pBase\&gcPeriod\oms15 WITH cdx 


 oExcel.Sheets(2).Select
 
 m.pnLpu = 0
 nCells=200
 FOR nCell=1 TO nCells
  m.nLpu  = oExcel.Cells(nCell,1).Value && ïîðÿäêîâûé íîìåð
  IF EMPTY(m.nLpu)
   LOOP 
  ENDIF 
  DO CASE 
   CASE VARTYPE(m.nLpu)='C'
    m.nLpu = INT(VAL(m.nLpu))
   CASE VARTYPE(m.nLpu)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  
  IF m.nLpu=0
   LOOP 
  ENDIF 
  
  IF m.nLpu-m.pnLpu != 1
   EXIT 
  ENDIF 
  m.nlpuid = oExcel.Cells(nCell,2).Value && lpu_id
  m.lpuid = m.nlpuid
  DO CASE 
   CASE VARTYPE(m.nlpuid)='C'
    m.nlpuid = INT(VAL(m.nlpuid))
   CASE VARTYPE(m.nlpuid)='N'
   OTHERWISE 
    EXIT 
  ENDCASE 
  
  m.s_sum    = oExcel.Cells(nCell,6).Value && ñóììà ïî ñòðîêå, êîíòðîëüíàÿ öèôðà

  m.ch0001   = oExcel.Cells(nCell,7).Value
  m.ch0104   = oExcel.Cells(nCell,9).Value
  m.ch0517   = oExcel.Cells(nCell,11).Value
  m.m1859    = oExcel.Cells(nCell,13).Value
  m.f1854    = oExcel.Cells(nCell,15).Value
  m.m6064    = oExcel.Cells(nCell,17).Value
  m.f5564    = oExcel.Cells(nCell,19).Value
  m.m6599    = oExcel.Cells(nCell,21).Value
  m.f6599    = oExcel.Cells(nCell,23).Value
  
  DO CASE 
   CASE VARTYPE(m.s_sum)='C'
    m.s_sum = INT(VAL(m.s_sum))
   CASE VARTYPE(m.s_sum)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  DO CASE 
   CASE VARTYPE(m.ch0001)='C'
    m.ch0001 = INT(VAL(m.ch0001))
   CASE VARTYPE(m.ch0001)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch0104)='C'
    m.ch0104 = INT(VAL(m.ch0104))
   CASE VARTYPE(m.ch0104)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch0517)='C'
    m.ch0517 = INT(VAL(m.ch0517))
   CASE VARTYPE(m.ch0517)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m1859)='C'
    m.m1859 = INT(VAL(m.m1859))
   CASE VARTYPE(m.m1859)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f1854)='C'
    m.f1854 = INT(VAL(m.f1854))
   CASE VARTYPE(m.f1854)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m6064)='C'
    m.m6064 = INT(VAL(m.m6064))
   CASE VARTYPE(m.m6064)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f5564)='C'
    m.f5564 = INT(VAL(m.f5564))
   CASE VARTYPE(m.f5564)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m6599)='C'
    m.m6599 = INT(VAL(m.m6599))
   CASE VARTYPE(m.m6599)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f6599)='C'
    m.f6599 = INT(VAL(m.f6599))
   CASE VARTYPE(m.f6599)='N'
   OTHERWISE 
    LOOP 
  ENDCASE

  m.t_sum = m.ch0001 + m.ch0104 + m.ch0517 + m.m1859 + m.f1854 + m.m6064 + m.f5564 + m.m6599 + m.f6599 
  IF m.s_sum <>  m.t_sum 
   MESSAGEBOX('ÑÓÌÌÀ '+TRANSFORM(m.s_sum)+' ÍÅ ÐÀÂÍÀ '+TRANSFORM(m.t_sum),0+64, STR(m.nlpuid,4))
  ENDIF 
  
  m.paz = m.paz + ;
  	m.ch0001 + m.ch0104 + m.ch0517 + m.m1859 + m.f1854 + m.m6064 + m.f5564 + m.m6599 + m.f6599

  INSERT INTO oms15st FROM MEMVAR 

  IF m.nLpu>=1
   m.pnLpu = m.nLpu
  ENDIF 
  
 NEXT 
 
 SELECT oms15st
 SET ORDER TO lpuid
 COPY TO &pBase\&gcPeriod\oms15st WITH cdx 

 oDoc.Close(0)

 WAIT CLEAR 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 
 
 RELEASE oExcel
 
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÍÅ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF OpenFile(pcommon+'\pnorm', 'pnorm', 'shar', 'period')>0
  IF USED('pnorm')
   USE IN pnorm
  ENDIF 
  USE IN aisoms
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÍÅ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ELSE 
  SELECT pnorm
  IF !SEEK(m.gcperiod, 'pnorm')
   GO BOTTOM 
   SCATTER FIELDS EXCEPT period MEMVAR 
   INSERT INTO pnorm FROM MEMVAR 
  ELSE 
   SCATTER FIELDS EXCEPT period MEMVAR 
  ENDIF 
  USE IN pnorm 
  SELECT aisoms
 ENDIF 
 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  USE IN aisoms
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÍÅ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF fso.FileExists(pBase+'\'+m.gcPeriod+'\nsi\pnorm_iskl.dbf')
  IF  OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\pnorm_iskl', 'pn', 'shar', 'mcod')>0
   IF USED('pn')
    USE IN pn
   ENDIF 
   SELECT aisoms
  ENDIF 
 ELSE 
  MESSAGEBOX('ÑÏÐÀÂÎ×ÍÈÊ PNORM_ISKL ÍÅ ÍÀÉÄÅÍ!'+CHR(13)+CHR(10),0+64,'')
 ENDIF 

SELECT AisOms 
SET RELATION TO lpuid INTO oms15

SCAN
 m.mcod  = mcod
 m.lpuid = lpuid

 m.is_pn = .F.
 IF USED('pn')
  IF SEEK(m.mcod, 'pn') && !EMPTY(pn.lpu_id)
   m.is_pn = .T.
  ENDIF 
 ENDIF 
   
 WAIT m.mcod WINDOW NOWAIT 
   
 m.IsPilot  = IIF(SEEK(m.lpuid, 'pilot'), .T., .F.)

 m.pazval = 0 
 m.finval = 0

 m.pazval = oms15.ch0001 + oms15.ch0104 + oms15.ch0517 + ;
 	oms15.m1859 + oms15.f1854 + oms15.m6064 + oms15.f5564 + ;
 	oms15.m6599 + oms15.f6599

 IF m.is_pn = .T.
  m.kch0001 = pn.ch0001
  m.kch0104 = pn.ch0104
  m.kch0517 = pn.ch0517
  m.km1859  = pn.m1859
  m.kf1854  = pn.f1854
  m.km6064  = pn.m6064
  m.kf5564  = pn.f5564
  m.km6599  = pn.m6599
  m.kf6599  = pn.f6599

  m.finval = oms15.ch0001*m.kch0001 + oms15.ch0104*m.kch0104 + oms15.ch0517*m.kch0517 + oms15.m1859*m.km1859 + ;
   oms15.f1854*m.kf1854 + oms15.m6064*m.km6064 + oms15.f5564*m.kf5564 + oms15.m6599*m.km6599 + oms15.f6599*m.kf6599
   
 ELSE
  m.finval = oms15.ch0001*m.ch0001 + oms15.ch0104*m.ch0104 + oms15.ch0517*m.ch0517 + oms15.m1859*m.m1859 + ;
  	oms15.f1854*m.f1854 + oms15.m6064*m.m6064 + oms15.f5564*m.f5564 + oms15.m6599*m.m6599 + oms15.f6599*m.f6599
 ENDIF 
 
 REPLACE finval WITH IIF(m.IsPilot, m.finval, 0), pazval WITH IIF(m.IsPilot, m.pazval, 0)

ENDSCAN 

SET RELATION OFF INTO oms15
IF USED('pn')
 USE IN pn
ENDIF 
USE IN pilot
*USE IN aisoms

*MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')

IF OpenFile(pcommon+'\pnorms', 'pnorms', 'shar', 'period')>0
 IF USED('pnorms')
  USE IN pnorms
 ENDIF 
 USE IN aisoms
 MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÍÅ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')
 RETURN 
ELSE 
 SELECT pnorms
 IF !SEEK(m.gcperiod, 'pnorms')
  GO BOTTOM 
  SCATTER FIELDS EXCEPT period MEMVAR 
  INSERT INTO pnorms FROM MEMVAR 
 ELSE 
  SCATTER FIELDS EXCEPT period MEMVAR 
 ENDIF 
 USE IN pnorms
 SELECT aisoms
ENDIF 
 
IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\pilots', 'pilots', 'shar', 'lpu_id')>0
 IF USED('pilots')
  USE IN pilots
 ENDIF 
 USE IN aisoms
 MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÍÅ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')
 RETURN 
ENDIF 

SELECT AisOms 
SET RELATION TO lpuid INTO oms15st

SCAN
 m.mcod  = mcod
 m.lpuid = lpuid
   
 m.IsPilot  = IIF(SEEK(m.lpuid, 'pilots'), .T., .F.)

 m.pazvals = 0 
 m.finvals = 0

 m.pazvals = oms15st.ch0001 + oms15st.ch0104 + oms15st.ch0517 + ;
 	oms15st.m1859 + oms15st.f1854 + oms15st.m6064 + oms15st.f5564 + ;
 	oms15st.m6599 + oms15st.f6599

 m.finvals = oms15st.ch0001*m.ch0001 + oms15st.ch0104*m.ch0104 + oms15st.ch0517*m.ch0517 + oms15st.m1859*m.m1859 + ;
   oms15st.f1854*m.f1854 + oms15st.m6064*m.m6064 + oms15st.f5564*m.f5564 + oms15st.m6599*m.m6599 + oms15st.f6599*m.f6599
 
 REPLACE finvals WITH IIF(m.IsPilot, m.finvals, 0), pazvals WITH IIF(m.IsPilot, m.pazvals, 0)

ENDSCAN 

SET RELATION OFF INTO oms15st
USE IN pilots
USE IN aisoms

USE IN oms15
USE IN oms15st

MESSAGEBOX(CHR(13)+CHR(10)+'ÎÁÐÀÁÎÒÊÀ ÇÀÊÎÍ×ÅÍÀ!'+CHR(13)+CHR(10),0+64,'')

RETURN 

PROCEDURE AppOMS15
 IF MESSAGEBOX(CHR(13)+CHR(10)+'«¿√–”«»“‹ ◊»—À≈ÕÕŒ—“‹ œ–» –≈œÀ≈Õ»ﬂ?', 4+32, '“≈–¿œ»ﬂ')=7
  RETURN 
 ENDIF 
 
 m.mmyy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),2)
 flname = 'OMS15'+m.qcod+m.mmyy+'.xls'
 IF !fso.FileExists(pbase+'\'+gcperiod+'\'+flname)
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈ Œ¡Õ¿–”∆≈Õ ‘¿…À '+flname+CHR(13)+CHR(10)+;
   '¬ ƒ»–≈ “Œ–»» '+pbase+'\'+gcperiod+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À AISOMS.DBF'+CHR(13)+CHR(10)+;
   '¬ ƒ»–≈ “Œ–»» '+pbase+'\'+gcperiod+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 CREATE CURSOR oms15 (lpuid n(4), paz_outs n(6), paz n(6), ch01m n(6), ch01f n(6), ch14m n(6), ch14f n(6), ch514m n(6), ch514f n(6), ;
 	ch1517m n(6), ch1517f n(6), m1824 n(6), f1824 n(6), m2534 n(6), f2534 n(6), m3544 n(6), f3544 n(6), ;
 	m4559 n(6), f4559 n(6), m6068 n(6), f5564 n(6), m69 n(6), f65 n(6))
 INDEX on lpuid TAG lpuid 
 
 oWMI = GETOBJECT('winmgmts://')

 cQuery = "select * from win32_process where name='excel.exe'"

 oResult = oWMI.ExecQuery(cQuery)
 IF oResult.Count>0
  MESSAGEBOX('Œ¡Õ¿–”∆≈ÕŒ'+STR(oResult.Count,2)+' «¿œ”Ÿ≈ÕÕ€’ œ–Œ÷≈——¿ EXCEL',0+64,'')
  FOR EACH oProcess IN oResult
   oProcess.Terminate(1)
  NEXT

  oResult = oWMI.ExecQuery(cQuery)
  IF oResult.Count>0
   MESSAGEBOX('Œ¡Õ¿–”∆≈ÕŒ'+STR(oResult.Count,2)+' «¿œ”Ÿ≈ÕÕ€’ œ–Œ÷≈——¿ EXCEL',0+64,'')
  ELSE 
   MESSAGEBOX('¬—≈ œ–Œ÷≈——€ EXCEL «¿ –€“€!',0+64,'')
  ENDIF 
 ENDIF 

 WAIT "«¿œ”—  EXCEL..." WINDOW NOWAIT 
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
  oDoc = oExcel.WorkBooks.Open(pbase+'\'+gcperiod+'\'+flname,.T.)
 CATCH TO oEx
  m.err = .t. 
 ENDTRY 

 IF m.err = .t. 
  USE IN oms15
  MESSAGEBOX('Õ≈ ”ƒ¿ÀŒ—‹ Œ“ –€“‹ ‘¿…À!'+CHR(13)+CHR(10)+oEx.Message, 0+64, IIF(m.lpu_id>0, STR(m.lpu_id,4), 'getBillStatuses'))
  RETURN .F.
 ENDIF 
 
 oExcel.Sheets(1).Select
 
 m.pnLpu = 0
 nCells=200
 FOR nCell=1 TO nCells
  m.nLpu  = oExcel.Cells(nCell,1).Value
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
  m.lpuid = oExcel.Cells(nCell,2).Value
  DO CASE 
   CASE VARTYPE(m.lpuid)='C'
    m.lpuid = INT(VAL(m.lpuid))
   CASE VARTYPE(m.lpuid)='N'
   OTHERWISE 
    EXIT 
  ENDCASE 
  
  m.paz = 0
  
  m.ch01m   = oExcel.Cells(nCell,8).Value
  m.ch01f   = oExcel.Cells(nCell,9).Value
  m.ch14m   = oExcel.Cells(nCell,10).Value && Ï‡Î¸˜ËÍË < 5
  m.ch14f   = oExcel.Cells(nCell,11).Value && ‰Â‚Ó˜ÍË  < 5
  m.ch514m  = oExcel.Cells(nCell,12).Value
  m.ch514f  = oExcel.Cells(nCell,13).Value
  m.ch1517m = oExcel.Cells(nCell,14).Value && ÏÛÊ˜ËÌ˚ >= 5 <18
  m.ch1517f = oExcel.Cells(nCell,15).Value && ÊÂÌ˘ËÌ˚ >= 5 <18

  m.m1824   = oExcel.Cells(nCell,17).Value
  m.f1824   = oExcel.Cells(nCell,18).Value
  m.m2534   = oExcel.Cells(nCell,19).Value
  m.f2534   = oExcel.Cells(nCell,20).Value
  m.m3544   = oExcel.Cells(nCell,21).Value
  m.f3544   = oExcel.Cells(nCell,22).Value
  m.m4559   = oExcel.Cells(nCell,23).Value && ÏÛÊ˜ËÌ˚ >= 18 <60
  m.f4559   = oExcel.Cells(nCell,24).Value && ÊÂÌ˘ËÌ˚ >= 18 <55

  m.m6068   = oExcel.Cells(nCell,26).Value
  m.f5564   = oExcel.Cells(nCell,27).Value
  m.m69     = oExcel.Cells(nCell,28).Value && ÏÛÊ˜ËÌ˚ >= 60
  m.f65     = oExcel.Cells(nCell,29).Value && ÊÂÌ˘ËÌ˚ >= 55
  
  DO CASE 
   CASE VARTYPE(m.ch01m)='C'
    m.ch01m = INT(VAL(m.ch01m))
   CASE VARTYPE(m.ch01m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch01f)='C'
    m.ch01f = INT(VAL(m.ch01f))
   CASE VARTYPE(m.ch01f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch14m)='C'
    m.ch14m = INT(VAL(m.ch14m))
   CASE VARTYPE(m.ch14m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch14f)='C'
    m.ch14f = INT(VAL(m.ch14f))
   CASE VARTYPE(m.ch14f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch514m)='C'
    m.ch514m = INT(VAL(m.ch514m))
   CASE VARTYPE(m.ch514m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch514f)='C'
    m.ch514f = INT(VAL(m.ch514f))
   CASE VARTYPE(m.ch514f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch1517m)='C'
    m.ch1517m = INT(VAL(m.ch1517m))
   CASE VARTYPE(m.ch1517m)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.ch1517f)='C'
    m.ch1517f = INT(VAL(m.ch1517f))
   CASE VARTYPE(m.ch1517f)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m1824)='C'
    m.m1824 = INT(VAL(m.m1824))
   CASE VARTYPE(m.m1824)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f1824)='C'
    m.f1824 = INT(VAL(m.f1824))
   CASE VARTYPE(m.f1824)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m2534)='C'
    m.m2534 = INT(VAL(m.m2534))
   CASE VARTYPE(m.m2534)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f2534)='C'
    m.f2534 = INT(VAL(m.f2534))
   CASE VARTYPE(m.f2534)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m3544)='C'
    m.m3544 = INT(VAL(m.m3544))
   CASE VARTYPE(m.m3544)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f3544)='C'
    m.f3544 = INT(VAL(m.f3544))
   CASE VARTYPE(m.f3544)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m4559)='C'
    m.m4559 = INT(VAL(m.m4559))
   CASE VARTYPE(m.m4559)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f4559)='C'
    m.f4559 = INT(VAL(m.f4559))
   CASE VARTYPE(m.f4559)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.m6068)='C'
    m.m6068 = INT(VAL(m.m6068))
   CASE VARTYPE(m.m6068)='N'
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
   CASE VARTYPE(m.m69)='C'
    m.m69 = INT(VAL(m.m69))
   CASE VARTYPE(m.m69)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  DO CASE 
   CASE VARTYPE(m.f65)='C'
    m.f65 = INT(VAL(m.f65))
   CASE VARTYPE(m.f65)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 
  
  m.paz = m.paz + ;
  	m.ch01m+m.ch01f+m.ch14m+m.ch14f+m.ch514m+m.ch514f+m.ch1517m+m.ch1517f+m.m1824+m.f1824+;
  	m.m2534+m.f2534+m.m3544+m.f3544+m.m4559+m.f4559+m.m6068+m.f5564+m.m69+m.f65

  INSERT INTO oms15 FROM MEMVAR 

  IF m.nLpu>=1
   m.pnLpu = m.nLpu
  ENDIF 
  
 NEXT 
 
 oDoc.Close(0)
 
 SELECT oms15
 SET ORDER TO lpuid
 COPY TO &pBase\&gcPeriod\oms15 WITH cdx 
 *USE 

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
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ Õ≈ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF OpenFile(pcommon+'\pnorm', 'pnorm', 'shar', 'period')>0
  IF USED('pnorm')
   USE IN pnorm
  ENDIF 
  USE IN aisoms
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ Õ≈ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
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
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ Õ≈ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
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
  MESSAGEBOX('—œ–¿¬Œ◊Õ»  PNORM_ISKL Õ≈ Õ¿…ƒ≈Õ!'+CHR(13)+CHR(10),0+64,'')
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

 m.pazval = oms15.ch01m+oms15.ch01f+oms15.ch14m+oms15.ch14f+;
 	oms15.ch514m+oms15.ch514f+oms15.ch1517m+oms15.ch1517f+;
 	oms15.m1824+oms15.f1824+oms15.m2534+oms15.f2534+;
 	oms15.m3544+oms15.f3544+oms15.m4559+oms15.f4559+;
 	oms15.m6068+oms15.f5564+oms15.m69+oms15.f65

 IF m.is_pn = .T.
  m.km0001 = pn.m0001
  m.kf0001 = pn.f0001
  m.km0104 = pn.m1517
  m.kf0104 = pn.f1517
  m.km0514 = pn.m1517
  m.kf0514 = pn.f1517
  m.km1517 = pn.m1517
  m.kf1517 = pn.f1517
  m.km1824 = pn.m4559
  m.kf1824 = pn.f4554
  m.km2534 = pn.m4559
  m.kf2534 = pn.f4554
  m.km3544 = pn.m4559
  m.kf3544 = pn.f4554
  m.km4559 = pn.m4559
  m.kf4554 = pn.f4554
  m.km6068 = pn.m6999
  m.kf5564 = pn.f6599
  m.km6999 = pn.m6999
  m.kf6599 = pn.f6599

  m.finval = oms15.ch01m*m.km0001 + oms15.ch01f*m.kf0001 + oms15.ch14m*m.km0104 + oms15.ch14f*m.kf0104 + ;
   oms15.ch514m*m.km0514 + oms15.ch514f*m.kf0514 + oms15.ch1517m*m.km1517 + oms15.ch1517f*m.kf1517 + ;
   oms15.m1824*m.km1824 + oms15.f1824*m.kf1824 + oms15.m2534*m.km2534 + oms15.f2534*m.kf2534 + ;
   oms15.m3544*m.km3544 + oms15.f3544*m.kf3544 + oms15.m4559*m.km4559 + oms15.f4559*m.kf4554 + ;
   oms15.m6068*m.km6068 + oms15.f5564*m.kf5564 + oms15.m69*m.km6999 + oms15.f65*m.kf6599
   
 ELSE
  m.finval = oms15.ch01m*m.m0001 + oms15.ch01f*m.f0001 + oms15.ch14m*m.m0104 + oms15.ch14f*m.f0104 + ;
  	oms15.ch514m*m.m0514 + oms15.ch514f*m.f0514 + oms15.ch1517m*m.m1517 + oms15.ch1517f*m.f1517 + ;
  	oms15.m1824*m.m1824 + oms15.f1824*m.f1824 + oms15.m2534*m.m2534 + oms15.f2534*m.f2534 + ;
  	oms15.m3544*m.m3544 + oms15.f3544*m.f3544 + oms15.m4559*m.m4559 + oms15.f4559*m.f4554 + ;
  	oms15.m6068*m.m6068 + oms15.f5564*m.f5564 + oms15.m69*m.m6999 + oms15.f65*m.f6599
 ENDIF 
 
 REPLACE finval WITH IIF(m.IsPilot, m.finval, 0), pazval WITH IIF(m.IsPilot, m.pazval, 0)

ENDSCAN 

SET RELATION OFF INTO oms15
IF USED('pn')
 USE IN pn
ENDIF 
USE IN pilot
USE IN aisoms

MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
RETURN 

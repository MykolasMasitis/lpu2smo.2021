PROCEDURE LoadS532
 IF MESSAGEBOX(CHR(13)+CHR(10)+'«¿√–”«»“‹ —Õﬂ“»… œŒ 5.3.2?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 m.mmyy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),2)
 flname = 's_532.xls'
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
 
 IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
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
 oDoc = oExcel.WorkBooks.Open(pbase+'\'+gcperiod+'\'+flname,.T.)
 
 oExcel.Sheets(1).Select
 
 nCells=200
 FOR nCell=1 TO nCells

  m.mcod = oExcel.Cells(nCell,1).Value
  DO CASE 
   CASE VARTYPE(m.mcod)='C'
   CASE VARTYPE(m.mcod)='N'
    m.mcod = PADL(INT(m.mcod),7,'0')
   OTHERWISE 
    EXIT 
  ENDCASE 
  
  m.s_532   = oExcel.Cells(nCell,2).Value

  DO CASE 
   CASE VARTYPE(m.s_532)='C'
    m.s_532 = INT(VAL(m.s_532))
   CASE VARTYPE(m.s_532)='N'
   OTHERWISE 
    LOOP 
  ENDCASE 

  IF SEEK(m.mcod, 'aisoms')
   UPDATE aisoms SET s_532=m.s_532 WHERE mcod = m.mcod
  ENDIF 
  
 NEXT 

 IF USED('aisoms')
  USE IN aisoms
 ENDIF 

 WAIT CLEAR 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'Œ¡–¿¡Œ“ ¿ «¿ ŒÕ◊≈Õ¿!'+CHR(13)+CHR(10),0+64,'')
RETURN 
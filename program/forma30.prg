PROCEDURE FormA30
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ —“¿“»—“» ” —Õﬂ“»… œŒ Ã› ?'+CHR(13)+CHR(10),4+32,'‘Œ–Ã¿ ¿3-0')=7
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pmee)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ '+UPPER(ALLTRIM(pmee))+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FolderExists(pmee+'\'+m.gcperiod)
  fso.CreateFolder(pmee+'\'+m.gcperiod)
 ENDIF 
 IF !fso.FolderExists(pbase+'\'+m.gcperiod)
  MESSAGEBOX('Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\aisoms.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À AISOMS.DBF!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+m.gcperiod+'\nsi\sookodxx.dbf')
  MESSAGEBOX('Œ“—”“—“¬”≈“ ‘¿…À SOOKODXX.DBF!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\sookodxx', 'sookod', 'shar', 'er_c')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 
 
 PUBLIC oExcel AS Excel.Application
 WAIT "«‡ÔÛÒÍ MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
*  oExcel.Quit
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 m.BookName = 'A30'
 m.nOpBooks = oExcel.Workbooks.Count 
 IF m.nOpBooks>0
  FOR m.nBook=1 TO m.nOpBooks
   m.cBookName = LOWER(ALLTRIM(oExcel.Workbooks.Item(m.nBook).Name))
   IF m.cBookName=m.BookName+'.xls'
    oExcel.Workbooks.Item(m.nBook).Close 
   ENDIF 
  NEXT 
 ENDIF 

 oExcel.SheetsInNewWorkbook = 1
 oBook  = oExcel.WorkBooks.Add
 oSheet = oexcel.ActiveSheet
 oSheet.PageSetup.Orientation=2
 oSheet.name='ÙÓÏ‡ ÿ-5'
 
 WITH oExcel
  .Rows(1).RowHeight=25
  .Rows(2).RowHeight=25
  .Rows(2).WrapText = .t.
  .Cells(1,1) = '—Ú‡ÚËÒÚËÍ‡ ÒÌˇÚËÈ ÔÓ Ã›  Á‡ '+NameOfMonth(m.tmonth)+' '+STR(tyear,4)+' „Ó‰‡'
  .Cells(1,1).HorizontalAlignment = -4108
  .Rows(1).VerticalAlignment = -4108
  .Rows(2).VerticalAlignment = -4108
  .Range(.Cells(1,1),.Cells(1,5)).Merge
  .Cells(2,1)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ Ã√‘ŒÃ—'
  .Cells(2,2)=' Ó‰ ÃŒ'
  .Cells(2,3)=' Ó‰ Àœ”'
  .Cells(2,4)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ ‘‘ŒÃ—'
  .Cells(2,5)='Õ‡ËÏÂÌÓ‚‡ÌËÂ ‰ÂÙÂÍÚ‡'
  .Cells(2,6)=' ÓÎ-‚Ó ‰ÂÙÂÍÚÓ‚'
  .Cells(2,7)='—ÛÏÏ‡ ÒÌˇÚËÈ ÔÓ ‰ÂÙÂÍÚÛ'
  .Cells(2,8)='¿œœ'
  .Cells(2,9)='ƒ—“'
  .Cells(2,10)='—“'
  .Cells(2,11)=' Ó‰ ÂÂÒÚ‡'
 ENDWITH 

 WAIT "–¿—◊≈“..." WINDOW NOWAIT 

 CREATE CURSOR curdata (c_err c(3), k_u n(6), s_all n(11,2), mcod c(7), lpuid n(4), s_app n(11,2), s_dst n(11,2), s_st n(11,2))
 INDEX ON c_err+mcod TAG c_err
 SET ORDER TO c_err 

 SELECT aisoms
 SCAN 
  m.mcod  = mcod
  m.lpuid = lpuid
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'merror', 'shar', 'rid')>0
   IF USED('merror')
    USE IN merror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF RECCOUNT('merror')<=0
   IF USED('merror')
    USE IN merror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  SELECT merror 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
   IF USED('merror')
    USE IN merror
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\people', 'people', 'shar', 'sn_pol')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('merror')
    USE IN merror
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF  
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'rerror', 'shar', 'rrid', 'again')>0
   IF USED('people')
    USE IN people
   ENDIF 
   IF USED('rerror')
    USE IN rerror
   ENDIF 
   IF USED('merror')
    USE IN merror
   ENDIF 
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 



  CREATE CURSOR curid (recid i)
  INDEX ON recid TAG recid
  SET ORDER TO recid 
  
  SELECT people
  SET RELATION TO recid INTO rerror
  SELECT talon
  SET RELATION TO sn_pol INTO people
  SELECT merror
  SET RELATION TO rid INTO talon 
  SCAN 
   IF DELETED()
    LOOP 
   ENDIF 
   m.fll = UPPER(f)
   IF m.fll!='S'
    LOOP 
   ENDIF 

   m.recid  = talon.recid
   m.cod    = talon.cod
   m.c_err  = IIF(c_err='PKA', rerror.c_err, c_err)
   m.s_all  = talon.s_all+talon.s_lek
   m.otd    = SUBSTR(talon.otd,2,2)
   m.s_app = IIF(IsApp(m.otd), m.s_all, 0)
   m.s_st  = IIF(IsGsp(m.otd), m.s_all, 0)
   m.s_dst = IIF(IsDst(m.otd), m.s_all, 0)
   m.ok_u   = 0
   m.os_all = 0
   m.os_app = 0
   m.os_st  = 0
   m.os_dst = 0

   IF SEEK(m.c_err+m.mcod, 'curdata')
    m.ok_u   = curdata.k_u
    m.os_all = curdata.s_all
    m.os_st  = curdata.s_st
    m.os_dst = curdata.s_dst
    m.os_app = curdata.s_app

    IF SEEK(m.recid, 'curid')
     m.s_all = 0
     m.s_st  = 0
     m.s_dst = 0
     m.s_app = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    
    m.nk_u   = m.ok_u + 1
    m.ns_all = m.os_all + m.s_all
    m.ns_app = m.os_app + IIF(m.IsApp(m.otd), m.s_all, 0)
    m.ns_st  = m.os_st  + IIF(m.IsGsp(m.otd), m.s_all, 0)
    m.ns_dst = m.os_dst + IIF(m.IsDst(m.otd), m.s_all, 0)
    UPDATE curdata SET k_u=m.nk_u, s_all=m.ns_all, s_app=m.ns_app, s_dst=m.ns_dst, s_st=m.ns_st WHERE mcod=m.mcod AND c_err=m.c_err
   ELSE 
    IF SEEK(m.recid, 'curid')
     m.s_all = 0
     m.s_st  = 0
     m.s_dst = 0
     m.s_app = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    INSERT INTO curdata (c_err,k_u,s_all,s_app,s_st,s_dst,mcod,lpuid) VALUES ;
    	(m.c_err,1,m.s_all,m.s_app,m.s_st,m.s_dst,m.mcod,m.lpuid)
   ENDIF 


  ENDSCAN 
  SET RELATION OFF INTO talon 
  USE IN talon 
  USE IN people
  USE IN merror 
  USE IN rerror 
  USE IN curid

  SELECT aisoms 
  
 ENDSCAN
 USE IN aisoms
 WAIT CLEAR 

 WITH oExcel
  .Columns(1).NumberFormat  = '@'
  .Columns(1).ColumnWidth   = 5
  .Columns(2).NumberFormat  = '@'
  .Columns(2).ColumnWidth   = 10
  .Columns(3).NumberFormat  = '@'
  .Columns(3).ColumnWidth   = 10
  .Columns(4).NumberFormat  = '@'
  .Columns(4).ColumnWidth   = 5
  .Columns(5).NumberFormat  = '@'
  .Columns(5).ColumnWidth   = 75
  .Columns(6).NumberFormat  = '0'
  .Columns(6).ColumnWidth   = 10
  .Columns(7).NumberFormat  = '0.00'
  .Columns(7).ColumnWidth   = 20
  .Columns(8).NumberFormat  = '0.00'
  .Columns(8).ColumnWidth   = 20
  .Columns(9).NumberFormat  = '0.00'
  .Columns(9).ColumnWidth   = 20
  .Columns(10).NumberFormat  = '0.00'
  .Columns(10).ColumnWidth   = 20
 ENDWITH 

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Œ“◊≈“¿..." WINDOW NOWAIT 
 SELECT curdata 
 SET RELATION TO LEFT(c_err,2) INTO sookod
 m.n = 3
 m.k_u = 0
 m.s_all = 0 
 SCAN 
  WITH oExcel
   .Cells(m.n,1) = c_err
   .Cells(m.n,2) = lpuid
   .Cells(m.n,3) = mcod
   .Cells(m.n,4) = sookod.osn230
   .Cells(m.n,5) = sookod.comment
   .Cells(m.n,6) = k_u
   .Cells(m.n,7) = s_all
   .Cells(m.n,8) = s_app
   .Cells(m.n,9) = s_dst
   .Cells(m.n,10) = s_st
   m.k_u = m.k_u + k_u
   m.s_all = m.s_all + s_all
  ENDWITH 
  m.n = m.n + 1
 ENDSCAN 
 COPY TO  &pmee\&gcperiod\mekstat
 SET RELATION OFF INTO sookod
 USE IN curdata 
 USE IN sookod

 oExcel.Range(oExcel.Cells(m.n,1),oExcel.Cells(m.n,3)).Merge
 oExcel.Cells(m.n,1) = '»ÚÓ„Ó:'
 oExcel.Rows(m.n).RowHeight=25
 oExcel.Cells(m.n,5) = m.k_u
 oExcel.Cells(m.n,6) = m.s_all
 oExcel.Rows(m.n).VerticalAlignment = -4108
 WAIT CLEAR 

 IF fso.FileExists(pmee+'\'+m.gcperiod+'\'+m.BookName+'.xls')
  TRY 
   fso.DeleteFile(pmee+'\'+m.gcperiod+'\'+m.BookName+'.xls')
   oBook.SaveAs(pmee+'\'+m.gcperiod+'\'+m.BookName,18)
  CATCH  
   MESSAGEBOX('‘¿…À '+m.BookName+'.XLS Œ “–€“!',0+64,'')
  ENDTRY 
 ELSE 
  oBook.SaveAs(pmee+'\'+m.gcperiod+'\'+m.BookName,18)
 ENDIF 
 oExcel.Visible = .t.

RETURN 

FUNCTION IsApp(para1)
RETURN IIF(INLIST(para1,'00','01','22','08','85','90','91','92','93') , .T., .F.)

FUNCTION IsDst(para1)
RETURN IIF(INLIST(para1,'80','81') , .T., .F.)

FUNCTION IsGsp(para1)
RETURN IIF(!IsApp(para1) AND !IsDst(para1), .T., .F.)
PROCEDURE FormSh55Bis
 IF MESSAGEBOX('—‘Œ–Ã»–Œ¬¿“‹ —“¿“»—“» ” —Õﬂ“»… œŒ Ã››?'+CHR(13)+CHR(10),4+32,'–¿—ÿ»–≈ÕÕ¿ﬂ')=7
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
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 m.BookName = 'sh55'
 m.nOpBooks = oExcel.Workbooks.Count 
 IF m.nOpBooks>0
  FOR m.nBook=1 TO m.nOpBooks
   m.cBookName = LOWER(ALLTRIM(oExcel.Workbooks.Item(m.nBook).Name))
   IF m.cBookName=m.BookName+'.xls'
    oExcel.Workbooks.Item(m.nBook).Close 
   ENDIF 
  NEXT 
 ENDIF 

 oExcel.SheetsInNewWorkbook = 3
 oBook  = oExcel.WorkBooks.Add

 oexcel.Sheets(3).Select
 oSheet = oexcel.ActiveSheet
 oSheet.PageSetup.Orientation=2
 oSheet.name='› Ãœ'
 WITH oExcel
  .Rows(1).RowHeight=25
  .Rows(2).RowHeight=25
  .Rows(2).WrapText = .t.
  .Cells(1,1) = '—Ú‡ÚËÒÚËÍ‡ ÒÌˇÚËÈ ÔÓ › Ãœ Á‡ '+NameOfMonth(m.tmonth)+' '+STR(tyear,4)+' „Ó‰‡'
  .Cells(1,1).HorizontalAlignment = -4108
  .Rows(1).VerticalAlignment = -4108
  .Rows(2).VerticalAlignment = -4108
  .Range(.Cells(1,1),.Cells(1,13)).Merge
  .Cells(2,1)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ Ã√‘ŒÃ—'
  .Cells(2,2)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ ‘‘ŒÃ—'
  .Cells(2,3)='Õ‡ËÏÂÌÓ‚‡ÌËÂ ‰ÂÙÂÍÚ‡'
  .Cells(2,4)='¿œœ'
  .Cells(2,5)='—“'
  .Cells(2,6)='ƒ—“'
  .Cells(2,7)='—Ãœ'
  .Cells(2,8)=' ÓÎ-‚Ó ‰ÂÙÂÍÚÓ‚'
  .Cells(2,9)='¿œœ'
  .Cells(2,10)='—“'
  .Cells(2,11)='ƒ—“'
  .Cells(2,12)='—Ãœ'
  .Cells(2,13)='—ÛÏÏ‡ ÒÌˇÚËÈ ÔÓ ‰ÂÙÂÍÚÛ'

  .Columns(1).NumberFormat  = '@'
  .Columns(1).ColumnWidth   = 5
  .Columns(2).NumberFormat  = '@'
  .Columns(2).ColumnWidth   = 5
  .Columns(3).NumberFormat  = '@'
  .Columns(3).ColumnWidth   = 75
  .Columns(8).NumberFormat  = '0'
  .Columns(8).ColumnWidth   = 10
  .Columns(9).NumberFormat  = '0.00'
  .Columns(9).ColumnWidth   = 20
  .Columns(10).NumberFormat  = '0.00'
  .Columns(10).ColumnWidth   = 20
  .Columns(11).NumberFormat  = '0.00'
  .Columns(11).ColumnWidth   = 20
  .Columns(12).NumberFormat  = '0.00'
  .Columns(12).ColumnWidth   = 20
  .Columns(13).NumberFormat  = '0.00'
  .Columns(13).ColumnWidth   = 20
 ENDWITH 

 oexcel.Sheets(1).Select
 oSheet = oexcel.ActiveSheet
 oSheet.PageSetup.Orientation=2
 oSheet.name='Ã› '
 WITH oExcel
  .Rows(1).RowHeight=25
  .Rows(2).RowHeight=25
  .Rows(2).WrapText = .t.
  .Cells(1,1) = '—Ú‡ÚËÒÚËÍ‡ ÒÌˇÚËÈ ÔÓ Ã›  Á‡ '+NameOfMonth(m.tmonth)+' '+STR(tyear,4)+' „Ó‰‡'
  .Cells(1,1).HorizontalAlignment = -4108
  .Rows(1).VerticalAlignment = -4108
  .Rows(2).VerticalAlignment = -4108
  .Range(.Cells(1,1),.Cells(1,13)).Merge
  .Cells(2,1)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ Ã√‘ŒÃ—'
  .Cells(2,2)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ ‘‘ŒÃ—'
  .Cells(2,3)='Õ‡ËÏÂÌÓ‚‡ÌËÂ ‰ÂÙÂÍÚ‡'
  .Cells(2,4)='¿œœ'
  .Cells(2,5)='—“'
  .Cells(2,6)='ƒ—“'
  .Cells(2,7)='—Ãœ'
  .Cells(2,8)=' ÓÎ-‚Ó ‰ÂÙÂÍÚÓ‚'
  .Cells(2,9)='¿œœ'
  .Cells(2,10)='—“'
  .Cells(2,11)='ƒ—“'
  .Cells(2,12)='—Ãœ'
  .Cells(2,13)='—ÛÏÏ‡ ÒÌˇÚËÈ ÔÓ ‰ÂÙÂÍÚÛ'

  .Columns(1).NumberFormat  = '@'
  .Columns(1).ColumnWidth   = 5
  .Columns(2).NumberFormat  = '@'
  .Columns(2).ColumnWidth   = 5
  .Columns(3).NumberFormat  = '@'
  .Columns(3).ColumnWidth   = 75
  .Columns(8).NumberFormat  = '0'
  .Columns(8).ColumnWidth   = 10
  .Columns(9).NumberFormat  = '0.00'
  .Columns(9).ColumnWidth   = 20
  .Columns(10).NumberFormat  = '0.00'
  .Columns(10).ColumnWidth   = 20
  .Columns(11).NumberFormat  = '0.00'
  .Columns(11).ColumnWidth   = 20
  .Columns(12).NumberFormat  = '0.00'
  .Columns(12).ColumnWidth   = 20
  .Columns(13).NumberFormat  = '0.00'
  .Columns(13).ColumnWidth   = 20
 ENDWITH 

 oexcel.Sheets(2).Select
 oSheet = oexcel.ActiveSheet
 oSheet.PageSetup.Orientation=2
 oSheet.name='Ã››'
 WITH oExcel
  .Rows(1).RowHeight=25
  .Rows(2).RowHeight=25
  .Rows(2).WrapText = .t.
  .Cells(1,1) = '—Ú‡ÚËÒÚËÍ‡ ÒÌˇÚËÈ ÔÓ Ã›› Á‡ '+NameOfMonth(m.tmonth)+' '+STR(tyear,4)+' „Ó‰‡'
  .Cells(1,1).HorizontalAlignment = -4108
  .Rows(1).VerticalAlignment = -4108
  .Rows(2).VerticalAlignment = -4108
  .Range(.Cells(1,1),.Cells(1,13)).Merge
  .Cells(2,1)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ Ã√‘ŒÃ—'
  .Cells(2,2)=' Ó‰ ‰ÂÙÂÍÚ‡ ÔÓ ‘‘ŒÃ—'
  .Cells(2,3)='Õ‡ËÏÂÌÓ‚‡ÌËÂ ‰ÂÙÂÍÚ‡'
  .Cells(2,4)='¿œœ'
  .Cells(2,5)='—“'
  .Cells(2,6)='ƒ—“'
  .Cells(2,7)='—Ãœ'
  .Cells(2,8)=' ÓÎ-‚Ó ‰ÂÙÂÍÚÓ‚'
  .Cells(2,9)='¿œœ'
  .Cells(2,10)='—“'
  .Cells(2,11)='ƒ—“'
  .Cells(2,12)='—Ãœ'
  .Cells(2,13)='—ÛÏÏ‡ ÒÌˇÚËÈ ÔÓ ‰ÂÙÂÍÚÛ'

  .Columns(1).NumberFormat  = '@'
  .Columns(1).ColumnWidth   = 5
  .Columns(2).NumberFormat  = '@'
  .Columns(2).ColumnWidth   = 5
  .Columns(3).NumberFormat  = '@'
  .Columns(3).ColumnWidth   = 75
  .Columns(8).NumberFormat  = '0'
  .Columns(8).ColumnWidth   = 10
  .Columns(9).NumberFormat  = '0.00'
  .Columns(9).ColumnWidth   = 20
  .Columns(10).NumberFormat  = '0.00'
  .Columns(10).ColumnWidth   = 20
  .Columns(11).NumberFormat  = '0.00'
  .Columns(11).ColumnWidth   = 20
  .Columns(12).NumberFormat  = '0.00'
  .Columns(12).ColumnWidth   = 20
  .Columns(13).NumberFormat  = '0.00'
  .Columns(13).ColumnWidth   = 20
 ENDWITH 

 WAIT "–¿—◊≈“..." WINDOW NOWAIT 

 CREATE CURSOR curdatamek (er_c c(2), osn230 c(5), ku_amb n(6), ku_dst n(6), ku_st n(6), ku_smp n(6), k_u n(6), ;
 	s_amb n(11,2), s_dst n(11,2), s_st n(11,2), s_smp n(11,2), s_all n(11,2))
 INDEX ON er_c TAG er_c
 SET ORDER TO er_c 
 CREATE CURSOR curdatamee (er_c c(2), osn230 c(5), ku_amb n(6), ku_dst n(6), ku_st n(6), ku_smp n(6), k_u n(6), ;
 	s_amb n(11,2), s_dst n(11,2), s_st n(11,2), s_smp n(11,2), s_all n(11,2))
 INDEX ON er_c TAG er_c
 SET ORDER TO er_c 
 CREATE CURSOR curdataekmp (er_c c(2), osn230 c(5), ku_amb n(6), ku_dst n(6), ku_st n(6), ku_smp n(6), k_u n(6), ;
 	s_amb n(11,2), s_dst n(11,2), s_st n(11,2), s_smp n(11,2), s_all n(11,2))
 INDEX ON er_c TAG er_c
 SET ORDER TO er_c 

 SELECT aisoms
 SCAN 
  m.mcod     = mcod
  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod+'.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\m'+m.mcod, 'merror', 'shar', 'recid')>0
   IF USED('merror')
    USE IN merror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\e'+m.mcod, 'eerror', 'shar', 'rid')>0
   USE IN merror
   IF USED('eerror')
    USE IN eerror
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF 
  IF RECCOUNT('merror')<=0 AND RECCOUNT('eerror')<=0
   USE IN eerror
   USE IN merror
   SELECT aisoms
   LOOP 
  ENDIF 
  SELECT merror 
  IF OpenFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\talon', 'talon', 'shar', 'recid')>0
   USE IN eerror
   USE IN merror
   IF USED('talon')
    USE IN talon
   ENDIF 
   SELECT aisoms
   LOOP 
  ENDIF
  
  CREATE CURSOR curid (recid i)
  INDEX ON recid TAG recid
  SET ORDER TO recid 

  SELECT merror
  SET RELATION TO recid INTO talon 
  SCAN 
   IF DELETED()
    LOOP 
   ENDIF 
   
   m.et = et
   IF !INLIST(m.et,'2','3','7','8') && Ã››
    LOOP 
   ENDIF 

   m.recid  = talon.recid
   m.cod    = talon.cod
   m.er_c   = LEFT(err_mee,2)
   m.osn230 = osn230
   m.s_all  = talon.s_all
   m.ok_u   = 0
   m.os_all = 0

   IF SEEK(m.er_c, 'curdatamee')
    m.ok_u    = curdatamee.k_u
    m.oku_amb = curdatamee.ku_amb
    m.oku_dst = curdatamee.ku_dst
    m.oku_st  = curdatamee.ku_st
    m.oku_smp = curdatamee.ku_smp
    m.os_amb  = curdatamee.s_amb
    m.os_dst  = curdatamee.s_dst
    m.os_st   = curdatamee.s_st
    m.os_smp  = curdatamee.s_smp
    m.os_all  = curdatamee.s_all
    IF SEEK(m.recid, 'curid')
     m.s_all = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    m.nk_u   = m.ok_u + 1
    m.nku_amb = m.oku_amb + IIF(IsPlk(m.cod), 1, 0)
    m.nku_dst = m.oku_dst + IIF(IsDst(m.cod), 1, 0)
    m.nku_st  = m.oku_st  + IIF(IsGsp(m.cod), 1, 0)
    m.nku_smp = m.oku_smp + IIF(Is02(m.cod), 1, 0)
    m.ns_all = m.os_all + m.s_all
    m.ns_amb = m.os_amb + IIF(IsPlk(m.cod), m.s_all, 0)
    m.ns_dst = m.os_dst + IIF(IsDst(m.cod), m.s_all, 0)
    m.ns_st  = m.os_st  + IIF(IsGsp(m.cod), m.s_all, 0)
    m.ns_smp = m.os_smp + IIF(Is02(m.cod), m.s_all, 0)
    UPDATE curdatamee SET k_u=m.nk_u, ku_amb=m.nku_amb, ku_dst=m.nku_dst, ku_st=m.nku_st, ku_smp=m.nku_smp, ;
    	s_amb=m.ns_amb, s_dst=m.ns_dst, s_st=m.ns_st, s_smp=m.ns_smp, s_all=m.ns_all WHERE er_c=m.er_c 
   ELSE 
    IF SEEK(m.recid, 'curid')
     m.s_all = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    INSERT INTO curdatamee (er_c,osn230,ku_amb,ku_dst,ku_st,ku_smp,k_u,s_amb,s_dst,s_st,s_smp,s_all) VALUES ;
    	(m.er_c,m.osn230,IIF(IsPlk(m.cod),1,0),IIF(IsDst(m.cod),1,0),IIF(IsGsp(m.cod),1,0),IIF(Is02(m.cod),1,0),1,;
    	 IIF(IsPlk(m.cod),m.s_all,0),IIF(IsDst(m.cod),m.s_all,0),IIF(IsGsp(m.cod),m.s_all,0),IIF(Is02(m.cod),m.s_all,0),m.s_all)
   ENDIF 

  ENDSCAN 
  
  USE IN curid
  CREATE CURSOR curid (recid i)
  INDEX ON recid TAG recid
  SET ORDER TO recid 
  
  SELECT merror
  SCAN 
   IF DELETED()
    LOOP 
   ENDIF 
   
   m.et = et
   IF !INLIST(m.et,'4','5','6','9') && › Ãœ
    LOOP 
   ENDIF 

   m.recid  = talon.recid
   m.cod    = talon.cod
   m.er_c   = LEFT(err_mee,2)
   m.osn230 = osn230
   m.s_all  = talon.s_all
   m.ok_u   = 0
   m.os_all = 0

   IF SEEK(m.er_c, 'curdataekmp')
    m.ok_u   = curdataekmp.k_u
    m.oku_amb = curdataekmp.ku_amb
    m.oku_dst = curdataekmp.ku_dst
    m.oku_st  = curdataekmp.ku_st
    m.oku_smp = curdataekmp.ku_smp
    m.os_amb  = curdataekmp.s_amb
    m.os_dst  = curdataekmp.s_dst
    m.os_st   = curdataekmp.s_st
    m.os_smp  = curdataekmp.s_smp
    m.os_all = curdataekmp.s_all
    IF SEEK(m.recid, 'curid')
     m.s_all = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    m.nk_u   = m.ok_u + 1
    m.nku_amb = m.oku_amb + IIF(IsPlk(m.cod),1,0)
    m.nku_dst = m.oku_dst + IIF(IsDst(m.cod),1,0)
    m.nku_st  = m.oku_st  + IIF(IsGsp(m.cod),1,0)
    m.nku_smp = m.oku_smp + IIF(Is02(m.cod),1,0)
    m.ns_all = m.os_all + m.s_all
    m.ns_amb = m.os_amb + IIF(IsPlk(m.cod), m.s_all, 0)
    m.ns_dst = m.os_dst + IIF(IsDst(m.cod), m.s_all, 0)
    m.ns_st  = m.os_st  + IIF(IsGsp(m.cod), m.s_all, 0)
    m.ns_smp = m.os_smp + IIF(Is02(m.cod), m.s_all, 0)
    UPDATE curdataekmp SET k_u=m.nk_u, ku_amb=m.nku_amb, ku_dst=m.nku_dst, ku_st=m.nku_st, ku_smp=m.nku_smp, ;
    	s_amb=m.ns_amb, s_dst=m.ns_dst, s_st=m.ns_st, s_smp=m.ns_smp, s_all=m.ns_all WHERE er_c=m.er_c 
   ELSE 
    IF SEEK(m.recid, 'curid')
     m.s_all = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    INSERT INTO curdataekmp (er_c,osn230,ku_amb,ku_dst,ku_st,ku_smp,k_u,s_amb,s_dst,s_st,s_smp,s_all) VALUES ;
    	(m.er_c,m.osn230,IIF(IsPlk(m.cod),1,0),IIF(IsDst(m.cod),1,0),IIF(IsGsp(m.cod),1,0),IIF(Is02(m.cod),1,0),1,;
    	 IIF(IsPlk(m.cod),m.s_all,0),IIF(IsDst(m.cod),m.s_all,0),IIF(IsGsp(m.cod),m.s_all,0),IIF(Is02(m.cod),m.s_all,0),m.s_all)
   ENDIF 

  ENDSCAN 
  SET RELATION OFF INTO talon 

  USE IN curid
  CREATE CURSOR curid (recid i)
  INDEX ON recid TAG recid
  SET ORDER TO recid 
  
  SELECT eerror
  SET RELATION TO rid INTO talon 
  SCAN 
   IF DELETED()
    LOOP 
   ENDIF 
   
   m.recid  = talon.recid
   m.cod    = talon.cod
   m.er_c   = LEFT(c_err,2)
*   m.osn230 = osn230
   m.s_all  = talon.s_all
   m.ok_u   = 0
   m.os_all = 0

   m.oku_amb = 0 
   m.oku_dst = 0 
   m.oku_st  = 0 
   m.oku_smp = 0 

   IF SEEK(m.er_c, 'curdatamek')
    m.ok_u   = curdatamek.k_u
    m.oku_amb = curdatamek.ku_amb
    m.oku_dst = curdatamek.ku_dst
    m.oku_st  = curdatamek.ku_st
    m.oku_smp = curdatamek.ku_smp
    m.os_amb  = curdatamek.s_amb
    m.os_dst  = curdatamek.s_dst
    m.os_st   = curdatamek.s_st
    m.os_smp  = curdatamek.s_smp
    m.os_all = curdatamek.s_all
    IF SEEK(m.recid, 'curid')
     m.s_all = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    m.nk_u   = m.ok_u + 1
    m.nku_amb = m.oku_amb + IIF(IsPlk(m.cod),1,0)
    m.nku_dst = m.oku_dst + IIF(IsDst(m.cod),1,0)
    m.nku_st  = m.oku_st  + IIF(IsGsp(m.cod),1,0)
    m.nku_smp = m.oku_smp + IIF(Is02(m.cod),1,0)
    m.ns_all = m.os_all + m.s_all
    m.ns_amb = m.os_amb + IIF(IsPlk(m.cod), m.s_all, 0)
    m.ns_dst = m.os_dst + IIF(IsDst(m.cod), m.s_all, 0)
    m.ns_st  = m.os_st  + IIF(IsGsp(m.cod), m.s_all, 0)
    m.ns_smp = m.os_smp + IIF(Is02(m.cod), m.s_all, 0)
    UPDATE curdatamek SET k_u=m.nk_u, ku_amb=m.nku_amb, ku_dst=m.nku_dst, ku_st=m.nku_st, ku_smp=m.nku_smp, ;
    	s_amb=m.ns_amb, s_dst=m.ns_dst, s_st=m.ns_st, s_smp=m.ns_smp, s_all=m.ns_all WHERE er_c=m.er_c 
   ELSE 
    IF SEEK(m.recid, 'curid')
     m.s_all = 0
    ELSE 
     INSERT INTO curid FROM MEMVAR 
    ENDIF 
    INSERT INTO curdatamek (er_c,ku_amb,ku_dst,ku_st,ku_smp,k_u,s_amb,s_dst,s_st,s_smp,s_all) VALUES ;
    	(m.er_c,IIF(IsPlk(m.cod),1,0),IIF(IsDst(m.cod),1,0),IIF(IsGsp(m.cod),1,0),IIF(Is02(m.cod),1,0),1,;
    	 IIF(IsPlk(m.cod),m.s_all,0),IIF(IsDst(m.cod),m.s_all,0),IIF(IsGsp(m.cod),m.s_all,0),IIF(Is02(m.cod),m.s_all,0),m.s_all)
   ENDIF 

  ENDSCAN 

  SET RELATION OFF INTO talon 
  USE IN talon 
  USE IN merror 
  USE IN eerror
  USE IN curid

  SELECT aisoms 
  
 ENDSCAN
 USE IN aisoms
 WAIT CLEAR 

 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ Œ“◊≈“¿..." WINDOW NOWAIT 
 oexcel.Sheets(2).Select
 oSheet = oexcel.ActiveSheet
 SELECT curdatamee
 SET RELATION TO er_c INTO sookod
 m.n = 3
 m.k_u = 0
 m.s_all = 0 
 SCAN 
  WITH oExcel
   .Cells(m.n,1) = osn230
   .Cells(m.n,2) = sookod.osn230
   .Cells(m.n,3) = sookod.comment
   .Cells(m.n,4) = ku_amb
   .Cells(m.n,5) = ku_st
   .Cells(m.n,6) = ku_dst
   .Cells(m.n,7) = ku_smp
   .Cells(m.n,8) = k_u
   .Cells(m.n,9) = s_amb
   .Cells(m.n,10) = s_st
   .Cells(m.n,11) = s_dst
   .Cells(m.n,12) = s_smp
   .Cells(m.n,13) = s_all
   m.k_u = m.k_u + k_u
   m.s_all = m.s_all + s_all
  ENDWITH 
  m.n = m.n + 1
 ENDSCAN 
 oExcel.Range(oExcel.Cells(m.n,1),oExcel.Cells(m.n,3)).Merge
 oExcel.Cells(m.n,1) = '»ÚÓ„Ó:'
 oExcel.Rows(m.n).RowHeight=25
 oExcel.Cells(m.n,8) = m.k_u
 oExcel.Cells(m.n,13) = m.s_all
 oExcel.Rows(m.n).VerticalAlignment = -4108
 SET RELATION OFF INTO sookod
 USE IN curdatamee

 oexcel.Sheets(3).Select
 oSheet = oexcel.ActiveSheet
 SELECT curdataekmp
 SET RELATION TO er_c INTO sookod
 m.n = 3
 m.k_u = 0
 m.s_all = 0 
 SCAN 
  WITH oExcel
   .Cells(m.n,1) = osn230
   .Cells(m.n,2) = sookod.osn230
   .Cells(m.n,3) = sookod.comment
   .Cells(m.n,4) = ku_amb
   .Cells(m.n,5) = ku_st
   .Cells(m.n,6) = ku_dst
   .Cells(m.n,7) = ku_smp
   .Cells(m.n,8) = k_u
   .Cells(m.n,9) = s_amb
   .Cells(m.n,10) = s_st
   .Cells(m.n,11) = s_dst
   .Cells(m.n,12) = s_smp
   .Cells(m.n,13) = s_all
   m.k_u = m.k_u + k_u
   m.s_all = m.s_all + s_all
  ENDWITH 
  m.n = m.n + 1
 ENDSCAN 
 oExcel.Range(oExcel.Cells(m.n,1),oExcel.Cells(m.n,3)).Merge
 oExcel.Cells(m.n,1) = '»ÚÓ„Ó:'
 oExcel.Rows(m.n).RowHeight=25
 oExcel.Cells(m.n,8) = m.k_u
 oExcel.Cells(m.n,13) = m.s_all
 oExcel.Rows(m.n).VerticalAlignment = -4108
 SET RELATION OFF INTO sookod
 USE IN curdataekmp

 oexcel.Sheets(1).Select
 oSheet = oexcel.ActiveSheet
 SELECT curdatamek
 SET RELATION TO er_c INTO sookod
 m.n = 3
 m.k_u = 0
 m.s_all = 0 
 SCAN 
  WITH oExcel
   .Cells(m.n,1) = er_c
   .Cells(m.n,2) = sookod.osn230
   .Cells(m.n,3) = sookod.comment
   .Cells(m.n,4) = ku_amb
   .Cells(m.n,5) = ku_st
   .Cells(m.n,6) = ku_dst
   .Cells(m.n,7) = ku_smp
   .Cells(m.n,8) = k_u
   .Cells(m.n,9) = s_amb
   .Cells(m.n,10) = s_st
   .Cells(m.n,11) = s_dst
   .Cells(m.n,12) = s_smp
   .Cells(m.n,13) = s_all
   m.k_u = m.k_u + k_u
   m.s_all = m.s_all + s_all
  ENDWITH 
  m.n = m.n + 1
 ENDSCAN 
 oExcel.Range(oExcel.Cells(m.n,1),oExcel.Cells(m.n,3)).Merge
 oExcel.Cells(m.n,1) = '»ÚÓ„Ó:'
 oExcel.Rows(m.n).RowHeight=25
 oExcel.Cells(m.n,8) = m.k_u
 oExcel.Cells(m.n,13) = m.s_all
 oExcel.Rows(m.n).VerticalAlignment = -4108
 SET RELATION OFF INTO sookod
 USE IN curdatamek

 USE IN sookod
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
PROCEDURE LoadVols
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ЗАГРУЗИТЬ ОБЪЕМЫ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 f_name = 'load_vl.xlsx'
 IF !fso.FileExists(pbase+'\'+gcperiod+'\'+f_name)
  MESSAGEBOX(CHR(13)+CHR(10)+'НЕ ОБНАРУЖЕН ФАЙЛ '+f_name+CHR(13)+CHR(10)+;
   'В ДИРЕКТОРИИ '+pbase+'\'+gcperiod+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod)
  RETURN 
 ENDIF 
 IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\nsi')
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\nsi\sprlpuxx.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(m.pCommon+'\n_str_prv', 'n_str_prv', 'shar', 'n_str')>0
  USE IN sprlpu
  IF USED('n_str_prv')
   USE IN n_str_prv
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR load_vl (lpu_id n(4), mcod c(7), s_app n(15,2), n_kt n(11), s_kt n(15,2), ;
 	ks n(15,2), n_ks n(11), s_ks n(15,2), pr c(250), n_str n(11), prv c(3), n_pr n(11),;
 	ds n(15,2), n_ds n(11), ds_gem n(15,2), n_gem n(11), ds_eco n(15,2), n_eco n(11))
 INDEX on lpu_id TAG lpu_id 
 SET ORDER TO lpu_id
 
 WAIT "ЗАПУСК EXCEL..." WINDOW NOWAIT 
 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 
 m.IsVisible = .f. 
 m.IsQuit    = .t.
 oDoc = oExcel.WorkBooks.Open(pbase+'\'+gcperiod+'\'+f_name,.T.)
 
 oExcel.Sheets('АПП Новое').Select && АПП
 
 nCells=600

 FOR nCell=1 TO nCells

  IF VARTYPE(oExcel.Cells(nCell,1).Value)<>'N'
   LOOP 
  ENDIF 

  m.lpu_id = oExcel.Cells(nCell,1).Value
  m.lpu_id = IIF(m.lpu_id>9999 AND FLOOR(m.lpu_id/10000)=77, m.lpu_id%10000, m.lpu_id)

  m.mcod = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.mcod, '')
  
  m.s_app = 0
  m.n_kt  = 0
  m.s_kt  = 0
  
  IF VARTYPE(oExcel.Cells(nCell,3).Value)='N'
   m.s_app = oExcel.Cells(nCell,3).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,4).Value)='N'
   m.s_kt = oExcel.Cells(nCell,4).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,5).Value)='N'
   m.n_kt = oExcel.Cells(nCell,5).Value
  ENDIF 
  
  INSERT INTO load_vl FROM MEMVAR 
  
 NEXT 

 oExcel.Sheets('КС-новое').Select && круглосуточный стационар
 FOR nCell=1 TO nCells

  IF VARTYPE(oExcel.Cells(nCell,1).Value)<>'N'
   LOOP 
  ENDIF 

  m.lpu_id = oExcel.Cells(nCell,1).Value
  m.lpu_id = IIF(m.lpu_id>9999 AND FLOOR(m.lpu_id/10000)=77, m.lpu_id%10000, m.lpu_id)

  m.ks   = 0
  m.n_ks = 0
  m.s_ks = 0
  m.pr    = ''
  m.n_str = 0
  m.n_pr  = 0
  
  *IF VARTYPE(oExcel.Cells(nCell,3).Value)='N'
  * m.ks = oExcel.Cells(nCell,3).Value
  *ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,3).Value)='N'
   m.s_ks = oExcel.Cells(nCell,3).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,4).Value)='N'
   m.n_ks = oExcel.Cells(nCell,4).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,5).Value)='C'
   m.pr = oExcel.Cells(nCell,5).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,6).Value)='N'
   m.n_str = oExcel.Cells(nCell,6).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,7).Value)='N'
   m.n_pr = oExcel.Cells(nCell,7).Value
  ENDIF 

  m.prv = IIF(!EMPTY(m.n_str) AND SEEK(m.n_str, 'n_str_prv'), n_str_prv.prv, '')
  UPDATE load_vl SET ks=m.ks, s_ks=m.s_ks, n_ks=m.n_ks, pr=m.pr, n_str=m.n_str, n_pr=m.n_pr, prv=m.prv WHERE lpu_id = m.lpu_id
  
 NEXT 

 oExcel.Sheets('ДС-новое ').Select && дневной стационар
 FOR nCell=1 TO nCells

  IF VARTYPE(oExcel.Cells(nCell,1).Value)<>'N'
   LOOP 
  ENDIF 

  m.lpu_id = oExcel.Cells(nCell,1).Value
  m.lpu_id = IIF(m.lpu_id>9999 AND FLOOR(m.lpu_id/10000)=77, m.lpu_id%10000, m.lpu_id)

  m.ds     = 0
  m.n_ds   = 0
  m.ds_gem = 0
  m.n_gem  = 0
  m.ds_eco = 0
  m.n_eco  = 0
  
  IF VARTYPE(oExcel.Cells(nCell,3).Value)='N'
   m.ds = oExcel.Cells(nCell,3).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,4).Value)='N'
   m.n_ds = oExcel.Cells(nCell,4).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,5).Value)='N'
   m.ds_gem = oExcel.Cells(nCell,5).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,6).Value)='N'
   m.n_gem = oExcel.Cells(nCell,6).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,7).Value)='N'
   m.ds_eco = oExcel.Cells(nCell,7).Value
  ENDIF 
  IF VARTYPE(oExcel.Cells(nCell,8).Value)='N'
   m.n_eco = oExcel.Cells(nCell,8).Value
  ENDIF 

  UPDATE load_vl SET ds=m.ds, n_ds=m.n_ds, ds_gem=m.ds_gem, n_gem=m.n_gem, ;
  	ds_eco=m.ds_eco, n_eco=m.n_eco WHERE lpu_id = m.lpu_id
  
 NEXT 

 USE IN sprlpu
 USE IN n_str_prv

 WAIT CLEAR 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  IF IsQuit
   oExcel.Quit
  ENDIF 
 ENDIF 
 
 SELECT load_vl
 COPY TO &pBase\&gcPeriod\load_vl WITH cdx 
 
 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\NSI\nsif.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\NSI\nsif', 'nsif', 'shar', 'lpu_id')>0
  IF USED('nsif')
   USE IN nsif 
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT nsif 
 SET RELATION TO lpu_id INTO load_vl
 SCAN 
  REPLACE app WITH load_vl.s_app, n_kt_plan WITH load_vl.n_kt, app_ptkt WITH load_vl.s_kt,;
  	ks WITH load_vl.s_ks, n_ks_plan WITH load_vl.n_ks, ds WITH load_vl.ds, n_ds_plan WITH load_vl.n_ds,;
  	ds_gem WITH load_vl.ds_gem, n_gem WITH load_vl.n_gem, ds_eco WITH load_vl.ds_eco, ;
  	n_eco_plan WITH load_vl.n_eco, prv WITH load_vl.prv
 ENDSCAN 
 SET RELATION OFF INTO load_vl
 USE 
 USE IN load_vl
 
 MESSAGEBOX(CHR(13)+CHR(10)+'ОБРАБОТКА ЗАКОНЧЕНА!'+CHR(13)+CHR(10),0+64,'')
RETURN 
PROCEDURE SvRS7
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ СВОДНЫЕ РЕЕСТРЫ?'+CHR(13)+CHR(10)+;
 'ДЛЯ БУХГАЛТЕРИИ СОГАЗ-МЕД'+CHR(13)+CHR(10),4+32,'СОГАЗ-МЕД')=7
  RETURN 
 ENDIF 
 
 IF OpenFile(pbase+'\'+gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod') > 0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 
 IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar', 'lpu_id') > 0
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 PUBLIC oExcel AS Excel.Application
 WAIT "Запуск MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
  oExcel.Quit
 CATCH 
 ENDTRY 
 oExcel=CREATEOBJECT("Excel.Application")
 WAIT CLEAR 

 m.BookName = 'svr4'+m.qcod+PADL(DAY(DATE()),2,'0')+PADL(MONTH(DATE()),2,'0')
 m.nOpBooks = oExcel.Workbooks.Count 
 IF m.nOpBooks>0
  FOR m.nBook=1 TO m.nOpBooks
   m.cBookName = LOWER(ALLTRIM(oExcel.Workbooks.Item(m.nBook).Name))
   IF m.cBookName=m.BookName+'.xls'
    oExcel.Workbooks.Item(m.nBook).Close 
   ENDIF 
  NEXT 
 ENDIF 

 oExcel.UseSystemSeparators = .F.
 oExcel.DecimalSeparator = '.'
 oExcel.ReferenceStyle= -4150  && xlR1C1
 oExcel.SheetsInNewWorkbook = 8
 oBook = oExcel.WorkBooks.Add

 WAIT "ФОРМИРОВАНИЕ РЕЕСТРА 1..." WINDOW NOWAIT 
 =MakeHeader(1, 'РЕЕСТР 1', 'Реестр актов, предъявленных медицинскими организациями по МЭЭ')

 nRow  = 12
 nnRow = 1

 m.cl0106 = 0
 m.cl0107 = 0
 m.cl0108 = 0
 m.cl0109 = 0

 SELECT aisoms
 SET RELATION TO lpuid INTO lpudogs
 SET RELATION TO lpuid INTO sprlpu ADDITIVE 
 SCAN 
  m.fcod    = sprlpu.fcod
  m.inn     = lpudogs.inn
  m.kpp     = lpudogs.kpp
  m.lpuname = ALLTRIM(sprlpu.fullname)
  m.n_dog   = ALLTRIM(lpudogs.dogs)
  m.d_dog   = DTOC(lpudogs.ddogs)
  
  m.col0106 = aisoms.e_mee && 01 лист, 06 колонка - нумерация с 0!
  m.col0107 = aisoms.e_mee - (ROUND(aisoms.e_mee*0.35,2)+ROUND(aisoms.e_mee*0.15,2))
  m.col0108 = ROUND(aisoms.e_mee*0.35,2)
  m.col0109 = ROUND(aisoms.e_mee*0.15,2)


  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.fcod
  oExcel.Cells(nRow,3).Value  = m.inn+'/'+m.kpp
  oExcel.Cells(nRow,4).Value  = m.lpuname
  oExcel.Cells(nRow,5).Value  = m.n_dog
  oExcel.Cells(nRow,6).Value  = m.d_dog

  oExcel.Cells(nRow,7).Value   = m.col0106
  oExcel.Cells(nRow,8).Value   = m.col0107
  oExcel.Cells(nRow,9).Value   = m.col0108
  oExcel.Cells(nRow,10).Value  = m.col0109

  m.cl0106 = m.cl0106 + m.col0106
  m.cl0107 = m.cl0107 + m.col0107
  m.cl0108 = m.cl0108 + m.col0108
  m.cl0109 = m.cl0109 + m.col0109

  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 
 WAIT CLEAR 
 
 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 oExcel.Cells(nRow,7).Value  = m.cl0106
 oExcel.Cells(nRow,8).Value  = m.cl0107
 oExcel.Cells(nRow,9).Value  = m.cl0108
 oExcel.Cells(nRow,10).Value = m.cl0109

 WAIT CLEAR 
 WAIT "ФОРМИРОВАНИЕ РЕЕСТРА 2..." WINDOW NOWAIT 
 =MakeHeader(2, 'РЕЕСТР 2', 'Реестр актов, предъявленных медицинскими организациями по ЭКМП')

 nRow  = 12
 nnRow = 1

 m.cl0106 = 0
 m.cl0107 = 0
 m.cl0108 = 0
 m.cl0109 = 0

 SELECT aisoms
 SCAN 
  m.fcod    = sprlpu.fcod
  m.inn     = lpudogs.inn
  m.kpp     = lpudogs.kpp
  m.lpuname = ALLTRIM(sprlpu.fullname)
  m.n_dog   = ALLTRIM(lpudogs.dogs)
  m.d_dog   = DTOC(lpudogs.ddogs)
  
  m.col0106 = aisoms.e_ekmp && 01 лист, 06 колонка - нумерация с 0!
  m.col0107 = aisoms.e_ekmp - (ROUND(aisoms.e_ekmp*0.35,2)+ROUND(aisoms.e_ekmp*0.15,2))
  m.col0108 = ROUND(aisoms.e_ekmp*0.35,2)
  m.col0109 = ROUND(aisoms.e_ekmp*0.15,2)

  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.fcod
  oExcel.Cells(nRow,3).Value  = m.inn+'/'+m.kpp
  oExcel.Cells(nRow,4).Value  = m.lpuname
  oExcel.Cells(nRow,5).Value  = m.n_dog
  oExcel.Cells(nRow,6).Value  = m.d_dog

  oExcel.Cells(nRow,7).Value   = m.col0106
  oExcel.Cells(nRow,8).Value   = m.col0107
  oExcel.Cells(nRow,9).Value   = m.col0108
  oExcel.Cells(nRow,10).Value  = m.col0109

  m.cl0106 = m.cl0106 + m.col0106
  m.cl0107 = m.cl0107 + m.col0107
  m.cl0108 = m.cl0108 + m.col0108
  m.cl0109 = m.cl0109 + m.col0109

  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 

 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 oExcel.Cells(nRow,7).Value  = m.cl0106
 oExcel.Cells(nRow,8).Value  = m.cl0107
 oExcel.Cells(nRow,9).Value  = m.cl0108
 oExcel.Cells(nRow,10).Value = m.cl0109
 
 WAIT CLEAR 

 WAIT "ФОРМИРОВАНИЕ РЕЕСТРА 3..." WINDOW NOWAIT 
 =MakeHeader(3, 'РЕЕСТР 3', 'Реестр актов, предъявленных медицинскими организациями по МЭК')

 nRow  = 12
 nnRow = 1

 m.cl0106 = 0
 m.cl0107 = 0
 m.cl0108 = 0
 m.cl0109 = 0

 SELECT aisoms
 SCAN 
  m.fcod    = sprlpu.fcod
  m.inn     = lpudogs.inn
  m.kpp     = lpudogs.kpp
  m.lpuname = ALLTRIM(sprlpu.fullname)
  m.n_dog   = ALLTRIM(lpudogs.dogs)
  m.d_dog   = DTOC(lpudogs.ddogs)
  
  m.col0106 = aisoms.sum_flk && 01 лист, 06 колонка - нумерация с 0!
  m.col0107 = aisoms.sum_flk - (ROUND(aisoms.sum_flk*0.35,2)+ROUND(aisoms.sum_flk*0.15,2))
  m.col0108 = ROUND(aisoms.sum_flk*0.35,2)
  m.col0109 = ROUND(aisoms.sum_flk*0.15,2)

  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.fcod
  oExcel.Cells(nRow,3).Value  = m.inn+'/'+m.kpp
  oExcel.Cells(nRow,4).Value  = m.lpuname
  oExcel.Cells(nRow,5).Value  = m.n_dog
  oExcel.Cells(nRow,6).Value  = m.d_dog

  oExcel.Cells(nRow,7).Value   = m.col0106
  oExcel.Cells(nRow,8).Value   = m.col0107
  oExcel.Cells(nRow,9).Value   = m.col0108
  oExcel.Cells(nRow,10).Value  = m.col0109

  m.cl0106 = m.cl0106 + m.col0106
  m.cl0107 = m.cl0107 + m.col0107
  m.cl0108 = m.cl0108 + m.col0108
  m.cl0109 = m.cl0109 + m.col0109

  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 
 
 WAIT CLEAR 

 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 oExcel.Cells(nRow,7).Value  = m.cl0106
 oExcel.Cells(nRow,8).Value  = m.cl0107
 oExcel.Cells(nRow,9).Value  = m.cl0108
 oExcel.Cells(nRow,10).Value = m.cl0109

 WAIT "ФОРМИРОВАНИЕ РЕЕСТРА 4..." WINDOW NOWAIT 
 =MakeHeader01(4, 'РЕЕСТР 4', 'Результаты проведенной экспертизы  ТФОМС к начислению (загружается по видам медицинской помощи и периодам)')

 nRow  = 12
 nnRow = 1

 SELECT aisoms
 SCAN 
  m.fcod    = sprlpu.fcod
  m.inn     = lpudogs.inn
  m.kpp     = lpudogs.kpp
  m.lpuname = ALLTRIM(sprlpu.fullname)
  m.n_dog   = ALLTRIM(lpudogs.dogs)
  m.d_dog   = DTOC(lpudogs.ddogs)
  
  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.fcod
  oExcel.Cells(nRow,3).Value  = m.inn+'/'+m.kpp
  oExcel.Cells(nRow,4).Value  = m.lpuname
  oExcel.Cells(nRow,5).Value  = m.n_dog
  oExcel.Cells(nRow,6).Value  = m.d_dog

  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 
 
 WAIT CLEAR 

 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 WAIT "ФОРМИРОВАНИЕ РЕЕСТРА 5..." WINDOW NOWAIT 
 =MakeHeader05(5, 'РЕЕСТР 5', 'Результаты проведенной реэкспертизы ТФОМС  по невыявленным дефектам ')

 nRow  = 12
 nnRow = 1

 SELECT aisoms
 SCAN 
  m.fcod    = sprlpu.fcod
  m.inn     = lpudogs.inn
  m.kpp     = lpudogs.kpp
  m.lpuname = ALLTRIM(sprlpu.fullname)
  m.n_dog   = ALLTRIM(lpudogs.dogs)
  m.d_dog   = DTOC(lpudogs.ddogs)
  
  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.fcod
  oExcel.Cells(nRow,3).Value  = m.inn+'/'+m.kpp
  oExcel.Cells(nRow,4).Value  = m.lpuname
  oExcel.Cells(nRow,5).Value  = m.n_dog
  oExcel.Cells(nRow,6).Value  = m.d_dog

  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 
 
 WAIT CLEAR 

 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 WAIT "ФОРМИРОВАНИЕ РЕЕСТРА 6..." WINDOW NOWAIT 
 =MakeHeader06(6, 'РЕЕСТР 6', 'Результаты проведенной реэкспертизы ТФОМС  по необоснованному снятию')

 nRow  = 12
 nnRow = 1

 SELECT aisoms
 SCAN 
  m.fcod    = sprlpu.fcod
  m.inn     = lpudogs.inn
  m.kpp     = lpudogs.kpp
  m.lpuname = ALLTRIM(sprlpu.fullname)
  m.n_dog   = ALLTRIM(lpudogs.dogs)
  m.d_dog   = DTOC(lpudogs.ddogs)
  
  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.fcod
  oExcel.Cells(nRow,3).Value  = m.inn+'/'+m.kpp
  oExcel.Cells(nRow,4).Value  = m.lpuname
  oExcel.Cells(nRow,5).Value  = m.n_dog
  oExcel.Cells(nRow,6).Value  = m.d_dog

  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 
 
 WAIT CLEAR 

 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 WAIT "ФОРМИРОВАНИЕ РЕЕСТРА 7..." WINDOW NOWAIT 
 =MakeHeader07(7, 'РЕЕСТР 7', ' Сводный реестр по операциям ОМС')

 nRow  = 12
 nnRow = 1

 SELECT aisoms
 SCAN 
  m.fcod    = sprlpu.fcod
  m.inn     = lpudogs.inn
  m.kpp     = lpudogs.kpp
  m.lpuname = ALLTRIM(sprlpu.fullname)
  m.n_dog   = ALLTRIM(lpudogs.dogs)
  m.d_dog   = DTOC(lpudogs.ddogs)
  
  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.fcod
  oExcel.Cells(nRow,3).Value  = m.inn+'/'+m.kpp
  oExcel.Cells(nRow,4).Value  = m.lpuname
  oExcel.Cells(nRow,5).Value  = m.n_dog
  oExcel.Cells(nRow,6).Value  = m.d_dog

  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 
 
 WAIT CLEAR 

 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 WAIT "ФОРМИРОВАНИЕ РЕЕСТРА 8..." WINDOW NOWAIT 
 =MakeHeader08(8, 'РЕЕСТР 8', 'Реестр счетов от МО, оплаченных из средств ОМС  (загружается по видам медицинской помощи и периодам)')

 nRow  = 12
 nnRow = 1

 SELECT aisoms
 SCAN 
  m.fcod    = sprlpu.fcod
  m.inn     = lpudogs.inn
  m.kpp     = lpudogs.kpp
  m.lpuname = ALLTRIM(sprlpu.fullname)
  m.n_dog   = ALLTRIM(lpudogs.dogs)
  m.d_dog   = DTOC(lpudogs.ddogs)
  
  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.fcod
  oExcel.Cells(nRow,3).Value  = m.inn+'/'+m.kpp
  oExcel.Cells(nRow,4).Value  = m.lpuname
  oExcel.Cells(nRow,5).Value  = m.n_dog
  oExcel.Cells(nRow,6).Value  = m.d_dog

  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 
 
 WAIT CLEAR 

 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+BookName+'.xls')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+BookName+'.xls')
 ENDIF 

 oBook.SaveAs(pbase+'\'+m.gcperiod+'\'+BookName+'.xls',18)
 oExcel.Visible = .T.
 
 SELECT aisoms 
 SET RELATION OFF INTO sprlpu
 SET RELATION OFF INTO lpudogs
 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 

RETURN 

FUNCTION MakeHeader(nList, cListName, cTitleName)
 oSheet = oBook.WorkSheets(nList)
 oSheet.Select
 oSheet.Name = cListName
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oExcel.Columns(1).ColumnWidth  = 3
 oExcel.Columns(2).ColumnWidth  = 6
 oExcel.Columns(3).ColumnWidth  = 20
 oExcel.Columns(4).ColumnWidth  = 50
 oExcel.Columns(5).ColumnWidth  = 10
 oExcel.Columns(6).ColumnWidth  = 10

 oExcel.Columns(7).ColumnWidth  = 13
 oExcel.Columns(8).ColumnWidth  = 13
 oExcel.Columns(9).ColumnWidth  = 13
 oExcel.Columns(10).ColumnWidth = 13
 oExcel.Columns(11).ColumnWidth = 13
 oExcel.Columns(12).ColumnWidth = 13
 oExcel.Columns(13).ColumnWidth = 13
 oExcel.Columns(14).ColumnWidth = 13
 oExcel.Columns(15).ColumnWidth = 13
 oExcel.Columns(16).ColumnWidth = 13
 oExcel.Columns(17).ColumnWidth = 13
 oExcel.Columns(18).ColumnWidth = 13
 oExcel.Columns(19).ColumnWidth = 13
 oExcel.Columns(20).ColumnWidth = 13

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,30))
 oRange.Merge
 oExcel.Cells(1,1).Value=cTitleName

 oRange = oExcel.Range(oExcel.Cells(2,1), oExcel.Cells(2,30))
 oRange.Merge
 oExcel.Cells(2,1).Value='Период: '+LOWER(NameOfMonth(m.tmonth))+' '+STR(m.tyear,4)

 oRange = oExcel.Range(oExcel.Cells(3,1), oExcel.Cells(3,30))
 oRange.Merge
 oExcel.Cells(3,1).Value='Вид медицинской помощи: 0021'
 
 oExcel.Rows(8).VerticalAlignment = -4160
 oExcel.Rows(8).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(8,1), oExcel.Cells(10,1))
 oRange.Merge
 oExcel.Cells(8,1).Value  = '№ п\п'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,2), oExcel.Cells(10,2))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,2).Value  = 'Код МО'

 oRange = oExcel.Range(oExcel.Cells(8,3), oExcel.Cells(10,3))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,3).Value  = 'ИНН/КПП'

 oRange = oExcel.Range(oExcel.Cells(8,4), oExcel.Cells(10,4))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,4).Value  = 'Наименование МО'

 oRange = oExcel.Range(oExcel.Cells(8,5), oExcel.Cells(10,5))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,5).Value  = 'Номер Договора с МО'

 oRange = oExcel.Range(oExcel.Cells(8,6), oExcel.Cells(10,6))
 oRange.Merge
 oExcel.Cells(8,6).Value  = 'Дата Договора с МО'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,7), oExcel.Cells(8,18))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,7).Value = 'Сумма санкций  к начислению по результатам МЭЭ'
 oRange = oExcel.Range(oExcel.Cells(8,19), oExcel.Cells(8,30))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,19).Value = 'Сумма  предъявленных штрафных санкций по результатам МЭЭ'

 oExcel.Rows(9).RowHeight = 30
 oExcel.Rows(9).VerticalAlignment = -4160
 oExcel.Rows(9).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(9,7), oExcel.Cells(9,10))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,7).WrapText = .t.
 oExcel.Cells(9,7).Value = 'удержанных с МО за отчетный период'

 oRange = oExcel.Range(oExcel.Cells(9,11), oExcel.Cells(9,14))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,11).WrapText = .t.
 oExcel.Cells(9,11).Value = 'подлежащих удержанию,но НЕ удержанных с МО за отчетный период'

 oRange = oExcel.Range(oExcel.Cells(9,15), oExcel.Cells(9,18))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,15).WrapText = .t.
 oExcel.Cells(9,15).Value = 'удержанных за предыдущие периоды (из ранее не удержанных)'

 oRange = oExcel.Range(oExcel.Cells(9,19), oExcel.Cells(9,22))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,19).WrapText = .t.
 oExcel.Cells(9,19).Value = 'удержанных/ полученных в  отчетном периоде'

 oRange = oExcel.Range(oExcel.Cells(9,23), oExcel.Cells(9,26))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,23).WrapText = .t.
 oExcel.Cells(9,23).Value = 'подлежащих удержанию/получению, но НЕ удержанных/не полученных в отчетном периоде'

 oRange = oExcel.Range(oExcel.Cells(9,27), oExcel.Cells(9,30))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,27).WrapText = .t.
 oExcel.Cells(9,27).Value = 'удержанных за предыдущие периоды (из ранее не удержанных/не полученных)'
 
 FOR ncel=7 TO 30
  oRange = oExcel.Range(oExcel.Cells(10,ncel), oExcel.Cells(10,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 
 
 oExcel.Cells(10,7).Value  = 'ВСЕГО'
 oExcel.Cells(10,8).Value  = 'Целевые средства'
 oExcel.Cells(10,9).Value  = 'НСЗ'
 oExcel.Cells(10,10).Value  = 'ВД'

 oExcel.Cells(10,11).Value  = 'ВСЕГО'
 oExcel.Cells(10,12).Value  = 'Целевые средства'
 oExcel.Cells(10,13).Value  = 'НСЗ'
 oExcel.Cells(10,14).Value  = 'ВД'

 oExcel.Cells(10,15).Value  = 'ВСЕГО'
 oExcel.Cells(10,16).Value  = 'Целевые средства'
 oExcel.Cells(10,17).Value  = 'НСЗ'
 oExcel.Cells(10,18).Value  = 'ВД'

 oExcel.Cells(10,19).Value  = 'ВСЕГО'
 oExcel.Cells(10,20).Value  = 'Целевые средства'
 oExcel.Cells(10,21).Value  = 'НСЗ'
 oExcel.Cells(10,22).Value  = 'ВД'

 oExcel.Cells(10,23).Value  = 'ВСЕГО'
 oExcel.Cells(10,24).Value  = 'Целевые средства'
 oExcel.Cells(10,25).Value  = 'НСЗ'
 oExcel.Cells(10,26).Value  = 'ВД'

 oExcel.Cells(10,27).Value  = 'ВСЕГО'
 oExcel.Cells(10,28).Value  = 'Целевые средства'
 oExcel.Cells(10,29).Value  = 'НСЗ'
 oExcel.Cells(10,30).Value  = 'ВД'

 FOR ncel=1 TO 30
  oRange = oExcel.Range(oExcel.Cells(11,ncel), oExcel.Cells(11,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 

 oExcel.Cells(11,1).Value  = '1'
 oExcel.Cells(11,2).Value  = '2'
 oExcel.Cells(11,3).Value  = '3'
 oExcel.Cells(11,4).Value  = '4'
 oExcel.Cells(11,5).Value  = '5'
 oExcel.Cells(11,6).Value  = '6'
 oExcel.Cells(11,7).Value  = '7'
 oExcel.Cells(11,8).Value  = '8'
 oExcel.Cells(11,9).Value  = '9'
 oExcel.Cells(11,10).Value = '10'
 oExcel.Cells(11,11).Value = '11'
 oExcel.Cells(11,12).Value = '12'
 oExcel.Cells(11,13).Value = '13'
 oExcel.Cells(11,14).Value = '14'
 oExcel.Cells(11,15).Value = '15'
 oExcel.Cells(11,16).Value = '16'
 oExcel.Cells(11,17).Value = '17'
 oExcel.Cells(11,18).Value = '18'
 oExcel.Cells(11,19).Value = '19'
 oExcel.Cells(11,20).Value = '20'
 oExcel.Cells(11,21).Value = '21'
 oExcel.Cells(11,22).Value = '22'
 oExcel.Cells(11,23).Value = '23'
 oExcel.Cells(11,24).Value = '24'
 oExcel.Cells(11,25).Value = '25'
 oExcel.Cells(11,26).Value = '26'
 oExcel.Cells(11,27).Value = '27'
 oExcel.Cells(11,28).Value = '28'
 oExcel.Cells(11,29).Value = '29'
 oExcel.Cells(11,30).Value = '30'
RETURN 

FUNCTION MakeHeader01(nList, cListName, cTitleName)
 oSheet = oBook.WorkSheets(nList)
 oSheet.Select
 oSheet.Name = cListName
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oExcel.Columns(1).ColumnWidth  = 3
 oExcel.Columns(2).ColumnWidth  = 6
 oExcel.Columns(3).ColumnWidth  = 20
 oExcel.Columns(4).ColumnWidth  = 50
 oExcel.Columns(5).ColumnWidth  = 10
 oExcel.Columns(6).ColumnWidth  = 10

 oExcel.Columns(7).ColumnWidth  = 13
 oExcel.Columns(8).ColumnWidth  = 13
 oExcel.Columns(9).ColumnWidth  = 13
 oExcel.Columns(10).ColumnWidth = 13
 oExcel.Columns(11).ColumnWidth = 13
 oExcel.Columns(12).ColumnWidth = 13
 oExcel.Columns(13).ColumnWidth = 13
 oExcel.Columns(14).ColumnWidth = 13
 oExcel.Columns(15).ColumnWidth = 13
 oExcel.Columns(16).ColumnWidth = 13
 oExcel.Columns(17).ColumnWidth = 13
 oExcel.Columns(18).ColumnWidth = 13
 oExcel.Columns(19).ColumnWidth = 13
 oExcel.Columns(20).ColumnWidth = 13

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,30))
 oRange.Merge
 oExcel.Cells(1,1).Value=cTitleName

 oRange = oExcel.Range(oExcel.Cells(2,1), oExcel.Cells(2,30))
 oRange.Merge
 oExcel.Cells(2,1).Value='Период: '+LOWER(NameOfMonth(m.tmonth))+' '+STR(m.tyear,4)

 oRange = oExcel.Range(oExcel.Cells(3,1), oExcel.Cells(3,30))
 oRange.Merge
 oExcel.Cells(3,1).Value='Вид медицинской помощи: 0021'
 
 oExcel.Rows(8).VerticalAlignment = -4160
 oExcel.Rows(8).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(8,1), oExcel.Cells(10,1))
 oRange.Merge
 oExcel.Cells(8,1).Value  = '№ п\п'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,2), oExcel.Cells(10,2))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,2).Value  = 'Код МО'

 oRange = oExcel.Range(oExcel.Cells(8,3), oExcel.Cells(10,3))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,3).Value  = 'ИНН/КПП'

 oRange = oExcel.Range(oExcel.Cells(8,4), oExcel.Cells(10,4))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,4).Value  = 'Наименование МО'

 oRange = oExcel.Range(oExcel.Cells(8,5), oExcel.Cells(10,5))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,5).Value  = 'Номер Договора с МО'

 oRange = oExcel.Range(oExcel.Cells(8,6), oExcel.Cells(10,6))
 oRange.Merge
 oExcel.Cells(8,6).Value  = 'Дата Договора с МО'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,7), oExcel.Cells(8,12))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,7).Value = 'Результаты проведенной экспертизы ТФОМС, удержанных с МО за отчетный период'

 oRange = oExcel.Range(oExcel.Cells(8,13), oExcel.Cells(8,18))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,13).Value = 'Результаты проведенной экспертизы ТФОМС, не удержанной с МО за отчетный период'

 oRange = oExcel.Range(oExcel.Cells(8,19), oExcel.Cells(8,24))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,19).Value = 'Результаты проведенной экспертизы ТФОМС, удержанной за предыдущие периоды (из ранее не удержанных)'

 oExcel.Rows(9).RowHeight = 30
 oExcel.Rows(9).VerticalAlignment = -4160
 oExcel.Rows(9).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(9,7), oExcel.Cells(9,9))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,7).WrapText = .t.
 oExcel.Cells(9,7).Value = 'МЭЭ'

 oRange = oExcel.Range(oExcel.Cells(9,10), oExcel.Cells(9,12))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,10).WrapText = .t.
 oExcel.Cells(9,10).Value = 'ЭКМП'

 oRange = oExcel.Range(oExcel.Cells(9,13), oExcel.Cells(9,15))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,13).WrapText = .t.
 oExcel.Cells(9,13).Value = 'МЭЭ'

 oRange = oExcel.Range(oExcel.Cells(9,16), oExcel.Cells(9,18))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,16).WrapText = .t.
 oExcel.Cells(9,16).Value = 'ЭКМП'

 oRange = oExcel.Range(oExcel.Cells(9,19), oExcel.Cells(9,21))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,19).WrapText = .t.
 oExcel.Cells(9,19).Value = 'МЭЭ'

 oRange = oExcel.Range(oExcel.Cells(9,22), oExcel.Cells(9,24))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,22).WrapText = .t.
 oExcel.Cells(9,22).Value = 'ЭКМП'
 
 FOR ncel=7 TO 30
  oRange = oExcel.Range(oExcel.Cells(10,ncel), oExcel.Cells(10,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 
 
 oExcel.Cells(10,7).Value  = 'ВСЕГО'
 oExcel.Cells(10,8).Value  = 'Целевые средства'
 oExcel.Cells(10,9).Value  = 'НСЗ'

 oExcel.Cells(10,10).Value  = 'ВСЕГО'
 oExcel.Cells(10,11).Value  = 'Целевые средства'
 oExcel.Cells(10,12).Value  = 'НСЗ'

 oExcel.Cells(10,13).Value  = 'ВСЕГО'
 oExcel.Cells(10,14).Value  = 'Целевые средства'
 oExcel.Cells(10,15).Value  = 'НСЗ'

 oExcel.Cells(10,16).Value  = 'ВСЕГО'
 oExcel.Cells(10,17).Value  = 'Целевые средства'
 oExcel.Cells(10,18).Value  = 'НСЗ'

 oExcel.Cells(10,19).Value  = 'ВСЕГО'
 oExcel.Cells(10,20).Value  = 'Целевые средства'
 oExcel.Cells(10,21).Value  = 'НСЗ'

 oExcel.Cells(10,22).Value  = 'ВСЕГО'
 oExcel.Cells(10,23).Value  = 'Целевые средства'
 oExcel.Cells(10,24).Value  = 'НСЗ'

 FOR ncel=1 TO 25
  oRange = oExcel.Range(oExcel.Cells(11,ncel), oExcel.Cells(11,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 

 oExcel.Cells(11,1).Value  = '1'
 oExcel.Cells(11,2).Value  = '2'
 oExcel.Cells(11,3).Value  = '3'
 oExcel.Cells(11,4).Value  = '4'
 oExcel.Cells(11,5).Value  = '5'
 oExcel.Cells(11,6).Value  = '6'
 oExcel.Cells(11,7).Value  = '7'
 oExcel.Cells(11,8).Value  = '8'
 oExcel.Cells(11,9).Value  = '9'
 oExcel.Cells(11,10).Value = '10'
 oExcel.Cells(11,11).Value = '11'
 oExcel.Cells(11,12).Value = '12'
 oExcel.Cells(11,13).Value = '13'
 oExcel.Cells(11,14).Value = '14'
 oExcel.Cells(11,15).Value = '15'
 oExcel.Cells(11,16).Value = '16'
 oExcel.Cells(11,17).Value = '17'
 oExcel.Cells(11,18).Value = '18'
 oExcel.Cells(11,19).Value = '19'
 oExcel.Cells(11,20).Value = '20'
 oExcel.Cells(11,21).Value = '21'
 oExcel.Cells(11,22).Value = '22'
 oExcel.Cells(11,23).Value = '23'
 oExcel.Cells(11,24).Value = '24'
RETURN 

FUNCTION MakeHeader05(nList, cListName, cTitleName)
 oSheet = oBook.WorkSheets(nList)
 oSheet.Select
 oSheet.Name = cListName
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oExcel.Columns(1).ColumnWidth  = 3
 oExcel.Columns(2).ColumnWidth  = 6
 oExcel.Columns(3).ColumnWidth  = 20
 oExcel.Columns(4).ColumnWidth  = 50
 oExcel.Columns(5).ColumnWidth  = 10
 oExcel.Columns(6).ColumnWidth  = 10

 oExcel.Columns(7).ColumnWidth  = 13
 oExcel.Columns(8).ColumnWidth  = 13
 oExcel.Columns(9).ColumnWidth  = 13
 oExcel.Columns(10).ColumnWidth = 13
 oExcel.Columns(11).ColumnWidth = 13
 oExcel.Columns(12).ColumnWidth = 13
 oExcel.Columns(13).ColumnWidth = 13
 oExcel.Columns(14).ColumnWidth = 13
 oExcel.Columns(15).ColumnWidth = 13
 oExcel.Columns(16).ColumnWidth = 13
 oExcel.Columns(17).ColumnWidth = 13
 oExcel.Columns(18).ColumnWidth = 13
 oExcel.Columns(19).ColumnWidth = 13
 oExcel.Columns(20).ColumnWidth = 13

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,30))
 oRange.Merge
 oExcel.Cells(1,1).Value=cTitleName

 oRange = oExcel.Range(oExcel.Cells(2,1), oExcel.Cells(2,30))
 oRange.Merge
 oExcel.Cells(2,1).Value='Период: '+LOWER(NameOfMonth(m.tmonth))+' '+STR(m.tyear,4)

 oRange = oExcel.Range(oExcel.Cells(3,1), oExcel.Cells(3,30))
 oRange.Merge
 oExcel.Cells(3,1).Value='Вид медицинской помощи: 0021'
 
 oExcel.Rows(8).VerticalAlignment = -4160
 oExcel.Rows(8).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(8,1), oExcel.Cells(10,1))
 oRange.Merge
 oExcel.Cells(8,1).Value  = '№ п\п'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,2), oExcel.Cells(10,2))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,2).Value  = 'Код МО'

 oRange = oExcel.Range(oExcel.Cells(8,3), oExcel.Cells(10,3))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,3).Value  = 'ИНН/КПП'

 oRange = oExcel.Range(oExcel.Cells(8,4), oExcel.Cells(10,4))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,4).Value  = 'Наименование МО'

 oRange = oExcel.Range(oExcel.Cells(8,5), oExcel.Cells(10,5))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,5).Value  = 'Номер Договора с МО'

 oRange = oExcel.Range(oExcel.Cells(8,6), oExcel.Cells(10,6))
 oRange.Merge
 oExcel.Cells(8,6).Value  = 'Дата Договора с МО'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,7), oExcel.Cells(8,15))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,7).Value = 'Результаты реэкспертизы к начислению в отчетном периоде (невыявление дефектов)'

 oRange = oExcel.Range(oExcel.Cells(8,16), oExcel.Cells(8,24))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,16).Value = 'Результаты реэкспертизы, не удержанные в отчетном периоде (невыявление дефектов)'

 oRange = oExcel.Range(oExcel.Cells(8,25), oExcel.Cells(8,33))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,25).Value = 'Результаты реэкспертизы, удержанные за предыдущие периоды (из ранее не удержанных) невыявление дефектов'

 oExcel.Rows(9).RowHeight = 30
 oExcel.Rows(9).VerticalAlignment = -4160
 oExcel.Rows(9).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(9,7), oExcel.Cells(9,9))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,7).WrapText = .t.
 oExcel.Cells(9,7).Value = 'Повторный МЭК'

 oRange = oExcel.Range(oExcel.Cells(9,10), oExcel.Cells(9,12))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,10).WrapText = .t.
 oExcel.Cells(9,10).Value = 'МЭЭ'

 oRange = oExcel.Range(oExcel.Cells(9,13), oExcel.Cells(9,15))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,13).WrapText = .t.
 oExcel.Cells(9,13).Value = 'ЭКМП'

 oRange = oExcel.Range(oExcel.Cells(9,16), oExcel.Cells(9,18))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,16).WrapText = .t.
 oExcel.Cells(9,16).Value = 'Повторный МЭК'

 oRange = oExcel.Range(oExcel.Cells(9,19), oExcel.Cells(9,21))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,19).WrapText = .t.
 oExcel.Cells(9,19).Value = 'МЭЭ'

 oRange = oExcel.Range(oExcel.Cells(9,22), oExcel.Cells(9,24))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,22).WrapText = .t.
 oExcel.Cells(9,22).Value = 'ЭКМП'
 
 oRange = oExcel.Range(oExcel.Cells(9,25), oExcel.Cells(9,27))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,25).WrapText = .t.
 oExcel.Cells(9,25).Value = 'Повторный МЭК'

 oRange = oExcel.Range(oExcel.Cells(9,28), oExcel.Cells(9,30))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,28).WrapText = .t.
 oExcel.Cells(9,28).Value = 'МЭЭ'

 oRange = oExcel.Range(oExcel.Cells(9,31), oExcel.Cells(9,33))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,31).WrapText = .t.
 oExcel.Cells(9,31).Value = 'ЭКМП'

 FOR ncel=7 TO 33
  oRange = oExcel.Range(oExcel.Cells(10,ncel), oExcel.Cells(10,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 
 
 oExcel.Cells(10,7).Value  = 'ВСЕГО'
 oExcel.Cells(10,8).Value  = 'Целевые средства'
 oExcel.Cells(10,9).Value  = 'НСЗ'

 oExcel.Cells(10,10).Value  = 'ВСЕГО'
 oExcel.Cells(10,11).Value  = 'Целевые средства'
 oExcel.Cells(10,12).Value  = 'НСЗ'

 oExcel.Cells(10,13).Value  = 'ВСЕГО'
 oExcel.Cells(10,14).Value  = 'Целевые средства'
 oExcel.Cells(10,15).Value  = 'НСЗ'

 oExcel.Cells(10,16).Value  = 'ВСЕГО'
 oExcel.Cells(10,17).Value  = 'Целевые средства'
 oExcel.Cells(10,18).Value  = 'НСЗ'

 oExcel.Cells(10,19).Value  = 'ВСЕГО'
 oExcel.Cells(10,20).Value  = 'Целевые средства'
 oExcel.Cells(10,21).Value  = 'НСЗ'

 oExcel.Cells(10,22).Value  = 'ВСЕГО'
 oExcel.Cells(10,23).Value  = 'Целевые средства'
 oExcel.Cells(10,24).Value  = 'НСЗ'

 oExcel.Cells(10,25).Value  = 'ВСЕГО'
 oExcel.Cells(10,26).Value  = 'Целевые средства'
 oExcel.Cells(10,27).Value  = 'НСЗ'

 oExcel.Cells(10,28).Value  = 'ВСЕГО'
 oExcel.Cells(10,29).Value  = 'Целевые средства'
 oExcel.Cells(10,30).Value  = 'НСЗ'

 oExcel.Cells(10,31).Value  = 'ВСЕГО'
 oExcel.Cells(10,32).Value  = 'Целевые средства'
 oExcel.Cells(10,33).Value  = 'НСЗ'

 FOR ncel=1 TO 33
  oRange = oExcel.Range(oExcel.Cells(11,ncel), oExcel.Cells(11,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 

 oExcel.Cells(11,1).Value  = '1'
 oExcel.Cells(11,2).Value  = '2'
 oExcel.Cells(11,3).Value  = '3'
 oExcel.Cells(11,4).Value  = '4'
 oExcel.Cells(11,5).Value  = '5'
 oExcel.Cells(11,6).Value  = '6'
 oExcel.Cells(11,7).Value  = '7'
 oExcel.Cells(11,8).Value  = '8'
 oExcel.Cells(11,9).Value  = '9'
 oExcel.Cells(11,10).Value = '10'
 oExcel.Cells(11,11).Value = '11'
 oExcel.Cells(11,12).Value = '12'
 oExcel.Cells(11,13).Value = '13'
 oExcel.Cells(11,14).Value = '14'
 oExcel.Cells(11,15).Value = '15'
 oExcel.Cells(11,16).Value = '16'
 oExcel.Cells(11,17).Value = '17'
 oExcel.Cells(11,18).Value = '18'
 oExcel.Cells(11,19).Value = '19'
 oExcel.Cells(11,20).Value = '20'
 oExcel.Cells(11,21).Value = '21'
 oExcel.Cells(11,22).Value = '22'
 oExcel.Cells(11,23).Value = '23'
 oExcel.Cells(11,24).Value = '24'
 oExcel.Cells(11,25).Value = '25'
 oExcel.Cells(11,26).Value = '26'
 oExcel.Cells(11,27).Value = '27'
 oExcel.Cells(11,28).Value = '28'
 oExcel.Cells(11,29).Value = '29'
 oExcel.Cells(11,30).Value = '30'
 oExcel.Cells(11,31).Value = '31'
 oExcel.Cells(11,32).Value = '32'
 oExcel.Cells(11,33).Value = '33'
RETURN 

FUNCTION MakeHeader06(nList, cListName, cTitleName) && ФОормирование заголовка Реестра 6
 oSheet = oBook.WorkSheets(nList)
 oSheet.Select
 oSheet.Name = cListName
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oExcel.Columns(1).ColumnWidth  = 3
 oExcel.Columns(2).ColumnWidth  = 6
 oExcel.Columns(3).ColumnWidth  = 20
 oExcel.Columns(4).ColumnWidth  = 50
 oExcel.Columns(5).ColumnWidth  = 10
 oExcel.Columns(6).ColumnWidth  = 10

 oExcel.Columns(7).ColumnWidth  = 13
 oExcel.Columns(8).ColumnWidth  = 13
 oExcel.Columns(9).ColumnWidth  = 13
 oExcel.Columns(10).ColumnWidth = 13
 oExcel.Columns(11).ColumnWidth = 13
 oExcel.Columns(12).ColumnWidth = 13
 oExcel.Columns(13).ColumnWidth = 13
 oExcel.Columns(14).ColumnWidth = 13
 oExcel.Columns(15).ColumnWidth = 13
 oExcel.Columns(16).ColumnWidth = 13
 oExcel.Columns(17).ColumnWidth = 13
 oExcel.Columns(18).ColumnWidth = 13
 oExcel.Columns(19).ColumnWidth = 13
 oExcel.Columns(20).ColumnWidth = 13

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,35))
 oRange.Merge
 oExcel.Cells(1,1).Value=cTitleName

 oRange = oExcel.Range(oExcel.Cells(2,1), oExcel.Cells(2,35))
 oRange.Merge
 oExcel.Cells(2,1).Value='Период: '+LOWER(NameOfMonth(m.tmonth))+' '+STR(m.tyear,4)

 oRange = oExcel.Range(oExcel.Cells(3,1), oExcel.Cells(3,35))
 oRange.Merge
 oExcel.Cells(3,1).Value='Вид медицинской помощи: 0021'
 
 oExcel.Rows(8).VerticalAlignment = -4160
 oExcel.Rows(8).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(8,1), oExcel.Cells(10,1))
 oRange.Merge
 oExcel.Cells(8,1).Value  = '№ п\п'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,2), oExcel.Cells(10,2))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,2).Value  = 'Код МО'

 oRange = oExcel.Range(oExcel.Cells(8,3), oExcel.Cells(10,3))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,3).Value  = 'ИНН/КПП'

 oRange = oExcel.Range(oExcel.Cells(8,4), oExcel.Cells(10,4))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,4).Value  = 'Наименование МО'

 oRange = oExcel.Range(oExcel.Cells(8,5), oExcel.Cells(10,5))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,5).Value  = 'Номер Договора с МО'

 oRange = oExcel.Range(oExcel.Cells(8,6), oExcel.Cells(10,6))
 oRange.Merge
 oExcel.Cells(8,6).Value  = 'Дата Договора с МО'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,7), oExcel.Cells(8,17))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,7).Value = 'Результаты реэкспертизы к начислению за отчетный период (необоснованное снятие) '

 oRange = oExcel.Range(oExcel.Cells(8,18), oExcel.Cells(8,26))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,18).Value = 'Результаты реэкспертизы, не урегулированные в отчетном периоде в части целевых средств  (необоснованное снятие)'

 oRange = oExcel.Range(oExcel.Cells(8,27), oExcel.Cells(8,35))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,27).Value = 'Результаты реэкспертизы, урегулированные за предыдущие периоды (из ранее не урегулированных) необоснованное снятие'

 oExcel.Rows(9).RowHeight = 30
 oExcel.Rows(9).VerticalAlignment = -4160
 oExcel.Rows(9).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(9,7), oExcel.Cells(9,9))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,7).WrapText = .t.
 oExcel.Cells(9,7).Value = 'Повторный МЭК'

 oRange = oExcel.Range(oExcel.Cells(9,10), oExcel.Cells(9,13))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,10).WrapText = .t.
 oExcel.Cells(9,10).Value = 'МЭЭ'

 oRange = oExcel.Range(oExcel.Cells(9,14), oExcel.Cells(9,17))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,14).WrapText = .t.
 oExcel.Cells(9,14).Value = 'ЭКМП'

 oRange = oExcel.Range(oExcel.Cells(9,18), oExcel.Cells(9,20))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,18).WrapText = .t.
 oExcel.Cells(9,18).Value = 'Повторный МЭК'

 oRange = oExcel.Range(oExcel.Cells(9,21), oExcel.Cells(9,23))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,21).WrapText = .t.
 oExcel.Cells(9,21).Value = 'МЭЭ'

 oRange = oExcel.Range(oExcel.Cells(9,24), oExcel.Cells(9,26))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,24).WrapText = .t.
 oExcel.Cells(9,24).Value = 'ЭКМП'
 
 oRange = oExcel.Range(oExcel.Cells(9,27), oExcel.Cells(9,29))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,27).WrapText = .t.
 oExcel.Cells(9,27).Value = 'Повторный МЭК'

 oRange = oExcel.Range(oExcel.Cells(9,30), oExcel.Cells(9,32))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,30).WrapText = .t.
 oExcel.Cells(9,30).Value = 'МЭЭ'

 oRange = oExcel.Range(oExcel.Cells(9,33), oExcel.Cells(9,35))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(9,33).WrapText = .t.
 oExcel.Cells(9,33).Value = 'ЭКМП'

 FOR ncel=7 TO 35
  oRange = oExcel.Range(oExcel.Cells(10,ncel), oExcel.Cells(10,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 
 
 oExcel.Cells(10,7).Value  = 'ВСЕГО'
 oExcel.Cells(10,8).Value  = 'Целевые средства'
 oExcel.Cells(10,9).Value  = 'НСЗ'

 oExcel.Cells(10,10).Value  = 'ВСЕГО'
 oExcel.Cells(10,11).Value  = 'Целевые средства'
 oExcel.Cells(10,12).Value  = 'НСЗ'
 oExcel.Cells(10,13).Value  = 'РВД'

 oExcel.Cells(10,14).Value  = 'ВСЕГО'
 oExcel.Cells(10,15).Value  = 'Целевые средства'
 oExcel.Cells(10,16).Value  = 'НСЗ'
 oExcel.Cells(10,17).Value  = 'РВД'

 oExcel.Cells(10,18).Value  = 'ВСЕГО'
 oExcel.Cells(10,19).Value  = 'Целевые средства'
 oExcel.Cells(10,20).Value  = 'НСЗ'

 oExcel.Cells(10,21).Value  = 'ВСЕГО'
 oExcel.Cells(10,22).Value  = 'Целевые средства'
 oExcel.Cells(10,23).Value  = 'НСЗ'

 oExcel.Cells(10,24).Value  = 'ВСЕГО'
 oExcel.Cells(10,25).Value  = 'Целевые средства'
 oExcel.Cells(10,26).Value  = 'НСЗ'

 oExcel.Cells(10,27).Value  = 'ВСЕГО'
 oExcel.Cells(10,28).Value  = 'Целевые средства'
 oExcel.Cells(10,29).Value  = 'НСЗ'

 oExcel.Cells(10,30).Value  = 'ВСЕГО'
 oExcel.Cells(10,31).Value  = 'Целевые средства'
 oExcel.Cells(10,32).Value  = 'НСЗ'

 oExcel.Cells(10,33).Value  = 'ВСЕГО'
 oExcel.Cells(10,34).Value  = 'Целевые средства'
 oExcel.Cells(10,35).Value  = 'НСЗ'

 FOR ncel=1 TO 35
  oRange = oExcel.Range(oExcel.Cells(11,ncel), oExcel.Cells(11,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 

 oExcel.Cells(11,1).Value  = '1'
 oExcel.Cells(11,2).Value  = '2'
 oExcel.Cells(11,3).Value  = '3'
 oExcel.Cells(11,4).Value  = '4'
 oExcel.Cells(11,5).Value  = '5'
 oExcel.Cells(11,6).Value  = '6'
 oExcel.Cells(11,7).Value  = '7'
 oExcel.Cells(11,8).Value  = '8'
 oExcel.Cells(11,9).Value  = '9'
 oExcel.Cells(11,10).Value = '10'
 oExcel.Cells(11,11).Value = '11'
 oExcel.Cells(11,12).Value = '12'
 oExcel.Cells(11,13).Value = '13'
 oExcel.Cells(11,14).Value = '14'
 oExcel.Cells(11,15).Value = '15'
 oExcel.Cells(11,16).Value = '16'
 oExcel.Cells(11,17).Value = '17'
 oExcel.Cells(11,18).Value = '18'
 oExcel.Cells(11,19).Value = '19'
 oExcel.Cells(11,20).Value = '20'
 oExcel.Cells(11,21).Value = '21'
 oExcel.Cells(11,22).Value = '22'
 oExcel.Cells(11,23).Value = '23'
 oExcel.Cells(11,24).Value = '24'
 oExcel.Cells(11,25).Value = '25'
 oExcel.Cells(11,26).Value = '26'
 oExcel.Cells(11,27).Value = '27'
 oExcel.Cells(11,28).Value = '28'
 oExcel.Cells(11,29).Value = '29'
 oExcel.Cells(11,30).Value = '30'
 oExcel.Cells(11,31).Value = '31'
 oExcel.Cells(11,32).Value = '32'
 oExcel.Cells(11,33).Value = '33'
 oExcel.Cells(11,34).Value = '34'
 oExcel.Cells(11,34).Value = '35'
RETURN 

FUNCTION MakeHeader07(nList, cListName, cTitleName) && Формирование заголовка Реестра 7
 oSheet = oBook.WorkSheets(nList)
 oSheet.Select
 oSheet.Name = cListName
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oExcel.Columns(1).ColumnWidth  = 3
 oExcel.Columns(2).ColumnWidth  = 6
 oExcel.Columns(3).ColumnWidth  = 20
 oExcel.Columns(4).ColumnWidth  = 50
 oExcel.Columns(5).ColumnWidth  = 10
 oExcel.Columns(6).ColumnWidth  = 10
 oExcel.Columns(7).ColumnWidth  = 13

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,26))
 oRange.Merge
 oExcel.Cells(1,1).Value=cTitleName

 oRange = oExcel.Range(oExcel.Cells(2,1), oExcel.Cells(2,26))
 oRange.Merge
 oExcel.Cells(2,1).Value='Период: '+LOWER(NameOfMonth(m.tmonth))+' '+STR(m.tyear,4)

 oRange = oExcel.Range(oExcel.Cells(3,1), oExcel.Cells(3,26))
 oRange.Merge
 oExcel.Cells(3,1).Value='Вид медицинской помощи: 0021'
 
 oExcel.Rows(8).VerticalAlignment = -4160
 oExcel.Rows(8).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(8,1), oExcel.Cells(10,1))
 oRange.Merge
 oExcel.Cells(8,1).Value  = '№ п\п'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,2), oExcel.Cells(10,2))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,2).Value  = 'Код МО'

 oRange = oExcel.Range(oExcel.Cells(8,3), oExcel.Cells(10,3))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,3).Value  = 'ИНН/КПП'

 oRange = oExcel.Range(oExcel.Cells(8,4), oExcel.Cells(10,4))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,4).Value  = 'Наименование МО'

 oRange = oExcel.Range(oExcel.Cells(8,5), oExcel.Cells(10,5))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,5).Value  = 'Номер Договора с МО'

 oRange = oExcel.Range(oExcel.Cells(8,6), oExcel.Cells(10,6))
 oRange.Merge
 oExcel.Cells(8,6).Value  = 'Дата Договора с МО'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,7), oExcel.Cells(10,7))
 oRange.Merge
 oExcel.Cells(8,7).Value  = 'Сумма предъявленных счетов, подлежащих оплате в отчетном периоде'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,8), oExcel.Cells(10,8))
 oRange.Merge
 oExcel.Cells(8,8).Value  = 'Результаты  операций по решению комиссии ТФОМС (доплата)'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,9), oExcel.Cells(10,9))
 oRange.Merge
 oExcel.Cells(8,9).Value  = 'Сумма снятий по МЭК формируется НСЗ'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,10), oExcel.Cells(10,10))
 oRange.Merge
 oExcel.Cells(8,10).Value  = 'Сумма снятий по МЭК  не формируется НСЗ'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,11), oExcel.Cells(10,11))
 oRange.Merge
 oExcel.Cells(8,11).Value  = 'Сумма принятых счетов с учетом МЭК'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,12), oExcel.Cells(10,12))
 oRange.Merge
 oExcel.Cells(8,12).Value  = 'Сумма, удержанная по результатам   МЭЭ'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,13), oExcel.Cells(10,13))
 oRange.Merge
 oExcel.Cells(8,13).Value  = 'Сумма, удержанная по результатам   ЭКМП'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,14), oExcel.Cells(10,14))
 oRange.Merge
 oExcel.Cells(8,14).Value  = 'Сумма, удержанная по  штрафам (НЕ ПОЛУЧЕННАЯ!!!)'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,15), oExcel.Cells(8,16))
 oRange.Merge
 oExcel.Cells(8,15).Value  = 'Cуммы, удержанные по результатам экспертизы, проведенной ТФОМС'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(9,15), oExcel.Cells(10,15))
 oRange.Merge
 oExcel.Cells(9,15).Value  = 'МЭЭ'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(9,16), oExcel.Cells(10,16))
 oRange.Merge
 oExcel.Cells(9,16).Value  = 'ЭКМП'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,17), oExcel.Cells(8,22))
 oRange.Merge
 oExcel.Cells(8,17).Value  = 'Суммы, урегулированные по результатам проведенной  реэкспертизы'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(9,17), oExcel.Cells(9,19))
 oRange.Merge
 oExcel.Cells(9,17).Value  = 'невыявление дефектов'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(9,20), oExcel.Cells(9,22))
 oRange.Merge
 oExcel.Cells(9,20).Value  = 'необоснованное снятие (только целевые средства без учета средств на ведение дела)'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oExcel.Cells(10,17).Value  = 'МЭК'
 oExcel.Cells(10,18).Value  = 'МЭЭ'
 oExcel.Cells(10,19).Value  = 'ЭКМП'
 oExcel.Cells(10,20).Value  = 'МЭК(100%)'
 oExcel.Cells(10,21).Value  = 'МЭЭ(85%)'
 oExcel.Cells(10,22).Value  = 'ЭКМП(85%)'

 FOR ncel=17 TO 22
  oRange = oExcel.Range(oExcel.Cells(10,ncel), oExcel.Cells(10,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 

 oRange = oExcel.Range(oExcel.Cells(8,23), oExcel.Cells(10,23))
 oRange.Merge
 oExcel.Cells(8,23).Value  = 'Итого принято к оплате по реестрам   с учетом примененных санкций '
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,24), oExcel.Cells(10,24))
 oRange.Merge
 oExcel.Cells(8,24).Value  = 'Сумма, полученная по  штрафам на р/сч и использованная на оплату счетов МО  (50% НА ОПЛАТУ СЧЕТОВ - без нсз и рвд)'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,25), oExcel.Cells(10,25))
 oRange.Merge
 oExcel.Cells(8,25).Value  = 'Сумма, полученная по регрессам на р/сч и использованная на оплату счетов МО'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,26), oExcel.Cells(10,26))
 oRange.Merge
 oExcel.Cells(8,26).Value  = ' ВСЕГО к  бухгалтерскому учету с учетом полученных штрафов, регрессов (Дт 48203 основное)'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 


 FOR ncel=1 TO 26
  oRange = oExcel.Range(oExcel.Cells(11,ncel), oExcel.Cells(11,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 

 oExcel.Cells(11,1).Value  = '1'
 oExcel.Cells(11,2).Value  = '2'
 oExcel.Cells(11,3).Value  = '3'
 oExcel.Cells(11,4).Value  = '4'
 oExcel.Cells(11,5).Value  = '5'
 oExcel.Cells(11,6).Value  = '6'
 oExcel.Cells(11,7).Value  = '7'
 oExcel.Cells(11,8).Value  = '8'
 oExcel.Cells(11,9).Value  = '9'
 oExcel.Cells(11,10).Value = '10'
 oExcel.Cells(11,11).Value = '11'
 oExcel.Cells(11,12).Value = '12'
 oExcel.Cells(11,13).Value = '13'
 oExcel.Cells(11,14).Value = '14'
 oExcel.Cells(11,15).Value  = '15'
 oExcel.Cells(11,16).Value  = '16'
 oExcel.Cells(11,17).Value  = '17'
 oExcel.Cells(11,18).Value = '18'
 oExcel.Cells(11,19).Value = '19'
 oExcel.Cells(11,20).Value = '20'
 oExcel.Cells(11,21).Value = '21'
 oExcel.Cells(11,22).Value = '22'
 oExcel.Cells(11,23).Value = '23'
 oExcel.Cells(11,24).Value = '24'
 oExcel.Cells(11,25).Value = '25'
 oExcel.Cells(11,26).Value = '26'
RETURN 

FUNCTION MakeHeader08(nList, cListName, cTitleName) && ФОормирование заголовка Реестра 8
 oSheet = oBook.WorkSheets(nList)
 oSheet.Select
 oSheet.Name = cListName
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oExcel.Columns(1).ColumnWidth  = 3
 oExcel.Columns(2).ColumnWidth  = 6
 oExcel.Columns(3).ColumnWidth  = 20
 oExcel.Columns(4).ColumnWidth  = 50
 oExcel.Columns(5).ColumnWidth  = 10
 oExcel.Columns(6).ColumnWidth  = 10
 oExcel.Columns(7).ColumnWidth  = 13

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,7))
 oRange.Merge
 oExcel.Cells(1,1).Value=cTitleName

 oRange = oExcel.Range(oExcel.Cells(2,1), oExcel.Cells(2,7))
 oRange.Merge
 oExcel.Cells(2,1).Value='Период: '+LOWER(NameOfMonth(m.tmonth))+' '+STR(m.tyear,4)

 oRange = oExcel.Range(oExcel.Cells(3,1), oExcel.Cells(3,7))
 oRange.Merge
 oExcel.Cells(3,1).Value='Вид медицинской помощи: 0021'
 
 oExcel.Rows(8).VerticalAlignment = -4160
 oExcel.Rows(8).WrapText = .t.

 oRange = oExcel.Range(oExcel.Cells(8,1), oExcel.Cells(10,1))
 oRange.Merge
 oExcel.Cells(8,1).Value  = '№ п\п'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,2), oExcel.Cells(10,2))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,2).Value  = 'Код МО'

 oRange = oExcel.Range(oExcel.Cells(8,3), oExcel.Cells(10,3))
 oRange.Merge
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oExcel.Cells(8,3).Value  = 'ИНН/КПП'

 oRange = oExcel.Range(oExcel.Cells(8,4), oExcel.Cells(10,4))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,4).Value  = 'Наименование МО'

 oRange = oExcel.Range(oExcel.Cells(8,5), oExcel.Cells(10,5))
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 
 oRange.Merge
 oExcel.Cells(8,5).Value  = 'Номер Договора с МО'

 oRange = oExcel.Range(oExcel.Cells(8,6), oExcel.Cells(10,6))
 oRange.Merge
 oExcel.Cells(8,6).Value  = 'Дата Договора с МО'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 oRange = oExcel.Range(oExcel.Cells(8,7), oExcel.Cells(10,7))
 oRange.Merge
 oExcel.Cells(8,7).Value  = 'Сумма'
 WITH oRange
 .Borders(07).LineStyle = 1 
 .Borders(07).Weight    = 3
 .Borders(08).LineStyle = 1 
 .Borders(08).Weight    = 3
 .Borders(09).LineStyle = 1 
 .Borders(09).Weight    = 3
 .Borders(10).LineStyle = 1 
 .Borders(10).Weight    = 3
 ENDWITH 

 FOR ncel=1 TO 7
  oRange = oExcel.Range(oExcel.Cells(11,ncel), oExcel.Cells(11,ncel))
  WITH oRange
  .Borders(07).LineStyle = 1 
  .Borders(07).Weight    = 3
  .Borders(08).LineStyle = 1 
  .Borders(08).Weight    = 3
  .Borders(09).LineStyle = 1 
  .Borders(09).Weight    = 3
  .Borders(10).LineStyle = 1 
  .Borders(10).Weight    = 3
  ENDWITH 
 ENDFOR 

 oExcel.Cells(11,1).Value  = '1'
 oExcel.Cells(11,2).Value  = '2'
 oExcel.Cells(11,3).Value  = '3'
 oExcel.Cells(11,4).Value  = '4'
 oExcel.Cells(11,5).Value  = '5'
 oExcel.Cells(11,6).Value  = '6'
 oExcel.Cells(11,7).Value  = '7'
RETURN 
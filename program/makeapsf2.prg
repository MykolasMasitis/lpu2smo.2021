PROCEDURE MakeAPSF2

IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ'+CHR(13)+CHR(10)+;
 'АКТ ПЕРЕДАЧИ СЧЕТОВ-ФАКТУР?'+CHR(13)+CHR(10))==7
 RETURN 
ENDIF 

*IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar', 'mcod')>0
IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar', 'lpu_id')>0
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 RETURN 
ENDIF 
IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\admokrxx', 'admokr', 'shar', 'cokr')>0
 IF USED('admokr')
  USE IN admokr
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 RETURN 
ENDIF 

IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
 IF USED('admokr')
  USE IN admokr
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 RETURN 
ENDIF 

IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'lpu_id')>0
 IF USED('pilot')
  USE IN pilot
 ENDIF 
 IF USED('admokr')
  USE IN admokr
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 RETURN .f.
ENDIF 

IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilots', 'pilots', 'shar', 'lpu_id')>0
 IF USED('pilot')
  USE IN pilot
 ENDIF 
 IF USED('pilots')
  USE IN pilots
 ENDIF 
 IF USED('admokr')
  USE IN admokr
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 RETURN .f.
ENDIF 

IF OpenFile(pbase+'\'+gcperiod+'\aisoms', 'aisoms', 'shar', 'mcod')>0
 IF USED('pilot')
  USE IN pilot
 ENDIF 
 IF USED('pilots')
  USE IN pilots
 ENDIF 
 IF USED('lpudogs')
  USE IN lpudogs
 ENDIF 
 IF USED('admokr')
  USE IN admokr
 ENDIF 
 IF USED('sprlpu')
  USE IN sprlpu
 ENDIF 
 IF USED('aisoms')
  USE IN aisoms
 ENDIF 
 RETURN 
ENDIF 

SELECT sprlpu
SET RELATION TO mcod INTO aisoms
*SET RELATION TO mcod INTO lpudogs

PUBLIC oExcel AS Excel.Application

WAIT "Запуск MS Excel..." WINDOW NOWAIT 
TRY 
 oExcel=GETOBJECT(,"Excel.Application")
CATCH 
 oExcel=CREATEOBJECT("Excel.Application")
ENDTRY 
WAIT CLEAR 

WAIT "ФОРМИРОВАНИЕ ОТЧЕТА..." WINDOW NOWAIT 

oexcel.UseSystemSeparators = .F.
oexcel.DecimalSeparator = '.'

oexcel.ReferenceStyle= -4150  && xlR1C1
 
oexcel.SheetsInNewWorkbook=1
oBook = oExcel.WorkBooks.Add
oexcel.Cells.Font.Name='Calibri'
oexcel.Cells.Font.Size=9
oexcel.ActiveSheet.PageSetup.Orientation=2

BookName = pout+'\APSF_'+gcperiod
oSheet = oBook.WorkSheets(1)
oSheet.Select
 
nCell = 1

WITH oExcel
 .Columns(01).ColumnWidth = 3
 .Columns(02).ColumnWidth = 20
 .Columns(03).ColumnWidth = 5
 .Columns(04).ColumnWidth = 7
 .Columns(05).ColumnWidth = 10
 .Columns(06).ColumnWidth = 10
 .Columns(07).ColumnWidth = 07
 .Columns(08).ColumnWidth = 07
 .Columns(09).ColumnWidth = 07
 .Columns(10).ColumnWidth = 07
 .Columns(11).ColumnWidth = 09
 .Columns(12).ColumnWidth = 09
 .Columns(13).ColumnWidth = 10
 .Columns(14).ColumnWidth = 10
 .Columns(15).ColumnWidth = 10
 .Columns(16).ColumnWidth = 5
 .Columns(17).ColumnWidth = 7
 .Columns(18).ColumnWidth = 20
 .Columns(19).ColumnWidth = 5
 .Columns(20).ColumnWidth = 20
 .Columns(21).ColumnWidth = 10
 .Columns(22).ColumnWidth = 15
 .Columns(23).ColumnWidth = 19
 .Columns(24).ColumnWidth = 19

 .Columns(02).NumberFormat = '@'
 .Columns(02).WrapText     = .t.
 .Columns(03).NumberFormat = '@'
 .Columns(04).NumberFormat = '@'
 .Columns(05).NumberFormat = '@'
 .Columns(06).NumberFormat = "#,##0.00"
 .Columns(07).NumberFormat = "#,##0.00"
 .Columns(08).NumberFormat = "#,##0.00"
 .Columns(09).NumberFormat = "#,##0.00"
 .Columns(10).NumberFormat = "#,##0.00"
 .Columns(11).NumberFormat = "#,##0.00"
 .Columns(12).NumberFormat = "#,##0.00"
 .Columns(13).NumberFormat = "#,##0.00"
 .Columns(14).NumberFormat = "#,##0.00"
 .Columns(15).NumberFormat = "#,##0.00"
 .Columns(16).NumberFormat = '@'
 .Columns(17).NumberFormat = '@'
 .Columns(18).NumberFormat = '@'
 .Columns(19).NumberFormat = '@'
 .Columns(20).NumberFormat = '@'
 .Columns(21).NumberFormat = '@'
 .Columns(22).NumberFormat = '@'
 .Columns(23).NumberFormat = '@'
 .Columns(24).NumberFormat = '@'
ENDWITH 

WITH oExcel.Sheets(1)
 .cells(1,1).Value2 = 'АКТ передачи счетов-фактур'
 .cells(2,1).Value2 = 'за медицинские услуги по Московской городской программе ОМС (по условиям оказания медицинcкой помощи)'
 .cells(3,1).Value2 = 'от СМО '+m.qname
 .cells(4,1).Value2 = 'за '+ NameOfMonth(tMonth)+ ' '+STR(tYear,4)+' года'
 .cells(1,1).Font.Size = 11
 .cells(2,1).Font.Size = 11
 .cells(3,1).Font.Size = 11
 .cells(4,1).Font.Size = 11
 .cells(1,1).Font.Bold = .T.
 .cells(2,1).Font.Bold = .T.
 .cells(3,1).Font.Bold = .T.
 .cells(4,1).Font.Bold = .T.
 .cells(1,1).HorizontalAlignment=-4108
 .cells(2,1).HorizontalAlignment=-4108
 .cells(3,1).HorizontalAlignment=-4108
 .cells(4,1).HorizontalAlignment=-4108
 FOR nRow=1 TO 4
  oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,24))
  oRange.Merge
 ENDFOR  
 
ENDWITH 

m.fstring = 7 && Первая строка

WITH oExcel.Sheets(1)
 .Cells(m.fstring,01).Value2 = '№ п\п'
 .Cells(m.fstring,02).Value2 = 'Наименование ЛПУ (юридического лица) в разрезе по районам'
 .Cells(m.fstring,03).Value2 = 'Округ'
 .Cells(m.fstring,04).Value2 = 'Номер счетов-фактур, представленных ЛПУ'
 .Cells(m.fstring,05).Value2 = 'Номер Договора с ЛПУ'
 .Cells(m.fstring,06).Value2 = 'Сумма счетов-фактур, представленных ЛПУ'
 .Cells(m.fstring,07).Value2 = 'МЭК'
 .Cells(m.fstring,08).Value2 = 'в т.ч. PPA'
 .Cells(m.fstring,09).Value2 = 'МЭЭ'
 .Cells(m.fstring,10).Value2 = 'ЭКМП'
 .Cells(m.fstring,11).Value2 = 'Аванс'
 .Cells(m.fstring,12).Value2 = 'Долг на начало периода'
 .Cells(m.fstring,13).Value2 = 'К начислению'
 .Cells(m.fstring,14).Value2 = 'К оплате'
 .Cells(m.fstring,15).Value2 = 'Долг на конец периода'
 .Cells(m.fstring,16).Value2 = 'Пилот'
 .Cells(m.fstring,17).Value2 = 'ФКОД'
 .Cells(m.fstring,18).Value2 = 'ИНН/КПП'
 .Cells(m.fstring,19).Value2 = 'ПилотC'
 .Cells(m.fstring,20).Value2 = 'Лицевой счет'
 .Cells(m.fstring,21).Value2 = 'Дата договора'
 .Cells(m.fstring,22).Value2 = 'Лицевой счет в получателе'
 .Cells(m.fstring,23).Value2 = 'Номер банковского счета'
 .Cells(m.fstring,24).Value2 = 'Номер договора на 2020 г.'

 oExcel.Range(oExcel.Cells(m.fstring-1,1), oExcel.Cells(m.fstring,1)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,2), oExcel.Cells(m.fstring,2)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,3), oExcel.Cells(m.fstring,3)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,4), oExcel.Cells(m.fstring,4)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,5), oExcel.Cells(m.fstring,5)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,6), oExcel.Cells(m.fstring,6)).Merge 

 oExcel.Range(oExcel.Cells(m.fstring-1,11), oExcel.Cells(m.fstring,11)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,12), oExcel.Cells(m.fstring,12)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,13), oExcel.Cells(m.fstring,13)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,14), oExcel.Cells(m.fstring,14)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,15), oExcel.Cells(m.fstring,15)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,16), oExcel.Cells(m.fstring,16)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,17), oExcel.Cells(m.fstring,17)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,18), oExcel.Cells(m.fstring,18)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,19), oExcel.Cells(m.fstring,19)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,20), oExcel.Cells(m.fstring,20)).Merge 

 oExcel.Range(oExcel.Cells(m.fstring-1,21), oExcel.Cells(m.fstring,21)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,22), oExcel.Cells(m.fstring,22)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,23), oExcel.Cells(m.fstring,23)).Merge 
 oExcel.Range(oExcel.Cells(m.fstring-1,24), oExcel.Cells(m.fstring,24)).Merge 

 .Cells(m.fstring-1,8).Value2 = 'Сумма уменьшения счетов-фактур'
 oExcel.Range(oExcel.Cells(m.fstring-1,7), oExcel.Cells(m.fstring-1,10)).Merge 
ENDWITH 

FOR ncol=1 TO 24
 WITH oExcel.Sheets(1)
  .Cells(m.fstring,ncol).Font.Size = 8
  .Cells(m.fstring,ncol).Font.Bold = .F.
  .Cells(m.fstring,ncol).WrapText = .t.
  .Cells(m.fstring,ncol).HorizontalAlignment = -4108
  .Cells(m.fstring,ncol).VerticalAlignment = -4108
 ENDWITH 
NEXT 

m.fstring = m.fstring + 1
FOR ncolumn=1 TO 24
 WITH oExcel.Sheets(1)
  .cells(m.fstring,ncolumn).Value2 = STR(ncolumn,2)
  .cells(m.fstring,ncolumn).HorizontalAlignment = -4108
 ENDWITH 
NEXT 

oExcel.Range(oExcel.Cells(m.fstring,1), oExcel.Cells(m.fstring,14)).NumberFormat='@'

m.ccokr = cokr
m.nincokr = 1

m.sum04 = 0
m.sum05 = 0
m.sum06 = 0
m.sum07 = 0
m.sum08 = 0
m.sum09 = 0
m.sum10 = 0
m.sum11 = 0
m.sum12 = 0
m.sum13 = 0
m.sum14 = 0

m.sum04okr = 0
m.sum05okr = 0
m.sum06okr = 0
m.sum07okr = 0
m.sum08okr = 0
m.sum09okr = 0
m.sum10okr = 0
m.sum11okr = 0
m.sum12okr = 0
m.sum13okr = 0
m.sum14okr = 0

SCAN
 m.fstring = m.fstring + 1
 m.mcod = mcod
 m.lpu_id = lpu_id
 m.llpuid = lpu_id
 m.cokr = cokr
 m.numdog = ''
 IF USED('lpudogs')
  *m.numdog = IIF(SEEK(m.mcod, 'lpudogs'), lpudogs.dogs, '')
  *m.numdog = ALLTRIM(m.numdog)
  *m.inn    = IIF(SEEK(m.mcod, 'lpudogs'), lpudogs.inn, '')
  *m.kpp    = IIF(SEEK(m.mcod, 'lpudogs'), lpudogs.kpp, '')
  *m.acc    = IIF(SEEK(m.mcod, 'lpudogs'), lpudogs.account, '')
  *m.ddogs = IIF(SEEK(m.mcod, 'lpudogs'), DTOC(lpudogs.ddogs), '')
  *m.old_acc = IIF(SEEK(m.mcod, 'lpudogs'), ALLTRIM(lpudogs.old_acc), '')
  *m.bank_acc = IIF(SEEK(m.mcod, 'lpudogs'), ALLTRIM(lpudogs.bank_acc), '')
  *m.dog2020 = IIF(SEEK(m.mcod, 'lpudogs'), ALLTRIM(lpudogs.dog2020), '')

  m.numdog = IIF(SEEK(m.lpu_id, 'lpudogs'), lpudogs.dogs, '')
  m.numdog = ALLTRIM(m.numdog)
  m.inn    = IIF(SEEK(m.lpu_id, 'lpudogs'), lpudogs.inn, '')
  m.kpp    = IIF(SEEK(m.lpu_id, 'lpudogs'), lpudogs.kpp, '')
  m.acc    = IIF(SEEK(m.lpu_id, 'lpudogs'), lpudogs.account, '')
  m.ddogs = IIF(SEEK(m.lpu_id, 'lpudogs'), DTOC(lpudogs.ddogs), '')
  m.old_acc = IIF(SEEK(m.lpu_id, 'lpudogs'), ALLTRIM(lpudogs.old_acc), '')
  m.bank_acc = IIF(SEEK(m.lpu_id, 'lpudogs'), ALLTRIM(lpudogs.bank_acc), '')
  m.dog2020 = IIF(SEEK(m.lpu_id, 'lpudogs'), ALLTRIM(lpudogs.dog2020), '')
 ENDIF 
 m.cokrname = IIF(SEEK(m.cokr, 'admokr'), admokr.name_okr, '')
 m.ispilot = IIF(SEEK(m.llpuid, 'pilot'), .t., .f.)
 m.ispilots = IIF(SEEK(m.llpuid, 'pilots'), .t., .f.)

 WITH oExcel.Sheets(1)
  .Cells(m.fstring,01).Value2 = m.nincokr
  .Cells(m.fstring,02).Value2 = ALLTRIM(fullname)
  .Cells(m.fstring,03).Value2 = m.cokrname
  .Cells(m.fstring,04).Value2 = mcod
  .Cells(m.fstring,05).Value2 = LEFT(m.numdog,9)
  .Cells(m.fstring,06).Value2 = aisoms.s_pred + IIF(FIELD('s_lek', 'aisoms')='S_LEK', aisoms.s_lek, 0)
  .Cells(m.fstring,07).Value2 = (aisoms.sum_flk+aisoms.ls_flk)
  .Cells(m.fstring,08).Value2 = IIF(FIELD('s_ppa','aisoms')='S_PPA', aisoms.s_ppa, 0)
  .Cells(m.fstring,09).Value2 = aisoms.e_mee
  .Cells(m.fstring,10).Value2 = aisoms.e_ekmp
  .Cells(m.fstring,11).Value2 = aisoms.s_avans
  .Cells(m.fstring,12).Value2 = aisoms.dolg_b
  .Cells(m.fstring,13).Value2 = aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b
  .Cells(m.fstring,14).Value2 = IIF(aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b>0, aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b, 0)
  .Cells(m.fstring,15).Value2 = IIF(aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b<0, -(aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b), 0)
  .Cells(m.fstring,16).Value2 = IIF(m.IsPilot, 'Да ', 'Нет')
  .Cells(m.fstring,17).Value2 = fcod
  .Cells(m.fstring,18).Value2 = m.inn+'/'+m.kpp
  .Cells(m.fstring,19).Value2 = IIF(m.IsPilots, 'Да ', 'Нет')
  .Cells(m.fstring,20).Value2 = m.acc
  .Cells(m.fstring,21).Value2 = m.ddogs
  
  .Cells(m.fstring,22).Value2 = m.old_acc
  .Cells(m.fstring,23).Value2 = m.bank_acc
  .Cells(m.fstring,24).Value2 = m.dog2020
  
 ENDWITH 

 m.sum04okr = m.sum04okr + aisoms.s_pred
 m.sum05okr = m.sum05okr
 m.sum06okr = m.sum06okr + (aisoms.sum_flk+aisoms.ls_flk)
 m.sum07okr = m.sum07okr + aisoms.e_mee
 m.sum08okr = m.sum08okr + aisoms.e_ekmp
 m.sum09okr = m.sum09okr + aisoms.s_avans
 m.sum10okr = m.sum10okr + aisoms.dolg_b
 m.sum11okr = m.sum11okr + (aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b)
 m.sum12okr = m.sum12okr + IIF(aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b>0, aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b, 0)
 m.sum13okr = m.sum13okr + IIF(aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b<0, -(aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b), 0)
 m.sum14okr = m.sum14okr + IIF(FIELD('s_ppa','aisoms')='S_PPA', aisoms.s_ppa, 0)

 m.sum04 = m.sum04 + aisoms.s_pred
 m.sum05 = m.sum05
 m.sum06 = m.sum06 + (aisoms.sum_flk+aisoms.ls_flk)
 m.sum07 = m.sum07 + aisoms.e_mee
 m.sum08 = m.sum08 + aisoms.e_ekmp
 m.sum09 = m.sum09 + aisoms.s_avans
 m.sum10 = m.sum10 + aisoms.dolg_b
 m.sum11 = m.sum11 + aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b
 m.sum12 = m.sum12 + IIF(aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) -  ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b>0, aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b, 0)
 m.sum13 = m.sum13 + IIF(aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b<0, -(aisoms.s_pred - (aisoms.sum_flk+aisoms.ls_flk) - ;
   (aisoms.e_mee + aisoms.e_ekmp) - aisoms.s_avans - aisoms.dolg_b), 0)
 m.sum14 = m.sum14 + IIF(FIELD('s_ppa','aisoms')='S_PPA', aisoms.s_ppa, 0)

 m.nincokr = m.nincokr + 1

 IF cokr != m.ccokr
  WITH oExcel.Sheets(1)
  ENDWITH 
  m.cokrsum = 0
  m.sum04okr = 0
  m.sum05okr = 0
  m.sum06okr = 0
  m.sum07okr = 0
  m.sum08okr = 0
  m.sum09okr = 0
  m.sum10okr = 0
  m.sum11okr = 0
  m.sum12okr = 0
  m.sum13okr = 0
  m.sum14okr = 0
 ENDIF 
ENDSCAN 

m.fstring = m.fstring + 1
WITH oExcel.Sheets(1)
  oExcel.Range(oExcel.Cells(m.fstring,2), oExcel.Cells(m.fstring,3)).Merge 
 .Cells(m.fstring,02).Value2 = 'Итого:'

 .Cells(m.fstring,06).Value2 = m.sum04
 .Cells(m.fstring,07).Value2 = m.sum06
 .Cells(m.fstring,08).Value2 = m.sum14
 .Cells(m.fstring,09).Value2 = m.sum07
 .Cells(m.fstring,10).Value2 = m.sum08
 .Cells(m.fstring,11).Value2 = m.sum09
 .Cells(m.fstring,12).Value2 = m.sum10
 .Cells(m.fstring,13).Value2 = m.sum11
 .Cells(m.fstring,14).Value2 = m.sum12
 .Cells(m.fstring,15).Value2 = m.sum13
ENDWITH 

IF fso.FileExists(pout+'\APSF_'+gcperiod+'.xls')
 fso.DeleteFile(pout+'\APSF_'+gcperiod+'.xls')
ENDIF 

oBook.SaveAs(BookName,18)

oExcel.Visible = .T.

SET RELATION OFF INTO aisoms
USE IN aisoms
USE IN sprlpu
USE IN admokr
USE IN pilot 
USE IN pilots
IF USED('lpudogs')
 USE IN lpudogs
ENDIF 

WAIT CLEAR 

RETURN 

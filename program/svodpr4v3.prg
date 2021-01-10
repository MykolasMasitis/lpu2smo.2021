PROCEDURE SvodPr4V3
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ СФОРМИРОВАТЬ СВОД ПРИЛОЖЕНИЯ 4?'+CHR(13)+CHR(10)+;
 '(ВАРИАНТ 3)'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\pr4.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ОТСУТСТВУЕТ ФАЙЛ PR4.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\FormMAG02.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'НЕ СФОРМИРОВАН МАГ02!'+CHR(13)+CHR(10),0+16,'')
  RETURN
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\pr4', 'pr4', 'shar')>0
  IF USED('pr4')
   USE IN pr4
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\nsi\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')>0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  USE IN pr4
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar', 'mcod') > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  USE IN pr4
  USE IN sprlpu
  RETURN 
 ENDIF 
 
 IF OpenFile(pcommon+'\lpudogs', 'lpudogs', 'shar', 'lpu_id') > 0
  IF USED('lpudogs')
   USE IN lpudogs
  ENDIF 
  USE IN pr4
  USE IN sprlpu
  USE IN aisoms
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\'+m.gcPeriod+'\FormMAG02', 'mag', 'shar', 'lpuid') > 0
  IF USED('mag')
   USE IN mag
  ENDIF 
  USE IN pr4
  USE IN sprlpu
  USE IN aisoms
  USE IN lpudogs
  RETURN 
 ENDIF 

 PUBLIC oExcel AS Excel.Application
 WAIT "Запуск MS Excel..." WINDOW NOWAIT 
 TRY 
  oExcel=GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel=CREATEOBJECT("Excel.Application")
 ENDTRY 
 WAIT CLEAR 

 WAIT "ФОРМИРОВАНИЕ ОТЧЕТА..." WINDOW NOWAIT 

 oExcel.UseSystemSeparators = .F.
 oExcel.DecimalSeparator = '.'
 oExcel.ReferenceStyle= -4150  && xlR1C1
 oExcel.SheetsInNewWorkbook = 1
 oBook = oExcel.WorkBooks.Add
 oSheet = oBook.WorkSheets(1)
 oSheet.Select
 oSheet.Name = 'Сводная'
 oexcel.Cells.Font.Name='Calibri'
 oexcel.ActiveSheet.PageSetup.Orientation=2
 oExcel.Cells.NumberFormat = '@'

 oRange = oExcel.Range(oExcel.Cells(1,1), oExcel.Cells(1,19))
 oRange.Merge
 oExcel.Cells(1,1).Value='Сводный акт об оплате расчетов по подушевому финансированию'
 oExcel.Cells(1,1).HorizontalAlignment = -4108
 oExcel.Cells(1,1).Font.Size = 12
 oExcel.Cells(1,1).Font.Bold = .F.
 oExcel.Cells(1,1).Font.Italic = .T.
 oExcel.Rows(1).RowHeight = 30
 oExcel.Rows(1).VerticalAlignment = -4108
 
 oExcel.Columns(1).ColumnWidth  = 3
 oExcel.Columns(2).ColumnWidth  = 21
 oExcel.Columns(3).ColumnWidth  = 8
 oExcel.Columns(4).ColumnWidth  = 12
 oExcel.Columns(5).ColumnWidth  = 12
 oExcel.Columns(6).ColumnWidth  = 18
 oExcel.Columns(7).ColumnWidth  = 18
 oExcel.Columns(8).ColumnWidth  = 18
 oExcel.Columns(9).ColumnWidth  = 18
 oExcel.Columns(10).ColumnWidth  = 18
 oExcel.Columns(11).ColumnWidth = 18
 oExcel.Columns(12).ColumnWidth = 18
 oExcel.Columns(13).ColumnWidth = 18
 oExcel.Columns(14).ColumnWidth = 18
 oExcel.Columns(15).ColumnWidth = 18
 oExcel.Columns(16).ColumnWidth = 18
 oExcel.Columns(17).ColumnWidth = 18
 oExcel.Columns(18).ColumnWidth = 18
 oExcel.Columns(19).ColumnWidth = 18
 oExcel.Columns(20).ColumnWidth = 18
 oExcel.Columns(21).ColumnWidth = 18
 oExcel.Columns(22).ColumnWidth = 18
 oExcel.Columns(23).ColumnWidth = 18
 oExcel.Columns(24).ColumnWidth = 25
 oExcel.Columns(25).ColumnWidth = 25
 oExcel.Columns(26).ColumnWidth = 10
 
 oExcel.Cells(2,1).Value  = '№ п\п'
 oExcel.Cells(2,2).Value  = 'Наименование ЛПУ (юридического лица) в разрезе по районам'
 oExcel.Cells(2,3).Value  = 'Округ'
 oExcel.Cells(2,4).Value  = 'Код МО'
 oExcel.Cells(2,5).Value  = 'ФКОД'
 oExcel.Cells(2,6).Value  = 'Кол-во прикрепленного населения'
 oExcel.Cells(2,7).Value  = 'МЭК'
 oExcel.Cells(2,8).Value  = 'Подушевой МЭК'
 oExcel.Cells(2,9).Value  = 'Внеподушевой МЭК'
 oExcel.Cells(2,10).Value  = 'Расчетный объем подушевого финаснирования'
 oExcel.Cells(2,11).Value  = 'АВАНС всего'
 oExcel.Cells(2,12).Value  = 'АВАНС подушевик'
 oExcel.Cells(2,13).Value = 'Сумма средств, подлежащих исключению из финансирования по ;
  результам взаиморасчета с другими ГП за пацентов, прикрепленных к данному ГП, но;
  получивших амбулаторную медицинскую помощь в других ГП'
 oExcel.Cells(2,14).Value = 'Сумма средств, полученных на пациентов, прикрепленных ;
  к другим ГП, но в отчетном периоде пролеченных в данном ГП'
 oExcel.Cells(2,15).Value  = 'Сумма дополнительных средств,полученных ГП за комплексные услуги ;
  профилактического направления и средств за оказанную медицинскую помощь в дневных ;
  стационарах'
 oExcel.Cells(2,16).Value  = 'Сумма средств за медицинскую помощь, оказанную в данном ГП ;
  гражданам, не прикрепленным к городским поликлиникам'
 oExcel.Cells(2,17).Value = 'Сумма средств по результатам проведения МЭЭ и ЭКМП'
 oExcel.Cells(2,18).Value = 'Сумма средств по результатам проведения МЭЭ'
 oExcel.Cells(2,19).Value = 'Сумма средств по результатам проведения ЭКМП'
 oExcel.Cells(2,20).Value = 'ИТОГО сумма средств к перечислению ;
  (гр.6+гр.7+гр.8+гр.9-гр.10'
 oExcel.Cells(2,21).Value = 'Стационар с учетом МЭК'
 oExcel.Cells(2,22).Value = 'Доплата за ЛС'
 oExcel.Cells(2,23).Value = 'ИТОГО К ОПЛАТЕ, с учетом стационара и доп. услуг,ЛС'
 oExcel.Cells(2,24).Value = 'ИНН/КПП'
 oExcel.Cells(2,25).Value = 'Лицевой счет'
 oExcel.Cells(2,26).Value = 'Дата договора'

 oExcel.Rows(2).RowHeight = 130
 oExcel.Rows(2).HorizontalAlignment = 1
 oExcel.Rows(2).VerticalAlignment = -4160
 oExcel.Rows(2).WrapText = .t.

 oExcel.Cells(3,1).Value  = '1'
 oExcel.Cells(3,2).Value  = '2'
 oExcel.Cells(3,3).Value  = '3'
 oExcel.Cells(3,4).Value  = '4'
 oExcel.Cells(3,5).Value  = '5'
 oExcel.Cells(3,6).Value  = '6'
 oExcel.Cells(3,7).Value  = '7'
 oExcel.Cells(3,8).Value  = '8'
 oExcel.Cells(3,9).Value  = '9'
 oExcel.Cells(3,10).Value = '10'
 oExcel.Cells(3,11).Value = '11'
 oExcel.Cells(3,12).Value = '12'
 oExcel.Cells(3,13).Value = '13'
 oExcel.Cells(3,14).Value = '14'
 oExcel.Cells(3,15).Value = '15'
 oExcel.Cells(3,16).Value = '16'
 oExcel.Cells(3,17).Value = '17'
 oExcel.Cells(3,18).Value = '18'
 oExcel.Cells(3,19).Value = '19'
 oExcel.Cells(3,20).Value = '20'
 oExcel.Cells(3,21).Value = '21'
 oExcel.Cells(3,22).Value = '22'
 oExcel.Cells(3,23).Value = '23'
 oExcel.Cells(3,24).Value = '24'
 oExcel.Cells(3,25).Value = '25'
 oExcel.Cells(3,26).Value = '26'

 nRow  = 4
 nnRow = 1

 m.col5  = 0
 m.col6  = 0
 m.col7  = 0
 m.col8  = 0
 m.col9  = 0
 m.col10 = 0
 m.col11 = 0
 m.col12 = 0
 m.col13 = 0
 m.col14 = 0
 m.col15 = 0
 m.col16 = 0
 m.col17 = 0
 m.col18 = 0
 m.col19 = 0
 m.col20 = 0
 m.col21 = 0
 m.col22 = 0
 m.col23 = 0

 SELECT pr4
 SET RELATION TO mcod INTO aisoms 
 SET RELATION TO lpuid INTO mag ADDITIVE 
 SCAN 
  SCATTER MEMVAR 
  m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
  m.cokr = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.cokr, '')
  m.lpudog = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.dogs), '')
  m.inn = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.inn), '')
  m.kpp = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.kpp), '')
  m.acc = IIF(SEEK(m.lpuid, 'lpudogs'), ALLTRIM(lpudogs.account), '')
  m.fcod = IIF(SEEK(m.lpuid, 'sprlpu'), sprlpu.fcod, '')
  m.ddogs = IIF(SEEK(m.lpuid, 'lpudogs'), DTOC(lpudogs.ddogs), '')

  m.str01 = finval
  m.str02 = aisoms.s_avans 
  m.str31 = s_others
  m.str32 = s_guests
  m.str33 = s_own - m.str02
  m.str03 = m.str02 - m.str31 + m.str32 + m.str33
  *m.str04 = aisoms.s_dop
  m.str04 = IIF(INLIST(m.qcod,'S7','R2'), mag.col23, mag.col09) && mag.col23
  m.str05 = m.str01 - IIF(m.str03<0, -1*m.str03, m.str03)
  m.str06 = s_npilot
  m.str07 = s_empty
  m.str08  = aisoms.e_mee+aisoms.e_ekmp
  *m.str09 = IIF(m.str05>=0, m.str05, 0)+m.str04+m.str06+m.str07
  m.koplate = m.str01 - m.str02 - m.str31 + m.str32 + m.str04 + m.str06 + m.str07 - m.str08

  oExcel.Cells(nRow,1).Value  = nnRow
  oExcel.Cells(nRow,2).Value  = m.lpuname
  oExcel.Cells(nRow,3).Value  = m.cokr
  oExcel.Cells(nRow,4).Value  = m.mcod
  oExcel.Cells(nRow,5).Value  = m.fcod
  oExcel.Cells(nRow,6).Value  = pazval
  oExcel.Cells(nRow,7).Value  = aisoms.sum_flk && mag.col16 + mag.col20
  *oExcel.Cells(nRow,8).Value  = aisoms.pf_flk
  oExcel.Cells(nRow,8).Value  = mag.col16 && mag.col14
  oExcel.Cells(nRow,9).Value  = mag.col20 && mag.col17
  oExcel.Cells(nRow,10).Value  = finval
  oExcel.Cells(nRow,11).Value = aisoms.s_avans
  oExcel.Cells(nRow,12).Value = aisoms.s_pr_avans
  oExcel.Cells(nRow,13).Value = m.str31
  oExcel.Cells(nRow,14).Value = m.str32
  oExcel.Cells(nRow,15).Value = m.str04
  oExcel.Cells(nRow,16).Value = m.str07 
  oExcel.Cells(nRow,17).Value = m.str08
  oExcel.Cells(nRow,18).Value = aisoms.e_mee
  oExcel.Cells(nRow,19).Value = aisoms.e_ekmp
  oExcel.Cells(nRow,20).Value = finval - aisoms.s_pr_avans - m.str31 + m.str32 +m.str07 - m.str08
  *oExcel.Cells(nRow,21).Value = mag.col22 + IIF(FIELD('s_lek', 'aisoms')='S_LEK', aisoms.s_lek, 0) && mag.col07
  oExcel.Cells(nRow,21).Value = mag.col22 && mag.col07
  oExcel.Cells(nRow,22).Value = IIF(FIELD('s_lek', 'aisoms')='S_LEK', aisoms.s_lek, 0)
  *oExcel.Cells(nRow,23).Value = finval - aisoms.s_pr_avans - m.str31 + m.str32 + m.str04 + m.str07 - m.str08 + ;
  	mag.col08 + IIF(FIELD('s_lek', 'aisoms')='S_LEK' , aisoms.s_lek, 0) - mag.col20 && гр 9-гр 10-гр 12+ гр 13+гр 14+гр 15-гр 16+гр 20+гр 21
  oExcel.Cells(nRow,23).Value = finval - aisoms.s_avans - m.str31 + m.str32 + m.str04 + m.str07 - m.str08 + ;
  	mag.col08 - mag.col20 && гр 9-гр 10-гр 12+ гр 13+гр 14+гр 15-гр 16+гр 20+гр 21
  oExcel.Cells(nRow,24).Value = m.inn+'/'+m.kpp
  oExcel.Cells(nRow,25).Value = m.acc
  oExcel.Cells(nRow,26).Value = m.ddogs
  
  m.col6  = m.col6 + pazval
  m.col7  = m.col7 + aisoms.sum_flk && mag.col16 + mag.col20
  *m.col8  = m.col8 + aisoms.pf_flk
  m.col8  = m.col8 + mag.col16 && mag.col14
  m.col9  = m.col9 + mag.col20 && mag.col17
  m.col10 = m.col10 + finval
  m.col11 = m.col11 + aisoms.s_avans
  m.col12 = m.col12 + aisoms.s_pr_avans
  m.col13 = m.col13 + m.str31
  m.col14 = m.col14 + m.str32
  m.col15 = m.col15 + m.str04
  m.col16 = m.col16 + m.str07
  m.col17 = m.col17 + m.str08
  m.col18 = m.col18 + aisoms.e_mee
  m.col19 = m.col19 + aisoms.e_ekmp
  m.col20 = m.col20 + (finval - aisoms.s_pr_avans - m.str31 + m.str32 +m.str07 - m.str08)
  *m.col21 = m.col21 + mag.col22 + IIF(FIELD('s_lek', 'aisoms')='S_LEK', aisoms.s_lek, 0) && mag.col07
  m.col21 = m.col21 + mag.col22 && mag.col07
  m.col22 = m.col22 + IIF(FIELD('s_lek', 'aisoms')='S_LEK', aisoms.s_lek, 0)
  *m.col23 = m.col23 + (finval - aisoms.s_pr_avans - m.str31 + m.str32 + m.str04 + m.str07 - m.str08 + ;
  	mag.col08 + IIF(FIELD('s_lek', 'aisoms')='S_LEK' , aisoms.s_lek, 0) - mag.col20)
  m.col23 = m.col23 + (finval - aisoms.s_avans - m.str31 + m.str32 + m.str04 + m.str07 - m.str08 + ;
  	mag.col08 - mag.col20)
  
  nRow  = nRow + 1
  nnRow = nnRow + 1
  
 ENDSCAN 
 SET RELATION OFF INTO aisoms
 SET RELATION OFF INTO mag 
 
 USE IN pr4
 USE IN sprlpu
 USE IN aisoms
 USE IN lpudogs
 USE IN mag 
 
 WAIT CLEAR 
 
 oRange = oExcel.Range(oExcel.Cells(nRow,1), oExcel.Cells(nRow,5))
 oRange.Merge
 oExcel.Cells(nRow,1).Value='Итого:'

 oExcel.Cells(nRow,6).Value  = m.col6
 oExcel.Cells(nRow,7).Value  = m.col7
 oExcel.Cells(nRow,8).Value  = m.col8
 oExcel.Cells(nRow,9).Value  = m.col9 
 oExcel.Cells(nRow,10).Value  = m.col10
 oExcel.Cells(nRow,11).Value  = m.col11
 oExcel.Cells(nRow,12).Value  = m.col12
 oExcel.Cells(nRow,13).Value = m.col13
 oExcel.Cells(nRow,14).Value = m.col14
 oExcel.Cells(nRow,15).Value = m.col15
 oExcel.Cells(nRow,16).Value = m.col16
 oExcel.Cells(nRow,17).Value = m.col17
 oExcel.Cells(nRow,18).Value = m.col18
 oExcel.Cells(nRow,19).Value = m.col19
 oExcel.Cells(nRow,20).Value = m.col20
 oExcel.Cells(nRow,21).Value = m.col21
 oExcel.Cells(nRow,22).Value = m.col22
 oExcel.Cells(nRow,23).Value = m.col23

 BookName = 'svpr4v3'+m.qcod+PADL(DAY(DATE()),2,'0')+PADL(MONTH(DATE()),2,'0')
 IF fso.FileExists(pbase+'\'+m.gcperiod+'\'+BookName+'.xls')
  fso.DeleteFile(pbase+'\'+m.gcperiod+'\'+BookName+'.xls')
 ENDIF 

 oBook.SaveAs(pbase+'\'+m.gcperiod+'\'+BookName+'.xls',18)
 oExcel.Visible = .T.
 
RETURN 
FUNCTION McPrn(para1, IsVisible, IsQuit)
 
 m.lcpath  = para1
 m.mcod    = mcod
 m.lpuid   = lpuid
 IF USED('sprlpu')
  m.lpu_name = IIF(SEEK(m.lpuid, 'sprlpu'), ALLTRIM(sprlpu.fullname), '')
 ELSE 
  m.lpu_name = ''
 ENDIF 
 
 IF OpFiles()>0
  =ClFiles()
  RETURN 
 ENDIF 

 m.mmy    = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)

 m.n_akt    = mcod+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 *m.d_akt    = DTOC(DATE()) &&  && 6-ой рабочий день
 m.d_akt = DTOC(goApp.d_acts) && 6-ой рабочий день
 m.akt_mon  = NameOfMonth(tMonth)
 m.akt_year = STR(tYear,4)
 m.nr_akt   = m.mcod+m.qcod+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),1)
 
 m.k_predst = 0
 m.s_lek    = 0
 m.s_predst = 0
 m.k_bad    = 0
 m.s_bad    = 0
 m.s_ok     = 0
 m.s_532    = s_532
 m.s_532st  = IIF(FIELD('s_532st', 'aisoms')=UPPER('s_532st'), aisoms.s_532st, 0)
 m.s_532app = IIF(FIELD('s_532app')=UPPER('s_532app'), s_532app, 0)
 m.s_532pet = IIF(FIELD('s_532pet')=UPPER('s_532pet'), s_532pet, 0)
 m.s_532dst = IIF(FIELD('s_532dst')=UPPER('s_532dst'), s_532dst, 0)
 m.s_532eco = IIF(FIELD('s_532eco')=UPPER('s_532eco'), s_532eco, 0)
 
 SELECT talon
 SET RELATION TO recid INTO serror
 SCAN 
  m.k_u   = k_u
  m.s_all = s_all
  m.cod   = cod
  m.k_predst = m.k_predst + IIF(IsMes(m.cod) OR IsVMP(m.cod), 1, m.k_u)
  m.s_lek    = m.s_lek + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
  m.s_predst = m.s_predst + m.s_all + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
  IF EMPTY(sError.rid)
   m.s_ok = m.s_ok + m.s_all + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
  ELSE 
   m.k_bad = m.k_bad + 1 && m.k_u
   m.s_bad = m.s_bad + m.s_all + IIF(FIELD('s_lek')='S_LEK', s_lek, 0)
  ENDIF 
 ENDSCAN 
 SET RELATION OFF INTO serror
 
 =ClFiles()

 LOCAL m.lcTmpName, m.lcRepName, m.lcDbfName, m.llResult
 m.lcTmpName = pTempl + "\McxxxxQQmmy.xls"
 m.lcRepName = lcPath + "\Mc" + STR(m.lpuid,4) + m.qcod + m.mmy+'.xls'
 m.lcRepName2 = lcPath + "\Mc" + STR(m.lpuid,4) + m.qcod + m.mmy

 CREATE CURSOR curdata (recid i, cod n(6), name c(10))
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 USE IN curdata 

 TRY 
  oExcel = GETOBJECT(,"Excel.Application")
 CATCH 
  oExcel = CREATEOBJECT("Excel.Application")
 ENDTRY 
 
 oExcel.DisplayAlerts=.f.
 
 IF fso.FileExists(m.lcRepName2+'.pdf')
  fso.DeleteFile(m.lcRepName2+'.pdf')
 ENDIF 
 oDoc = oExcel.Workbooks.Add(m.lcRepName)
 
 oSheet = oDoc.ActiveSheet
 IF m.s_532eco>0
  osheet.Rows("33").Insert(-4121, 1) && xlDown, xlFormatFromRightOrBelow
  oExcel.Cells(33,2).Value = " МП в  условиях дневного стационара ЭКО:"
  oExcel.Range(oExcel.Cells(33,2), oExcel.Cells(33,7)).Merge
  oExcel.Range(oExcel.Cells(33,8), oExcel.Cells(33,9)).Merge
  oExcel.Cells(33,8).Value = m.s_532eco
 ENDIF 
 IF m.s_532dst>0
  osheet.Rows("33").Insert(-4121, 1) && xlDown, xlFormatFromRightOrBelow
  oExcel.Cells(33,2).Value = " МП в  условиях дневного стационара без учета ЭКО:"
  oExcel.Range(oExcel.Cells(33,2), oExcel.Cells(33,7)).Merge
  oExcel.Range(oExcel.Cells(33,8), oExcel.Cells(33,9)).Merge
  oExcel.Cells(33,8).Value = m.s_532dst
 ENDIF 
 IF m.s_532st>0
  osheet.Rows("33").Insert(-4121, 1) && xlDown, xlFormatFromRightOrBelow
  oExcel.Cells(33,2).Value = " МП в стационарных условиях:"
  oExcel.Range(oExcel.Cells(33,2), oExcel.Cells(33,7)).Merge
  oExcel.Range(oExcel.Cells(33,8), oExcel.Cells(33,9)).Merge
  oExcel.Cells(33,8).Value = m.s_532st
 ENDIF 
 IF m.s_532pet>0
  osheet.Rows("33").Insert(-4121, 1) && xlDown, xlFormatFromRightOrBelow
  oExcel.Cells(33,2).Value = " МП в амбулаторных условиях ПЭТ/КТ:"
  oExcel.Range(oExcel.Cells(33,2), oExcel.Cells(33,7)).Merge
  oExcel.Range(oExcel.Cells(33,8), oExcel.Cells(33,9)).Merge
  oExcel.Cells(33,8).Value = m.s_532pet
 ENDIF 
 IF m.s_532app>0
  osheet.Rows("33").Insert(-4121, 1) && xlDown, xlFormatFromRightOrBelow
  oExcel.Cells(33,2).Value = "в том числе: МП в амбулаторных условиях без учета ПЭТ/КТ:"
  oExcel.Range(oExcel.Cells(33,2), oExcel.Cells(33,7)).Merge
  oExcel.Range(oExcel.Cells(33,8), oExcel.Cells(33,9)).Merge
  oExcel.Cells(33,8).Value = m.s_532app
 ENDIF 
 
 odoc.SaveAs(m.lcRepName2,56)
 TRY 
  odoc.SaveAs(m.lcRepName2,57)
 CATCH 
 ENDTRY 

 SELECT aisoms

RETURN 

FUNCTION OpFiles
 tn_rslt = 0 
 tn_rslt = tn_rslt + OpenFile(lcpath+'\talon', 'talon', 'shar')
 tn_rslt = tn_rslt + OpenFile(lcpath+'\e'+m.mcod, 'serror', 'shar', 'rid')
RETURN tn_rslt

FUNCTION ClFiles
 IF USED('talon')
  USE IN talon
 ENDIF 
 IF USED('serror')
  USE IN serror
 ENDIF 
RETURN 
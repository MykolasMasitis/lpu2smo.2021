PROCEDURE mkAppFP(para1, para2, para3)
 LOCAL m.mcod, m.lpuid, m.IsVisible, m.IsQuit
 
 m.oPrm      = para1
 m.IsVisible = para2
 m.IsQuit    = para3
 
 WITH oPrm
  m.pBase    = .pBase
  m.pTempl   = .pTempl
  m.gcPeriod = .gcPeriod
  m.qCod     = .qCod
  m.qName    = .qname
 
  m.tMonth   = .tMonth
  m.tYear    = .tYear

  m.mcod     = .mcod
  m.lpuid    = .lpuid
 ENDWITH 

 *m.IsLpuTpn  = IIF(SEEK(m.lpuid, 'lputpn'), .t., .f.)
 m.IsLpuTpn  = oPrm.IsLpuTpn
 *m.IsHorLpuS = IIF(SEEK(m.lpuid, 'horlpus'), .t., .f.)
 m.DotName = m.pTempl+'\ActPFn.xlt'
 m.mmy    = SUBSTR(m.gcperiod,5,2) + SUBSTR(m.gcperiod,4,1)
 *DocName = pbase+'\'+m.gcPeriod+'\'+m.mcod+'\pdf'+m.qcod+m.mmy
 *m.DocName = m.pbase+'\'+m.gcPeriod+'\'+m.mcod+'\pdf'+m.mcod+m.mmy
 DocName = pbase+'\'+m.gcperiod+'\'+m.mcod+'\pdf'+m.qcod+m.mmy
 
 IF !fso.FileExists(dotname)
  *MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ ƒŒ ”Ã≈Õ“¿ ActPFn.xlt'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
* WAIT "«¿œ”—  EXCEL..." WINDOW NOWAIT 
* TRY 
*  oExcel=GETOBJECT(,"Excel.Application")
* CATCH 
*  oExcel=CREATEOBJECT("Excel.Application")
* ENDTRY 
* WAIT CLEAR 
 
 Local Worker as Worker
 TRY 
  Worker = NewObject("Worker", "ParallelFox.vcx")
 CATCH 
 ENDTRY 

 IF VARTYPE(Worker)='O'
  Worker.StartCriticalSection("XReport")
 ENDIF 
 oDoc = oExcel.WorkBooks.Add(dotname)
 IF VARTYPE(Worker)='O'
  Worker.EndCriticalSection("XReport") 
 ENDIF 

 IF VARTYPE(Worker)='O'
  RELEASE m.Worker
 ENDIF  

 m.lpuname = oPrm.lpuname
 *m.lpuname = IIF(SEEK(m.lpuid, 'sprlpu'), ;
  ALLTRIM(sprlpu.fullname)+', '+m.mcod+' ('+CokrName(VAL(sprlpu.cokr))+')', '')
 *IF m.IsLpuTpn
 * m.s_avans = IIF(SEEK(m.mcod, 'aisoms'), aisoms.s_pr_avans, 0)
 *ELSE 
 * m.s_avans = IIF(SEEK(m.mcod, 'aisoms'), aisoms.s_avans, 0)
 *ENDIF 
 m.s_avans = oPrm.s_avans
 
 *m.e_mee   = IIF(SEEK(m.mcod, 'aisoms'), aisoms.e_mee, 0)
 *m.e_ekmp  = IIF(SEEK(m.mcod, 'aisoms'), aisoms.e_ekmp, 0)
 m.e_mee  = oPrm.e_mee
 m.e_ekmp = oPrm.e_ekmp

 *m.str01  = finval
 *m.str011 = ppr4.finval
 *m.str02  = m.s_avans
 *m.str03  = s_others
 *m.str031 = ppr4.s_others
 *m.str04  = s_guests
 *m.str041 = ppr4.s_guests && + IIF(m.IsHorLpuS AND SEEK(m.lpuid, 'mag01'), mag01.str41, 0)
 *m.str051 = s_npilot + s_empty
 *m.str052 = ppr4.s_npilot + ppr4.s_empty && +IIF(m.IsHorLpuS AND SEEK(m.lpuid, 'mag01'), mag01.str52, 0)
 *m.str06  = m.e_mee+m.e_ekmp
 
 m.str01  = oPrm.str01
 m.str011 = oPrm.str011
 m.str02  = oPrm.str02
 m.str03  = oPrm.str03
 m.str031 = oPrm.str031
 m.str04  = oPrm.str04
 m.str041 = oPrm.str041
 m.str051 = oPrm.str051
 m.str052 = oPrm.str052
 m.str06  = oPrm.str06

 m.koplate = (m.str01+m.str011) - m.str02 - (m.str03+m.str031) + (m.str04+m.str041) + (m.str051+m.str052) - m.str06

 *m.str10 = aisoms.s_pred - aisoms.sum_flk
 m.str10 = oPrm.str10
 
 WITH oExcel
  .Cells(4,1) = m.lpuname
  .Cells(7,1)='ÒÓ —ÃŒ '+ALLTRIM(m.qname)
  .Cells(9,1)='Á‡  '+NameOfMonth(INT(VAL(SUBSTR(m.gcPeriod,5,2))))+' '+SUBSTR(m.gcPeriod,1,4)+' „.'

  .Cells(12,8).Value = m.str01 + m.str011
  .Cells(13,8).Value = m.str01
  .Cells(14,8).Value = m.str011

  .Cells(15,8).Value = m.str02

  .Cells(16,8).Value = m.str03 + m.str031

  .Cells(17,8).Value = m.str04 + m.str041

  .Cells(18,8).Value = m.str051 + m.str052
  .Cells(19,8).Value = m.str051
  .Cells(20,8).Value = m.str052

  .Cells(21,8).Value = m.str06

 ENDWITH 

 oExcel.Cells(22,1)='»ÚÓ„Ó Í ÓÔÎ‡ÚÂ: (Ô.1-Ô.2-Ô.3+Ô.4+Ô.5-Ô.6)'+;
  TRANSFORM(m.koplate, '99 999 999.99')+' Û·.'
 oExcel.Cells(24,1)='(Ò Û˜ÂÚÓÏ ÂÁÛÎ¸Ú‡ÚÓ‚ Ã› ) '+;
  TRANSFORM(m.str10,'99999999.99')+' Û·.'

 IF fso.FileExists(docname+'.xls')
  fso.DeleteFile(docname+'.xls')
 ENDIF 
 oDoc.SaveAs(DocName, 18)
 TRY 
  IF fso.FileExists(docname+'.pdf')
   fso.DeleteFile(docname+'.pdf')
  ENDIF 
  oDoc.SaveAs(DocName, 57)
 CATCH 
 FINALLY 
*  oDoc.Close(0)
 ENDTRY 
 
 IF IsVisible == .t. 
  oExcel.Visible = .t.
 ELSE 
  oDoc.Close(0)
  IF IsQuit
   *oExcel.Quit
  ENDIF 
 ENDIF 
 
RETURN 
PROCEDURE AllMt
 PUBLIC oExcel as Excel.Application

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

* TRY 
*  oExcel = GETOBJECT(,"Excel.Application")
* CATCH 
*  oExcel = CREATEOBJECT("Excel.Application")
* ENDTRY 

 m.mmy = PADL(m.tMonth,2,'0') + SUBSTR(STR(m.tYear,4),4,1)

 SCAN 
  MailView.refresh

  IF !fso.FolderExists(pbase+'\'+m.gcperiod+'\'+mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\people.dbf') OR ;
   !fso.FileExists(pbase+'\'+m.gcperiod+'\'+mcod+'\talon.dbf')
   LOOP 
  ENDIF 
  
  m.mcod  = mcod 
  m.lpuid = lpuid

  m.l_path  = m.pbase+'\'+m.gcperiod+'\'+m.mcod
  m.mt_file = "Mt" + STR(m.lpuid,4) + m.qcod + m.mmy

  IF fso.FileExists(m.l_path+'\'+m.mt_file+'.pdf')
   LOOP 
  ENDIF 

  m.t_beg = SECONDS()
  =MtPrn2(pbase+'\'+m.gcperiod+'\'+mcod, .f., .f.)
  m.t_end = SECONDS()
  *m.t_last = ROUND((m.t_end - m.t_beg)/60,2)
  m.t_mt = m.t_end - m.t_beg

  UPDATE aisoms SET t_mt=m.t_mt WHERE mcod=m.mcod

  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 

 ENDSCAN 


 SET ESCAPE &OldEscStatus

 oExcel.Quit
 
* MESSAGEBOX(TRANSFORM(m.t_last,'99999.99'),0+64,'')
 
RETURN 
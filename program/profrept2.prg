PROCEDURE ProfRepT2
 IF MESSAGEBOX('ÑÂÅÄÅÍÈß ÎÁ ÎÁÚÅÌÀÕ È ÑÒÎÈÌÎÑÒÈ'+CHR(13)+CHR(10)+;
 	'ÏÐÎÔÌÅÄÎÑÌÎÒÐÀÕ ÂÇÐÎÑËÎÃÎ ÍÀÑÅËÅÍÈß?'+CHR(13)+CHR(10)+'',4+64, 'ÒÀÁËÈÖÀ 2')=7
 	RETURN 
 ENDIF 
 IF !fso.FileExists(pTempl+'\Prof_t2.xls')
  MESSAGEBOX('ØÀÁËÎÍ '+UPPER('Prof_t2.xls')+' ÍÅ ÍÀÉÄÅÍ!',4+64,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pcommon+'\dspcodes.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË DSPCODES.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pBase+'\'+m.gcPeriod+'\dsp.dbf')
  MESSAGEBOX('ÔÀÉË dsp.dbf ÇÀ ÂÛÁÐÀÍÍÛÉ ÏÅÐÈÎÎÄ ÍÅ ÑÔÎÐÌÈÐÎÂÀÍ!', 4+64, '')
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+m.gcPeriod+'\dsp', 'dsp', 'shar')>0
  IF USED('dsp')
   USE IN dsp 
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pcommon+'\dspcodes', 'dspcodes', 'shar', 'cod')>0
  IF USED('dspcodes')
   USE IN dspcodes
  ENDIF 
  IF USED('dsp')
   USE IN dsp
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR curdsp (sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 DIMENSION dimdsp(11,17)
 dimdsp = 0
 DIMENSION dimdata(3,3)
 dimdata = 0
 DIMENSION dimtb4(11,20)
 dimtb4 = 0

 SELECT dsp 
 SCAN 
  m.cod = cod 
  m.rslt = rslt
  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF m.tipofcod!=2
   LOOP 
  ENDIF 
  IF !INLIST(m.rslt,343,344,345) && âìåñòî 345 373/374
*   LOOP 
  ENDIF 

  m.vozr = ROUND((m.tdat1 - dr)/365.25,2)

  IF m.vozr>=18 AND w=1
   dimtb4(2,1) = dimtb4(2,1) + 1
   dimtb4(2,2) = dimtb4(2,2) + s_all
   IF EMPTY(er)
    dimtb4(2,3) = dimtb4(2,3) + 1
    dimtb4(2,4) = dimtb4(2,4) + s_all
   ENDIF
   DO CASE 
    CASE m.rslt = 343 
     dimtb4(2,5) = dimtb4(2,5) + 1
     dimtb4(2,6) = dimtb4(2,6) + s_all
     IF EMPTY(er)
      dimtb4(2,7) = dimtb4(2,7) + 1
      dimtb4(2,8) = dimtb4(2,8) + s_all
     ENDIF
    CASE m.rslt = 344
     dimtb4(2,9) = dimtb4(2,9) + 1
     dimtb4(2,10) = dimtb4(2,10) + s_all
     IF EMPTY(er)
      dimtb4(2,11) = dimtb4(2,11) + 1
      dimtb4(2,12) = dimtb4(2,12) + s_all
     ENDIF
    CASE m.rslt = 345
     dimtb4(2,13) = dimtb4(2,13) + 1
     dimtb4(2,14) = dimtb4(2,14) + s_all
     IF EMPTY(er)
      dimtb4(2,15) = dimtb4(2,15) + 1
      dimtb4(2,16) = dimtb4(2,16) + s_all
     ENDIF
   ENDCASE 
  ENDIF 

  IF m.vozr>=18 and w=2
   dimtb4(3,1) = dimtb4(3,1) + 1
   dimtb4(3,2) = dimtb4(3,2) + s_all
   IF EMPTY(er)
    dimtb4(3,3) = dimtb4(3,3) + 1
    dimtb4(3,4) = dimtb4(3,4) + s_all
   ENDIF 
   DO CASE 
    CASE m.rslt = 343 
     dimtb4(3,5) = dimtb4(3,5) + 1
     dimtb4(3,6) = dimtb4(3,6) + s_all
     IF EMPTY(er)
      dimtb4(3,7) = dimtb4(3,7) + 1
      dimtb4(3,8) = dimtb4(3,8) + s_all
     ENDIF
    CASE m.rslt = 344
     dimtb4(3,9) = dimtb4(3,9) + 1
     dimtb4(3,10) = dimtb4(3,10) + s_all
     IF EMPTY(er)
      dimtb4(3,11) = dimtb4(3,11) + 1
      dimtb4(3,12) = dimtb4(3,12) + s_all
     ENDIF
    CASE m.rslt = 345
    dimtb4(3,13) = dimtb4(3,13) + 1
    dimtb4(3,14) = dimtb4(3,14) + s_all
    IF EMPTY(er)
     dimtb4(3,15) = dimtb4(3,15) + 1
     dimtb4(3,16) = dimtb4(3,16) + s_all
    ENDIF
  ENDCASE 
 ENDIF 

  IF m.vozr=65
   dimtb4(10,1) = dimtb4(10,1) + 1
   dimtb4(10,2) = dimtb4(10,2) + s_all
   IF EMPTY(er)
    dimtb4(10,3) = dimtb4(10,3) + 1
    dimtb4(10,4) = dimtb4(10,4) + s_all
   ENDIF 
   DO CASE 
    CASE m.rslt = 343 
     dimtb4(10,5) = dimtb4(10,5) + 1
     dimtb4(10,6) = dimtb4(10,6) + s_all
     IF EMPTY(er)
      dimtb4(10,7) = dimtb4(10,7) + 1
      dimtb4(10,8) = dimtb4(10,8) + s_all
     ENDIF
    CASE m.rslt = 344
     dimtb4(10,9) = dimtb4(10,9) + 1
     dimtb4(10,10) = dimtb4(10,10) + s_all
     IF EMPTY(er)
      dimtb4(10,11) = dimtb4(10,11) + 1
      dimtb4(10,12) = dimtb4(10,12) + s_all
     ENDIF
    CASE m.rslt = 345
     dimtb4(10,13) = dimtb4(10,13) + 1
     dimtb4(10,14) = dimtb4(10,14) + s_all
     IF EMPTY(er)
      dimtb4(10,15) = dimtb4(10,15) + 1
      dimtb4(10,16) = dimtb4(10,16) + s_all
     ENDIF
   ENDCASE 
  ENDIF 

  IF m.vozr>65
   dimtb4(11,1) = dimtb4(11,1) + 1
   dimtb4(11,2) = dimtb4(11,2) + s_all
   IF EMPTY(er)
    dimtb4(11,3) = dimtb4(11,3) + 1
    dimtb4(11,4) = dimtb4(11,4) + s_all
   ENDIF 
   DO CASE 
    CASE m.rslt = 343 
     dimtb4(11,5) = dimtb4(11,5) + 1
     dimtb4(11,6) = dimtb4(11,6) + s_all
     IF EMPTY(er)
      dimtb4(11,7) = dimtb4(11,7) + 1
      dimtb4(11,8) = dimtb4(11,8) + s_all
     ENDIF
    CASE m.rslt = 344
     dimtb4(11,9) = dimtb4(11,9) + 1
     dimtb4(11,10) = dimtb4(11,10) + s_all
     IF EMPTY(er)
      dimtb4(11,11) = dimtb4(11,11) + 1
      dimtb4(11,12) = dimtb4(11,12) + s_all
     ENDIF
    CASE m.rslt = 345
     dimtb4(11,13) = dimtb4(11,13) + 1
     dimtb4(11,14) = dimtb4(11,14) + s_all
     IF EMPTY(er)
      dimtb4(11,15) = dimtb4(11,15) + 1
      dimtb4(11,16) = dimtb4(11,16) + s_all
     ENDIF
   ENDCASE 
  ENDIF 

 ENDSCAN 


 USE IN curdsp
 USE IN dsp 
 USE IN dspcodes
 
 dimdata(1,2) = TRANSFORM(dimtb4(2,3)+dimtb4(3,3), '99999')
 dimdata(1,3) = TRANSFORM(dimtb4(2,4)+dimtb4(3,4), '999999999.99')

 dimdata(2,2) = TRANSFORM(dimtb4(10,3), '99999')
 dimdata(2,3) = TRANSFORM(dimtb4(10,4), '999999999.99')

 dimdata(3,2) = TRANSFORM(dimtb4(11,3), '99999')
 dimdata(3,3) = TRANSFORM(dimtb4(11,4), '999999999.99')

 m.lcTmpName = pTempl+'\Prof_t2.xls'
 m.lcRepName = pBase+'\'+m.gcPeriod+'\Prof_t2.xls'
 m.IsVisible = .T.
 
 CREATE CURSOR curdata (recid i)
 m.llResult = X_Report(m.lcTmpName, m.lcRepName, m.IsVisible)
 USE IN curdata
 
RETURN 

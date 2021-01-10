# DEFINE CURMONTH .T.
# DEFINE ALLPERIOD .F.

PROCEDURE DspMonitorProf(para1, para2)

 m.NeedOpen = .t.
 m.IsSilent = .f.
 IF PARAMETERS()>0
  m.NeedOpen = para1
 ENDIF 
 IF PARAMETERS()>1
  m.IsSilent = para2
 ENDIF 

 IF !m.IsSilent
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÕÎÒÈÒÅ ÑÔÎÐÌÈÐÎÂÀÒÜ ÎÒ×ÅÒ'+CHR(13)+CHR(10)+;
 'ÏÎ ÏÐÎÔÎÑÌÎÒÐÀÌ ÂÇÐÎÑËÎÃÎ ÍÀÑÅËÅÍÈß?'+CHR(13)+CHR(10),4+32,'ÏðîôÎ_Â_3_032018_50046')=7
  RETURN 
 ENDIF 
 ENDIF 
 
 m.regim = ALLPERIOD
 IF !m.IsSilent
 IF MESSAGEBOX('ÍÀÐÀÑÒÀÞÙÈÉ ÈÒÎÃ (ÄÀ) ÈËÈ ÇÀ ÌÅÑßÖ (ÍÅÒ)?',4+32,'')=6
  m.regim = ALLPERIOD
 ELSE 
  m.regim = CURMONTH
 ENDIF 
 ENDIF 
 
 IF !m.IsSilent
  MESSAGEBOX('ÂÛ ÂÛÁÐÀËÈ '+IIF(m.regim = ALLPERIOD,'ÍÀÐÀÑÒÀÞÙÈÉ ÈÒÎÃ','ÇÀ ÌÅÑßÖ'),0+64,'')
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+gcperiod)
  IF !m.IsSilent
   MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß ÏÅÐÈÎÄÀ!'+CHR(13)+CHR(10),0+16,gcperiod)
  ENDIF 
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\dsp.dbf')
  IF !m.IsSilent
   MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË DSP.DBF!'+CHR(13)+CHR(10),0+16,gcperiod)
  ENDIF 
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pcommon+'\dspcodes.dbf')
  IF !m.IsSilent
  MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË DSPCODES.DBF!'+CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN 
 ENDIF 

 IF OpenFile(pbase+'\'+gcperiod+'\dsp', 'dsp', 'shar')>0
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

 m.period = NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))

 m.mmyy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),2)
 DotName = 'ÏðîôÎ_Â_3_032018_50046.xls'
 DocName = 'ÏðîôÎ_Â_3_032018_50046'

 IF !fso.FileExists(ptempl+'\'+dotname)
  IF !m.IsSilent
   MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ØÀÁËÎÍ ÎÒ×ÅÒÀ ' + ptempl+'\'+dotname + CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN 
 ENDIF 
 
 DIMENSION dimtb4(40,40)
 dimtb4 = 0

 =MakePage2old()

 IF USED('dsp')
  USE IN dsp
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 

 IF fso.FileExists(pbase+'\'+gcperiod+'\'+DocName+'.xls')
  fso.DeleteFile(pbase+'\'+gcperiod+'\'+DocName+'.xls')
 ENDIF 
 
 CREATE CURSOR curdata (n_rec i)
 
 dimtb4(1,1) = TRANSFORM(dimtb4(2,1) + dimtb4(3,1), '99999999')
 dimtb4(1,2) = TRANSFORM(dimtb4(2,2) + dimtb4(3,2), '999999999.99')
 dimtb4(1,3) = TRANSFORM(dimtb4(2,3) + dimtb4(3,3), '99999999')
 dimtb4(1,4) = TRANSFORM(dimtb4(2,4) + dimtb4(3,4), '999999999.99')

 dimtb4(2,1) = TRANSFORM(dimtb4(2,1), '9999999')
 dimtb4(2,2) = TRANSFORM(dimtb4(2,2), '99999999.99')
 dimtb4(2,3) = TRANSFORM(dimtb4(2,3), '9999999')
 dimtb4(2,4) = TRANSFORM(dimtb4(2,4), '99999999.99')

 dimtb4(3,1) = TRANSFORM(dimtb4(3,1), '9999999')
 dimtb4(3,2) = TRANSFORM(dimtb4(3,2), '99999999.99')
 dimtb4(3,3) = TRANSFORM(dimtb4(3,3), '9999999')
 dimtb4(3,4) = TRANSFORM(dimtb4(3,4), '99999999.99')
 
 m.llResult = X_Report(ptempl+'\'+m.dotname, pbase+'\'+gcperiod+'\'+DocName+'.xls', .F.)
 
 USE IN curdata
 
 IF !m.IsSilent
  MESSAGEBOX('ÎÒ×¨Ò ÑÔÎÐÌÈÐÎÂÀÍ. ÔÀÉË ÑÎÕÐÀÍ¨Í ÏÎ ÀÄÐÅÑÓ:'+CHR(13)+CHR(10)+UPPER(pbase+'\'+gcperiod+'\'+DocName+'.xls'),0+64,'')
 ENDIF 
 
RETURN 


FUNCTION MakePage2Old
 DIMENSION dimtb4(9,20)
 dimtb4 = 0
 
 SELECT dsp
 
 SCAN 
  m.d_u  = d_u
  m.cod = cod 
  m.rslt = rslt
  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF m.tipofcod!=2
   LOOP 
  ENDIF 

  *m.vozr = ROUND((m.tdat1 - dr)/365.25,2)
  m.vozr = ROUND((m.d_u - dr)/365.25,2)

  IF m.vozr>=18 AND w=1
   dimtb4(2,1) = dimtb4(2,1) + 1
   dimtb4(2,2) = dimtb4(2,2) + s_all
   IF EMPTY(er)
    dimtb4(2,3) = dimtb4(2,3) + 1
    dimtb4(2,4) = dimtb4(2,4) + s_all
   ENDIF
  ENDIF 
  IF m.vozr=65 AND w=1
   dimtb4(4,1) = dimtb4(4,1) + 1
   dimtb4(4,2) = dimtb4(4,2) + s_all
   IF EMPTY(er)
    dimtb4(4,3) = dimtb4(4,3) + 1
    dimtb4(4,4) = dimtb4(4,4) + s_all
   ENDIF
  ENDIF 
  IF m.vozr>65 AND w=1
   dimtb4(5,1) = dimtb4(5,1) + 1
   dimtb4(5,2) = dimtb4(5,2) + s_all
   IF EMPTY(er)
    dimtb4(5,3) = dimtb4(5,3) + 1
    dimtb4(5,4) = dimtb4(5,4) + s_all
   ENDIF
  ENDIF 

  IF m.vozr>=18 and w=2
   dimtb4(3,1) = dimtb4(3,1) + 1
   dimtb4(3,2) = dimtb4(3,2) + s_all
   IF EMPTY(er)
    dimtb4(3,3) = dimtb4(3,3) + 1
    dimtb4(3,4) = dimtb4(3,4) + s_all
   ENDIF 
  ENDIF 
  IF m.vozr=65 and w=2
   dimtb4(6,1) = dimtb4(6,1) + 1
   dimtb4(6,2) = dimtb4(6,2) + s_all
   IF EMPTY(er)
    dimtb4(6,3) = dimtb4(6,3) + 1
    dimtb4(6,4) = dimtb4(6,4) + s_all
   ENDIF 
  ENDIF 
  IF m.vozr>65 and w=2
   dimtb4(7,1) = dimtb4(7,1) + 1
   dimtb4(7,2) = dimtb4(7,2) + s_all
   IF EMPTY(er)
    dimtb4(7,3) = dimtb4(7,3) + 1
    dimtb4(7,4) = dimtb4(7,4) + s_all
   ENDIF 
  ENDIF 

 ENDSCAN 

 
RETURN 

FUNCTION IsWDR(w, pol, age, vozr1, vozr2)
 PRIVATE w, pol, age, dr1, dr2
 IF m.w!=m.pol
  RETURN .F.
 ENDIF 
 IF !BETWEEN(m.age, m.vozr1, m.vozr2)
  RETURN .F.
 ENDIF 
RETURN .T.


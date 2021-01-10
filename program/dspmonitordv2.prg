# DEFINE CURMONTH .T.
# DEFINE ALLPERIOD .F.

PROCEDURE DspMonitorDV2(para1, para2)
 m.NeedOpen = .t.
 m.IsSilent = .f.
 IF PARAMETERS()>0
  m.NeedOpen = para1
 ENDIF 
 IF PARAMETERS()>1
  m.IsSilent = para2
 ENDIF 

 IF !m.IsSilent
  IF MESSAGEBOX(CHR(13)+CHR(10)+'’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“'+CHR(13)+CHR(10)+;
   'œŒ ƒ»—œ¿Õ—≈–»«¿÷»» ¬«–Œ—ÀŒ√Œ Õ¿—≈À≈Õ»ﬂ?'+CHR(13)+CHR(10),4+32,'ƒ¬_1_032018_50046, ¬≈–—»ﬂ 2')=7
  RETURN 
  ENDIF 
 ENDIF 
 
 IF !m.IsSilent
  IF MESSAGEBOX('Õ¿–¿—“¿ﬁŸ»… »“Œ√ (ƒ¿) »À» «¿ Ã≈—ﬂ÷ (Õ≈“)?',4+32,'')=6
   m.regim = ALLPERIOD
  ELSE 
   m.regim = CURMONTH
  ENDIF 
 ELSE 
  m.regim = ALLPERIOD
 ENDIF 
 
 IF !m.IsSilent
  MESSAGEBOX('¬€ ¬€¡–¿À» '+IIF(m.regim = ALLPERIOD,'Õ¿–¿—“¿ﬁŸ»… »“Œ√','«¿ Ã≈—ﬂ÷'),0+64,'')
 ENDIF 
 
 IF m.regim = ALLPERIOD
  m.pgdat1 = {01.04.2019}
 ELSE 
  m.pgdat1 = m.tdat1
 ENDIF 
 m.pgdat2 = m.tdat2
 
 IF m.regim = ALLPERIOD
  DO FORM selperiod
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+gcperiod)
  IF !m.IsSilent
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿!'+CHR(13)+CHR(10),0+16,gcperiod)
  ENDIF 
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\dsp.dbf')
  IF !m.IsSilent
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À DSP.DBF!'+CHR(13)+CHR(10),0+16,gcperiod)
  ENDIF 
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(pcommon+'\dspcodes.dbf')
  IF !m.IsSilent
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À DSPCODES.DBF!'+CHR(13)+CHR(10),0+16,'')
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
 DotName = 'ƒ¬_1_032018_50046_1.xls'
 DocName = 'ƒ¬_1_032018_50046_1'
 IF !fso.FileExists(ptempl+'\'+dotname)
  IF !m.IsSilent
   MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ Œ“◊≈“¿ ' + ptempl+'\'+dotname + CHR(13)+CHR(10),0+16,'')
  ENDIF 
  RETURN 
 ENDIF 
 
 DIMENSION dimdata(40,40)
 dimdata = 0

 =MakePage3()

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
 
 FOR n=16 TO 38
  dimdata(n,9) = ROUND(VAL(dimdata(n,9))/1000,2)
  dimdata(n,12) = ROUND(VAL(dimdata(n,12))/1000,2)
  dimdata(n,15) = ROUND(VAL(dimdata(n,15))/1000,2)
  dimdata(n,18) = ROUND(VAL(dimdata(n,18))/1000,2)
  dimdata(n,23) = ROUND(VAL(dimdata(n,23))/1000,2)
  dimdata(n,28) = ROUND(VAL(dimdata(n,28))/1000,2)
  
  dimdata(n,33) = VAL(dimdata(n,20)) - (VAL(dimdata(n,34)) + VAL(dimdata(n,35)) + VAL(dimdata(n,36)))

 ENDFOR 
 
 
 m.llResult = X_Report(ptempl+'\'+m.dotname, pbase+'\'+gcperiod+'\'+DocName+'.xls', .F.)
 
 USE IN curdata
 
 IF !m.IsSilent
  MESSAGEBOX('Œ“◊®“ —‘Œ–Ã»–Œ¬¿Õ. ‘¿…À —Œ’–¿Õ®Õ œŒ ¿ƒ–≈—”:'+CHR(13)+CHR(10)+UPPER(pbase+'\'+gcperiod+'\'+DocName+'.xls'),0+64,'')
 ENDIF 
 
RETURN 


FUNCTION MakePage3

 CREATE CURSOR curdsp (sn_pol c(25))
 INDEX ON sn_pol TAG sn_pol
 SET ORDER TO sn_pol

 CREATE CURSOR curlpu1 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu2 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu3 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu4 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu5 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu6 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu7 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu8 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu9 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu10 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu11 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu12 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 
 CREATE CURSOR curlpu13 (mcod c(7))
 INDEX ON mcod TAG mcod 
 SET ORDER TO mcod 

 SELECT dsp

 SCAN 
  m.cod    = cod
  m.w      = w
  m.ages   = ages
  m.sn_pol = sn_pol
  m.mcod   = mcod
  m.d_u    = d_u

  IF m.regim = CURMONTH
   IF !BETWEEN(m.d_u, m.tdat1, m.tdat2)
    LOOP
   ENDIF 
  ELSE 
   *IF !BETWEEN(m.d_u, IIF(m.qcod='S7', {01.04.2019}, {01.05.2019}), m.tdat2)
   IF !BETWEEN(m.d_u, m.pgdat1, m.pgdat2)
    LOOP
   ENDIF 
  ENDIF 

  IF !SEEK(m.cod, 'dspcodes')
   LOOP
  ENDIF 
  m.tipofcod = dspcodes.tip
  IF m.tipofcod!=1
   LOOP 
  ENDIF 
  IF m.ages<18
   LOOP 
  ENDIF 
  IF INLIST(m.cod, 25204, 35401)
   LOOP 
  ENDIF 

  =incdimdsp(16)
  IF !SEEK(m.mcod, 'curlpu1')
   INSERT INTO curlpu1 (mcod) VALUES (m.mcod)
  ENDIF 
  
  IF m.w=1  && ‚ÚÓ‡ˇ ÒÚÓÍ‡, ÏÛÊ˜ËÌ˚
   IF !SEEK(m.mcod, 'curlpu2')
    INSERT INTO curlpu2 (mcod) VALUES (m.mcod)
   ENDIF 
   =incdimdsp(17)
  ENDIF 
  IF m.w=1 AND BETWEEN(m.ages, 18, 39)
   IF !SEEK(m.mcod, 'curlpu3')
    INSERT INTO curlpu3 (mcod) VALUES (m.mcod)
   ENDIF 
   =incdimdsp(18)
  ENDIF 
  IF m.w=1 AND BETWEEN(m.ages, 40, 64)
   IF !SEEK(m.mcod, 'curlpu4')
    INSERT INTO curlpu4 (mcod) VALUES (m.mcod)
   ENDIF 
   =incdimdsp(24)
   =incdimdsp(33)
  ENDIF 
  IF m.w=1 AND m.ages=65
   =incdimdsp(24)
   =incdimdsp(35)
  ENDIF 
  IF m.w=1 AND m.ages>65
   =incdimdsp(24)
   =incdimdsp(36)
  ENDIF 

  IF m.w=2 && ‚ÚÓ‡ˇ ÒÚÓÍ‡, ÊÂÌ˘ËÌ˚
   IF !SEEK(m.mcod, 'curlpu8')
    INSERT INTO curlpu8 (mcod) VALUES (m.mcod)
   ENDIF 
   =incdimdsp(25)
  ENDIF 
  IF m.w=2 AND BETWEEN(m.ages, 18, 39)
   IF !SEEK(m.mcod, 'curlpu9')
    INSERT INTO curlpu9 (mcod) VALUES (m.mcod)
   ENDIF 
   =incdimdsp(26)
  ENDIF 
  IF m.w=2 AND BETWEEN(m.ages, 40, 64)
   IF !SEEK(m.mcod, 'curlpu10')
    INSERT INTO curlpu10 (mcod) VALUES (m.mcod)
   ENDIF 
   =incdimdsp(32)
   =incdimdsp(34)
  ENDIF 
  IF m.w=2 AND m.ages=65
   =incdimdsp(32)
   =incdimdsp(37)
  ENDIF 
  IF m.w=2 AND m.ages>65
   =incdimdsp(32)
   =incdimdsp(38)
  ENDIF 

  IF !SEEK(m.sn_pol, 'curdsp')
  	INSERT INTO curdsp (sn_pol) VALUES (m.sn_pol)
  ENDIF 

 ENDSCAN 

 dimdata(16,4) = RECCOUNT('curlpu1')
 dimdata(17,4) = RECCOUNT('curlpu2')
 dimdata(19,4) = RECCOUNT('curlpu3')
 dimdata(20,4) = RECCOUNT('curlpu4')
 dimdata(21,4) = RECCOUNT('curlpu5')
 dimdata(22,4) = RECCOUNT('curlpu6')
 dimdata(23,4) = RECCOUNT('curlpu7')
 dimdata(25,4) = RECCOUNT('curlpu8')
 dimdata(27,4) = RECCOUNT('curlpu9')
 dimdata(28,4) = RECCOUNT('curlpu10')
 dimdata(29,4) = RECCOUNT('curlpu11')
 dimdata(30,4) = RECCOUNT('curlpu12')
 dimdata(31,4) = RECCOUNT('curlpu13')

 FOR m.nstr=16 TO 38
   dimdata(m.nstr,09) = TRANSFORM(dimdata(m.nstr,12)+dimdata(m.nstr,15), '999999999.99') && 1-˚È + 2-ÓÈ ˝Ú‡Ô
   dimdata(m.nstr,10) = TRANSFORM(dimdata(m.nstr,10), '99999')
   dimdata(m.nstr,12) = TRANSFORM(dimdata(m.nstr,12), '999999999.99') && 1-˚È
   dimdata(m.nstr,13) = TRANSFORM(dimdata(m.nstr,13), '99999') 
   dimdata(m.nstr,14) = TRANSFORM(dimdata(m.nstr,13), '99999')  && »ÁÏÂÌÂÌÓ 18.04.2019 ‚ÏÂÒÚÂ Ò 3-ÂÈ ÒÚ‡ÌËˆÂÈ dispmon
   dimdata(m.nstr,15) = TRANSFORM(dimdata(m.nstr,15), '999999999.99') && 2-ÓÈ ˝Ú‡Ô
   
   dimdata(m.nstr,18) = TRANSFORM(dimdata(m.nstr,23)+dimdata(m.nstr,28), '999999999.99')
   dimdata(m.nstr,20) = TRANSFORM(dimdata(m.nstr,22)-dimdata(m.nstr,24), '99999')
   dimdata(m.nstr,23) = TRANSFORM(dimdata(m.nstr,23), '999999999.99')
   dimdata(m.nstr,24) = TRANSFORM(dimdata(m.nstr,24), '99999')
   dimdata(m.nstr,25) = TRANSFORM(dimdata(m.nstr,24), '99999')  && »ÁÏÂÌÂÌÓ 18.04.2019 ‚ÏÂÒÚÂ Ò 3-ÂÈ ÒÚ‡ÌËˆÂÈ dispmon
   dimdata(m.nstr,28) = TRANSFORM(dimdata(m.nstr,28), '999999999.99')

   dimdata(m.nstr,31) = TRANSFORM(dimdata(m.nstr,31), '99999') &&
   dimdata(m.nstr,33) = TRANSFORM(dimdata(m.nstr,33), '99999') &&
   dimdata(m.nstr,34) = TRANSFORM(dimdata(m.nstr,34), '99999') &&

   dimdata(m.nstr,38) = TRANSFORM(dimdata(m.nstr,35) + dimdata(m.nstr,36), '99999') &&

   dimdata(m.nstr,35) = TRANSFORM(dimdata(m.nstr,35), '99999') &&
   dimdata(m.nstr,36) = TRANSFORM(dimdata(m.nstr,36), '99999') &&

 ENDFOR 
 
 USE IN curdsp
 USE IN curlpu1
 USE IN curlpu2
 USE IN curlpu3
 USE IN curlpu4
 USE IN curlpu5
 USE IN curlpu6
 USE IN curlpu7
 USE IN curlpu8
 USE IN curlpu9
 USE IN curlpu10
 USE IN curlpu11
 USE IN curlpu12
 USE IN curlpu13

RETURN 

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

FUNCTION incdimdsp(nstr)
 PRIVATE nstr
  dimdata(m.nstr,10) = dimdata(m.nstr,10) + 1
  dimdata(m.nstr,12) = dimdata(m.nstr,12) + s_all

  IF !EMPTY(k_u2)
   dimdata(m.nstr,13) = dimdata(m.nstr,13) + 1
   dimdata(m.nstr,14) = dimdata(m.nstr,14) + k_u2
   dimdata(m.nstr,15) = dimdata(m.nstr,15) + s_all2
  ENDIF 

  IF EMPTY(er)
   dimdata(m.nstr,22) = dimdata(m.nstr,22) + 1
   dimdata(m.nstr,23) = dimdata(m.nstr,23) + s_all
   IF !EMPTY(k_u2)
    dimdata(m.nstr,24) = dimdata(m.nstr,24) + 1
    dimdata(m.nstr,25) = dimdata(m.nstr,25) + k_u2ok
    dimdata(m.nstr,28) = dimdata(m.nstr,28) + s_all2ok
   ENDIF 

   
   DO CASE 
    CASE INLIST(rslt,352,353,357,358) && INLIST(rslt,316,352,353,354)
     dimdata(m.nstr,31) = dimdata(m.nstr,31) + 1
    CASE INLIST(rslt,317,352)
     dimdata(m.nstr,33) = dimdata(m.nstr,33) + 1
    CASE INLIST(rslt,318,353)
     dimdata(m.nstr,34) = dimdata(m.nstr,34) + 1
    CASE INLIST(rslt,355,357)
     dimdata(m.nstr,35) = dimdata(m.nstr,35) + 1
     dimdata(m.nstr,37) = dimdata(m.nstr,37) + IIF(ds='C', 1, 0)
    CASE INLIST(rslt,356,358)
     dimdata(m.nstr,36) = dimdata(m.nstr,36) + 1
     dimdata(m.nstr,37) = dimdata(m.nstr,37) + IIF(ds='C', 1, 0)
   ENDCASE 
  ENDIF 

RETURN 

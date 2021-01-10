PROCEDURE DspMonitorXX
 IF MESSAGEBOX(CHR(13)+CHR(10)+'’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“'+CHR(13)+CHR(10)+;
 'œŒ ƒ»—œ¿Õ—≈–»«¿÷»» œ–»ÀŒ∆≈Õ»≈ 4?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+gcperiod)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿!'+CHR(13)+CHR(10),0+16,gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À PEOPLE.DBF!'+CHR(13)+CHR(10),0+16,gcperiod)
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\talon.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À TALON.DBF!'+CHR(13)+CHR(10),0+16,gcperiod)
  RETURN 
 ENDIF 
 

 IF OpenFile(pbase+'\'+gcperiod+'\people', 'people', 'shar', 'sn_pol')>0
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\talon', 'talon', 'shar')>0
  USE IN people
  IF USED('talon')
   USE IN talon
  ENDIF 
  RETURN 
 ENDIF 
 
 CREATE CURSOR dsp (recid i, period c(6), mcod c(7), sn_pol c(17), c_i c(30), ;
 	fam c(20), im c(20), ot c(20), w n(1), dr d, ages n(2), cod n(6), rslt n(3), ;
 	d_u d, s_all n(11,2), k_u2 n(3), s_all2 n(11,2), k_u2ok n(3), s_all2ok n(11,2), er c(3))
 
 SELECT talon 
 SET RELATION TO sn_pol INTO people
 SCAN 
  *m.dn = dn
  *IF !INLIST(m.dn,1,2) && EMPTY(m.dn)
  * LOOP 
  *ENDIF 
  m.p_cel = p_cel
  IF m.p_cel != '1.3'
   LOOP 
  ENDIF 
 
  m.d_u   = d_u
  m.k_u   = k_u
  m.s_all = s_all
  m.er    = ''

  m.id_smo = recid
  m.sn_pol = sn_pol
  m.fam    = people.fam
  m.im     = people.im
  m.ot     = people.ot
  m.w      = people.w
  m.dr     = people.dr
   
  m.ages   = YEAR(tdat1) - YEAR(m.dr)
  
  m.c_i    = c_i 
  
  INSERT INTO dsp FROM MEMVAR 
 ENDSCAN 
 SET RELATION OFF INTO people
 USE 
 USE IN people 

 m.period = NameOfMonth(VAL(SUBSTR(m.gcperiod,5,2)))

 m.mmyy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),2)
 DotName = 'œËÎÓÊÂÌËÂ4.xls'
 DocName = 'œËÎÓÊÂÌËÂ4'+m.qcod+m.mmyy
 IF !fso.FileExists(ptempl+'\'+dotname)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ÿ¿¡ÀŒÕ Œ“◊≈“¿ œËÎÓÊÂÌËÂ4.xls'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 SELECT dsp
 
 DIMENSION dimdata(9,17)
 dimdata = 0

 SCAN 
  m.cod    = cod
  m.w      = w
  m.ages   = ages
  m.sn_pol = sn_pol
  m.mcod   = mcod

  IF m.ages<18
   LOOP 
  ENDIF 
  IF m.w=1 AND m.cod = 25204 AND !INLIST(m.ages,49,53,55,59,61,65,67,71,73)
   LOOP 
  ENDIF 
  IF m.w=2 AND !INLIST(m.ages,49,53,55,59,61,65,67,71,73) AND !INLIST(m.ages,50,52,56,58,62,64,68,70) AND INLIST(m.cod, 25204, 35401)
   LOOP 
  ENDIF 
  
  =incdimdata(1)

  IF m.w=1 && ‚ÚÓ‡ˇ ÒÚÓÍ‡, ÏÛÊ˜ËÌ˚
   =incdimdata(2)
  ENDIF 
  IF m.w=1 AND m.ages=65
   =incdimdata(3)
  ENDIF 
  IF m.w=1 AND m.ages>65
   =incdimdata(4)
  ENDIF 

  IF m.w=2 && ‚ÚÓ‡ˇ ÒÚÓÍ‡, ÊÂÌ˘ËÌ˚
   =incdimdata(5)
  ENDIF 
  IF m.w=2 AND m.ages=65
   =incdimdata(6)
  ENDIF 
  IF m.w=2 AND m.ages>65
   =incdimdata(7)
  ENDIF 

 ENDSCAN 
 
 IF USED('dsp')
  USE IN dsp
 ENDIF 
 IF USED('dspcodes')
  USE IN dspcodes
 ENDIF 
 
 CREATE CURSOR curdata (n_rec i)

 m.llResult = X_Report(ptempl+'\'+m.dotname, pbase+'\'+gcperiod+'\'+DocName+'.xls', .F.)
 
 USE IN curdata
 
 MESSAGEBOX('Œ“◊®“ —‘Œ–Ã»–Œ¬¿Õ. ‘¿…À —Œ’–¿Õ®Õ œŒ ¿ƒ–≈—”:'+CHR(13)+CHR(10)+UPPER(pbase+'\'+gcperiod+'\'+DocName+'.xls'),0+64,'')
 

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

FUNCTION incdimdata(nstr)
 PRIVATE nstr
  *dimdata(m.nstr,3) = dimdata(m.nstr,3) + 1
  dimdata(m.nstr,3) = dimdata(m.nstr,3) + s_all
  dimdata(m.nstr,4) = dimdata(m.nstr,4) + 1
  dimdata(m.nstr,5) = dimdata(m.nstr,5) + s_all

  *IF EMPTY(er)
  * dimdata(m.nstr,6) = dimdata(m.nstr,6) + 1
  * dimdata(m.nstr,7) = dimdata(m.nstr,7) + s_all
  *ENDIF 

RETURN 


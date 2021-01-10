PROCEDURE MakeYFiles

 IF MESSAGEBOX('ÑÔÎÐÌÈÐÎÂÀÜ ÏÅÐÑÎÒ×ÅÒ?',4+32,'')=7
  RETURN 
 ENDIF 

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar', 'mcod')>0
  RETURN
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\TarifN', 'tarif', 'shar', 'cod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', 'sprlpu', 'shar', 'mcod')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sookodxx', "sookod", "SHARED", "er_c")>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilot', 'pilot', 'shar', 'mcod')>0
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\pilots', 'pilots', 'shar', 'mcod')>0
  IF USED('pilots')
   USE IN pilots
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+m.gcperiod+'\nsi\lputpn', 'lputpn', 'shar', 'lpu_id')>0
  IF USED('pilots')
   USE IN pilots
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  RETURN 
 ENDIF 

DO CASE 
 CASE m.ffoms = 77011
  m.LoadPeriod={01.02.2011}+6*365
 CASE m.ffoms = 77002
  m.LoadPeriod={13.01.2011}+6*365
 CASE m.ffoms = 77008
  m.LoadPeriod={03.01.2011}+6*365
 CASE m.ffoms = 77013
  m.LoadPeriod={26.01.2011}+6*365
 CASE m.ffoms = 77012
  m.LoadPeriod={30.01.2011}+6*365
 OTHERWISE 
  m.LoadPeriod={30.01.2011}+6*365
 
ENDCASE 

IF DATE()>m.LoadPeriod
* =ChkDirsBrief()
ENDIF 

 mmy = PADL(tMonth,2,'0') + RIGHT(STR(tYear,4),1)
 
 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 SELECT AisOms

 SCAN
  m.mcod = mcod 
  IF mcod = '0371001'
   LOOP 
  ENDIF 
  IF s_pred <= 0
   LOOP 
  ENDIF 
  
  m.nIsDoubles = 0
  
  m.sum_st1 = 0  && Ñóììà ê îïëàòå, ïîëó÷åííàÿ âû÷èòàíèåì ÔËÊ èç ïðåäñòàâëåííûõ
  m.sum_st1 = s_pred - sum_flk

  m.sum_st2 = 0  && Ñóììà ê îïëàòå ïî äàííûì àãðåãàòêè
    
  m.mcod = mcod
  m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)
  m.IsPilot = IIF(SEEK(m.mcod, 'pilot'), .t., .f.)
  m.IsPilotS= IIF(SEEK(m.mcod, 'pilots'), .t., .f.)
  
  m.s_calc_pf = finval + finvals

  WAIT m.mcod WINDOW NOWAIT 

  m.lpu_id = lpuid
  m.IsLpuTpn = IIF(SEEK(m.lpu_id, 'lputpn'), .t., .f.)

  lcPath = pBase+'\'+m.gcperiod+'\'+m.mcod
  IF !fso.FolderExists(lcPath)
   LOOP 
  ENDIF 
  
  MmyName = 'D'+m.qcod+STR(m.lpu_id,4)+'.'+mmy
  IF fso.FileExists(lcpath+'\'+MmyName)
   *LOOP 
  ENDIF 
  
  =MakeYFilesOne(lcPath)

  SET DEFAULT TO (pBin)

  SELECT AisOms

  IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÏÐÅÐÂÀÒÜ ÎÁÐÀÁÎÒÊÓ?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 
 
 ENDSCAN 
 WAIT CLEAR 
 USE 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('tarif')
   USE IN tarif
  ENDIF 
  IF USED('sookod')
   USE IN sookod
  ENDIF 
  IF USED('pilot')
   USE IN pilot
  ENDIF 
  IF USED('pilots')
   USE IN pilots
  ENDIF 
  IF USED('lputpn')
   USE IN lputpn
  ENDIF 

 SET ESCAPE &OldEscStatus
 
 MESSAGEBOX("ÎÁÐÀÁÎÒÊÀ ÇÀÊÎÍ×ÅÍÀ!",0+64,"")
 
RETURN 



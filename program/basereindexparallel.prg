PROCEDURE BaseReindexParallel
 IF MESSAGEBOX('ÏÅÐÅÈÍÄÅÊÑÈÐÎÂÀÒÜ ÐÀÁÎ×ÈÅ ÁÀÇÛ?',4+32,'')=7
  RETURN 
 ENDIF  
 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 Local Parallel as Parallel of ParallelFox.vcx
 Parallel = NewObject("Parallel", "ParallelFox.vcx")

 Parallel.StartWorkers(FullPath("lpu2smo.exe"),,.t.)

 Local i, lnTimer
 lnTimer = Seconds()
 SELECT aisoms

 SCAN 
  m.mcod = mcod 
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  *=PeopleReindex(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people')
  Parallel.Do("PeopleReindex","BaseReindexParallel.prg",, m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people')
  WAIT CLEAR 
  
  SELECT aisoms 
  
 ENDSCAN 
 
 Parallel.Wait()
 
 MESSAGEBOX("Total Time: "+TRANSFORM(Seconds()-lnTimer, '999999.99')+' ñåê.',0+64,'')
 
RETURN 

FUNCTION PeopleReindex(para1)
 SET SAFETY OFF 
 SET PROCEDURE TO Utils.prg
 
 LOCAL tcFile
 tcFile = para1
 
 *USE &tcFile IN 0 ALIAS people EXCLUSIVE 
 IF OpenFile(tcFile, 'people', 'excl')>0
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 0
 ENDIF 
 
 SELECT People 
 DELETE TAG ALL 
 SET FULLPATH OFF 
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON sn_pol TAG sn_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX on dr TAG dr
 INDEX ON s_all TAG s_all 
 SET FULLPATH OFF 
 USE IN people 
RETURN 1

PROCEDURE myParallel

* Example calling functions
Local i, lnTimer

CreateAppObject()

lnTimer = Seconds()

For i = 1 to 10
	? "Running Units of Work", i * 10
	_Screen.oApp.Test(i * 10)
EndFor 
	
MESSAGEBOX("Total Time: "+TRANSFORM(Seconds()-lnTimer, '999999.99')+' ñåê.',0+64,'')

IF 1=2
* Example calling functions

Local i, lnTimer

Local Parallel as Parallel of ParallelFox.vcx
Parallel = NewObject("Parallel", "ParallelFox.vcx")

Parallel.StartWorkers(FullPath("lpu2smo.exe"),,.f.)

Parallel.Call("CreateAppObject",.t.)

lnTimer = Seconds()

For i = 1 to 10
	? "Running Units of Work", i * 10
	Parallel.Call("_Screen.oApp.Test",,i * 10)
EndFor 
	
Parallel.Wait()
*? "Total Time", Seconds() - lnTimer
MESSAGEBOX("Total Time: "+TRANSFORM(Seconds()-lnTimer, '999999.99')+' ñåê.',0+64,'')
ENDIF 

Return 





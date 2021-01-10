PROCEDURE BaseReindInParallel

 IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\aisoms.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.pBase+'\'+m.gcPeriod+'\aisoms', 'aisoms', 'shar')>0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 

 Local Parallel as Parallel
 Parallel = NewObject("Parallel", "ParallelFox.vcx")
 
 Parallel.StartWorkers(FULLPATH("lpu2smo.prg"),,.f.)
 
 m.t_0 = SECONDS()
 
 SELECT aisoms
 SCAN 

  m.mcod = mcod 
  IF !fso.FolderExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\people.dbf')
   LOOP  
  ENDIF 

  TEXT TO cScript NOSHOW 
   LPARAMETERS pBase, gcPeriod, mcod
   
   SET SAFETY OFF
   *SET PROCEDURE TO Utils
   
   USE &pBase\&gcPeriod\&mcod\people IN 0 ALIAS people EXCLUSIVE 

   SELECT People 
   DELETE TAG ALL 
   INDEX ON RecId TAG recid CANDIDATE 
   INDEX ON recid_lpu TAG recid_lpu
   INDEX ON sn_pol TAG sn_pol
   INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
   INDEX on dr TAG dr
   INDEX ON s_all TAG s_all 
   USE IN people 
   ? m.mcod
   
   RETURN 
  ENDTEXT 
  
  Parallel.ExecScript(cScript,, m.Pbase, m.gcPeriod, m.mcod)
  
  SELECT aisoms

 ENDSCAN 
 
 Parallel.Wait()

 USE IN aisoms 
 
 m.t_1 = SECONDS()
 
 MESSAGEBOX(TRANSFORM(m.t_1-m.t_0,'99999.99'),0+64,'')
 
RETURN 
PROCEDURE TryParallel
 IF MESSAGEBOX('гюосярхрэ лмнцнонрнвмнярэ?',4+32,'')=7
  RETURN 
 ENDIF 

 Local Parallel as Parallel
 Parallel = NewObject("Parallel", "ParallelFox.vcx")
 
 Parallel.StartWorkers("lpu2smo.exe",,.t.)
  
 Parallel.Do('SetEnv',,.T.,pCommon)
 Parallel.do('comreind')
 
 Parallel.Wait()
 
 MESSAGEBOX('OK!',0+64,'')
RETURN 
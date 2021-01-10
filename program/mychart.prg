PROCEDURE myChart
 USE d:\lpu2smo\mee\201905\mekstat.dbf IN 0 SHARED
 SELECT c_err, SUM(k_u) as k_u FROM mekstat GROUP BY c_err INTO CURSOR curss
 
 DO FORM chart_mekdef

RETURN 
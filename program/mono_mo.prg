PROCEDURE mono_mo
SELECT * FROM nsio WHERE lpu_id in ;
	(SELECT lpu_id FROM nsio GROUP BY lpu_id HAVING coun(*)=1;
	WHERE kol>0) AND  kol>0 INTO TABLE  d:\lpu2smo\base\202001\mono
RETURN 
PROCEDURE qwerty
  SCAN  
   SCATTER MEMVAR 

   m.IsTpnR    = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='r' AND !(IsKdS(m.cod) OR IsEko(m.cod)), .T., .F.)
   
   m.prmcod = people.prmcod
   m.prik   = IIF(SEEK(m.prmcod, 'sprlpu'), sprlpu.lpu_id, 0)
   
   m.lIs02 = IIF(SEEK(m.cod, 'tarif') AND tarif.tpn='q', .t., .f.)
   m.lpu_ord = IIF(!EMPTY(FIELD('lpu_ord')), lpu_ord, 0)
   m.paztip = TipOfPaz(m.mcod, m.prmcod) && 0 (не прикреплен),1 (прикреплен по месту обращени€),2 (к пилоту),3 (не к пилоту)
   
   DO CASE 
    CASE IsMes(m.cod) OR IsVmp(m.cod) OR IsKDS(m.cod)
     m.f_type = 'ft' && по тарифу

    CASE m.IsTpnR OR IsPat(m.cod) OR IsEKO(m.cod) OR INLIST(SUBSTR(m.otd,2,2),'70','73','93') OR m.d_type='s'
     m.f_type = 'fh' && допуслуги

    CASE (INLIST(SUBSTR(m.otd,2,2),'01','90') AND IsStac(m.mcod)) AND TipOfPr(m.mcod, m.prmcod) = 1
     m.f_type = 'fp' && из средств подушевого финансировани€

    CASE (m.ord=7 AND m.lpu_ord=7665) AND TipOfPr(m.mcod, m.prmcod) = 1
     m.f_type = 'fp' && из средств подушевого финансировани€
     
    OTHERWISE 

    DO CASE 
     CASE TipOfPr(m.mcod, m.prmcod) = 0 && неприкреплен
      m.f_type = 'fp'
      CASE TipOfPr(m.mcod, m.prmcod) = 2 && к пилоту
      IF (m.lIs02 OR INLIST(m.otd,'08','91') OR (m.profil='100' AND INLIST(m.otd,'00','92'))) OR m.lpu_ord>0
       m.f_type = 'vz' && взаимозачеты (из средств иного Ћѕ”)
      ELSE 
       m.f_type = 'fp'
      ENDIF 
      CASE TipOfPr(m.mcod, m.prmcod) = 3 && свой
      m.f_type = 'fp'
      OTHERWISE 
      m.f_type = ''
    ENDCASE 
   ENDCASE 
   
   IF !m.IsPilot
    IF IsKDS(m.cod)
     m.f_type=' '
    ELSE 
     m.f_type='ft'
    ENDIF 
   ENDIF 

   INSERT INTO syfile FROM MEMVAR 
  
  ENDSCAN 

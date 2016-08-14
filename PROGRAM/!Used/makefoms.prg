PROCEDURE MakeFOMS
 IF MESSAGEBOX(CHR(13)+CHR(10)+'—‘Œ–Ã»–Œ¬¿“‹ Œ“◊≈“ ¬ Ã√‘ŒÃ—?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 IF !fso.FileExists(ptempl+'\prqqmmyy.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À ÿ¿¡ÀŒÕ¿ '+'PRQQMMYY.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF !fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'—¬ŒƒÕ€… –≈√»—“– Õ≈ —Œ¡–¿Õ!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\people', 'people', 'excl')>0
  RETURN 
 ENDIF 
 
 prfile = 'pr'+LOWER(m.qcod)+PADL(tMonth,2,'0')+RIGHT(STR(tYear,4),2)
 fso.CopyFile(ptempl+'\prqqmmyy.dbf', ;
  pbase+'\'+gcperiod+'\'+prfile+'.dbf')

 IF OpenFile(pbase+'\'+gcperiod+'\'+prfile, 'prfile', 'shar')>0
  IF USED('prfile')
   USE IN prfile
  ENDIF 
  fso.DeleteFile(pbase+'\'+gcperiod+'\'+prfile+'.dbf')
  RETURN 
 ENDIF 
 
 WAIT "Œ¡–¿¡Œ“ ¿..." WINDOW NOWAIT 
 SELECT people
 INDEX ON lpu_id TAG lpu_id
 SET ORDER TO lpu_id
 SCAN 
  m.recid   = PADL(RECNO(),7,'0')
  m.lpu_id  = lpu_id
  m.date_in = date_in
  m.spos    = spos
  m.s_pol   = s_pol2
  m.n_pol   = n_pol2
  m.tip_d   = IIF(tip_d2='—', '1', '3')
  m.q       = m.qcod

  INSERT INTO prfile FROM MEMVAR 
 ENDSCAN 
 SET ORDER TO 
 DELETE TAG ALL 
 USE IN people 
 USE IN prfile
 WAIT CLEAR 

RETURN 
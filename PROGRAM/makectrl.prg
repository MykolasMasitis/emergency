PROCEDURE MakeCTRL(m.lcDir, m.mcod)
 PRIVATE lcdir,mcod
 
 m.thisid = lpuid
 m.mmy    = PADL(m.tmonth,2,'0')+RIGHT(STR(m.tyear,4),1)
 
 IF !fso.FolderExists(lcDir)
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcdir+'\calls.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcdir+'\e'+m.mcod+'.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(lcdir+'\ans_sv4.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(lcdir+'\calls', 'calls', 'shar', 'recid')>0
  IF USED('calls')
   USE IN calls
  ENDIF 
 ENDIF 
 IF OpenFile(lcdir+'\e'+m.mcod, 'serror', 'shar')>0
  USE IN calls
  IF USED('serror')
   USE IN serror
  ENDIF 
 ENDIF 
 IF OpenFile(lcdir+'\ans_sv4', 'answers', 'excl')>0
  USE IN calls
  USE IN serror
  IF USED('answers')
   USE IN answers
  ENDIF 
 ENDIF 
 
 IF fso.FileExists(lcdir+'\ctrl'+m.mmy+'.dbf')
  fso.DeleteFile(lcdir+'\ctrl'+m.mmy+'.dbf')
 ENDIF 
 fso.CopyFile(ptempl+'\ctrlmmy.dbf', lcdir+'\ctrl'+m.mmy+'.dbf', .t.)
 IF !fso.FileExists(lcdir+'\ctrl'+m.mmy+'.dbf')
  USE IN calls
  USE IN serror
  SELECT aisoms
  RETURN 
 ENDIF 
 
 SELECT answers
 INDEX on recid TAG recid
 SET ORDER TO recid
 =OpenFile(lcdir+'\ctrl'+m.mmy, 'ctrl', 'shar')
 
 SELECT calls
 SET RELATION TO PADL(recid,6,'0') INTO answers
 SELECT serror
 SET RELATION TO rid INTO calls
 SCAN 
  m.sp_id = 4708
  m.recid = calls.recid_lpu
  m.err = LEFT(c_err,2)
  m.comment = "Ошибка страховой принадлежности документа ОМС, в т.ч. категории пациента"
  m.smo = answers.q
  m.polis = IIF(!EMPTY(answers.s_pol), answers.s_pol+' '+answers.n_pol, answers.n_pol)
  
  INSERT INTO ctrl FROM MEMVAR 
  
 ENDSCAN 
 SET RELATION OFF INTO serror
 SELECT calls
 SET RELATION OFF INTO answers
 
 USE IN calls
 USE IN serror
 USE IN ctrl
 USE IN answers
 SELECT aisoms
RETURN 

PROCEDURE locexp(m.lcDir, m.mcod)
 PRIVATE m.lcdir,m.mcod
 
 m.lpuid = lpuid
 
 M.ECA = .T.

 IF !fso.FolderExists(m.lcDir)
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(m.lcDir + '\calls.dbf')
  RETURN .F.
 ENDIF 
 IF !fso.FileExists(m.lcDir + '\e'+m.mcod+'.dbf')
  RETURN .F.
 ENDIF 

 IF OpenFile(m.lcDir+'\calls', "calls", "SHARE")>0
  IF USED('calls')
   USE IN calls
  ENDIF 
  RETURN .F.
 ENDIF 
 IF OpenFile(m.lcDir+'\e'+m.mcod, "serror", "SHARE", 'rid')>0
  USE IN calls
  IF USED('serror')
   USE IN serror
  ENDIF 
  RETURN .F.
 ENDIF 
 
 IF RECCOUNT('calls')<=0
  USE IN calls 
  USE IN serror 
  RETURN .F.
 ENDIF 

 DELETE FOR SUBSTR(c_err,3,1)='A' IN serror
 
 m.mcod = mcod
 WAIT m.mcod+'...' WINDOW NOWAIT 
 
 SELECT calls
 m.sumdef = 0 
 m.pazdef = 0 
 m.callsdef = 0 
 SCAN
  IF M.ECA == .T. && Алгоритм EC
   IF !EMPTY(calls.ans_r)
    m.IsGood = IIF(calls.q = m.qcod, .T., .F.)
    IF IsVS(calls.sn_pol) AND LEFT(calls.sn_pol,2)=m.qcod
     IF USED('kms')
      m.vvs = INT(VAL(SUBSTR(ALLTRIM(calls.sn_pol),7)))
      IF SEEK(m.vvs, 'kms')
       m.IsGood = .t.
      ENDIF 
     ENDIF 
    ENDIF 
    IF IsGood == .f.                 
     m.polis = sn_pol
     m.recid = recid
     m.tar = tar 
     rval =InsError('S', 'ECA', m.recid)
     m.sumdef   = m.sumdef   + IIF(rval==.T., tar, 0)
     m.pazdef   = m.pazdef   + IIF(rval==.T., 1, 0)
     m.callsdef = m.callsdef + IIF(rval==.T., 1, 0)
    ENDIF 
   ENDIF 
  ENDIF 
 ENDSCAN 

 SELECT aisoms
 REPLACE sumdef WITH m.sumdef, pazdef WITH m.pazdef, callsdef WITH m.callsdef
 
 WAIT CLEAR 

RETURN 

FUNCTION InsError(WFile, cError, cRecId)
 IF WFile == 'R'
  IF !SEEK(cRecId, 'rError')
   INSERT INTO rError (f, c_err, rid) VALUES ('R', cError, cRecId)
  ENDIF
 ENDIF 
 IF WFile == 'S'
  IF !SEEK(cRecId, 'sError')
   INSERT INTO sError (f, c_err, rid) VALUES ('S', cError, cRecId)
   RETURN .T.
  ELSE 
   IF cError != sError.c_err
    INSERT INTO sError (f, c_err, rid) VALUES ('S', cError, cRecId)
   ENDIF
  ENDIF
 ENDIF 
RETURN .F.

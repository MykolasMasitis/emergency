FUNCTION chkBase

IsArcExists = fso.FolderExists(pArc)
IF IsArcExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pArc!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pArc)
 ENDIF 
ENDIF 

IF !fso.FolderExists(pArc+'\IPS')
 fso.CreateFolder(pArc+'\IPS')
ENDIF 

IF !fso.FolderExists(pArc+'\IPS\INPUT')
 fso.CreateFolder(pArc+'\IPS\INPUT')
ENDIF 

IF !fso.FolderExists(pArc+'\IPS\OUTPUT')
 fso.CreateFolder(pArc+'\IPS\OUTPUT')
ENDIF 

IF !fso.FolderExists(pArc+'\ERZ')
 fso.CreateFolder(pArc+'\ERZ')
ENDIF 

IF !fso.FolderExists(pArc+'\ERZ\INPUT')
 fso.CreateFolder(pArc+'\ERZ\INPUT')
ENDIF 

IF !fso.FolderExists(pArc+'\ERZ\OUTPUT')
 fso.CreateFolder(pArc+'\ERZ\OUTPUT')
ENDIF 

IsDirExists = fso.FolderExists(pBase)
IF IsDirExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pBase!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pBase)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pBase+'\'+gcPeriod)
IF IsDirExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pBase\&gcPeriod!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pBase+'\'+gcPeriod)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pDouble)
IF IsDirExists == .F.
 fso.CreateFolder(pDouble)
ENDIF 

IsDirExists = fso.FolderExists(pOut)
IF IsDirExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pOut!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pOut)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pTempl)
IF IsDirExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pTempl!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  fso.CreateFolder(pTempl)
 ENDIF 
ENDIF 

IsDirExists = fso.FolderExists(pTrash)
IF IsDirExists == .F.
 fso.CreateFolder(pTrash)
ENDIF 

IsFileExists = fso.FileExists(pCommon+'\UsrLpu.dbf')
IF IsFileExists == .F.
 CREATE TABLE &pCommon\UsrLpu (mcod c(7), lpu_id n(4), cokr c(2), usr n(2)) 
 APPEND FROM &pCommon\sprlpuxx
 REPLACE ALL usr WITH 1
 INDEX ON mcod TAG mcod 
 INDEX ON lpu_id TAG lpu_id
 USE 
ENDIF 

IsFileExists = fso.FileExists(pCommon+'\Users.dbf')
IF IsFileExists == .F.
 CREATE TABLE &pCommon\Users (name c(6), fam c(25), im c(20), ot c(20), fio c(40))
 INDEX ON name TAG name CANDIDATE  
 INSERT INTO Users (name) VALUES ('OMS')
 INSERT INTO Users (name) VALUES ('NSI')
 FOR bnm=1 TO 10
  uuser = 'USR'+PADL(bnm,3,'0')
  INSERT INTO Users (name) VALUES (uuser)
 ENDFOR  
 USE 
ENDIF 

IsDirExists = fso.FolderExists(pBase+'\'+gcPeriod+'\NSI')
IF IsDirExists == .F.
 IF MESSAGEBOX("нрясрярбсер дхпейрнпхъ" + CHR(13) + "&pBase\&gcPeriod\NSI!" + CHR(13) + "янгдюрэ?",4+32, "") == 6
  WAIT "янгдюеряъ кнйюкэмши мях..." WINDOW NOWAIT 
  fso.CopyFolder(pCommon, pBase+'\'+gcPeriod+'\NSI')
  WAIT CLEAR 
 ENDIF 
ENDIF 


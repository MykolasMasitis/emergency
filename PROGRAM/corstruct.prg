PROCEDURE CorStruct
 IF MESSAGEBOX('ÂÛ ÕÎÒÈÒÅ ÏÐÎÂÅÑÒÈ '+CHR(13)+CHR(10)+;
               'ÊÎÐÐÅÊÒÈÐÎÂÊÓ ÑÒÐÓÊÒÓÐÛ ÁÄ?!'+CHR(13)+CHR(10)+;
               '',4+48,'') != 6
  RETURN 
 ENDIF 

 IF MESSAGEBOX('ÂÛ ÀÁÑÎËÞÒÍÎ ÓÂÅÐÅÍÛ Â ÑÂÎÈÕ ÄÅÉÑÒÂÈßÕ?',4+48,'') != 6
  RETURN 
 ENDIF 

 ppdir  = pbase+'\'+m.gcperiod
 IF !fso.FolderExists(ppdir)
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÄÈÐÅÊÒÎÐÈß '+ppdir,0+16,'')  
  RETURN
 ENDIF 
 
 aisfile = ppdir+'\AisOms'
 IF !fso.FileExists(aisfile+'.dbf')
  MESSAGEBOX('ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË '+aisfile,0+16,'')  
  RETURN
 ENDIF 
 
 IF OpenFile(aisfile, 'AisOms', 'shared', 'mcod')>0
  RETURN 
 ENDIF 

 SELECT AisOms
 SCAN
  m.mcod = mcod
  m.lpuid = STR(lpuid,4)

  WAIT m.mcod WINDOW NOWAIT 

  IF !fso.FolderExists(ppdir+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppdir+'\'+m.mcod+'\calls.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppdir+'\'+m.mcod+'\teams.dbf')
   LOOP 
  ENDIF 
  IF !fso.FileExists(ppdir+'\'+m.mcod+'\e'+m.mcod+'.dbf')
   CREATE TABLE &ppdir\&mcod\e&mcod (f c(1), c_err c(3), rid i)
   INDEX FOR UPPER(f)='R' ON rid TAG rrid
   INDEX FOR UPPER(f)='S' ON rid TAG rid
   USE 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\calls', 'calls', 'excl')>0
   LOOP 
  ENDIF 
  IF OpenFile(ppdir+'\'+m.mcod+'\teams', 'teams', 'excl')>0
   USE IN People
   LOOP 
  ENDIF 
  
  SELECT calls
  *IF FSIZE('c_i')!=25
  * ALTER TABLE Talon ALTER COLUMN c_i c(25)
  *ENDIF 
  *IF FIELD('IsPr')!='ISPR'
  * ALTER TABLE Talon ADD COLUMN IsPr L
  *ENDIF 
  USE 
  
  SELECT teams
  USE 

 ENDSCAN 
 USE 
 
 WAIT CLEAR 

 MESSAGEBOX('OK!', 0+64, '')

RETURN 
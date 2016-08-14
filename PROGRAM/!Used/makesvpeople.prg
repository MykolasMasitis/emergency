PROCEDURE MakeSvPeople
 IF fso.FileExists(pbase+'\'+gcperiod+'\people.dbf')
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ÑÂÎÄÍÛÉ ÐÅÃÈÑÒÐ ÓÆÅ ÑÎÁÐÀÍ.'+CHR(13)+CHR(10)+;
   'ÕÎÒÈÒÅ ÏÅÐÅÔÎÐÌÈÐÎÂÀÒÜ ÅÃÎ?',4+32,'')=7
   RETURN 
  ENDIF 
 ENDIF 

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÂÛ ÕÎÒÈÒÅ ÑÎÁÐÀÒÜ'+CHR(13)+CHR(10)+;
  'ÑÂÎÄÍÛÉ ÐÅÃÈÑÒÐ?'+CHR(13)+CHR(10),4+32,'')=7
  RETURN 
 ENDIF 

 IF MESSAGEBOX(CHR(13)+CHR(10)+'ÏÎÑËÅ ÑÎÇÄÀÍÈß ÑÂÎÄÍÎÃÎ ÐÅÃÈÑÒÐÀ'+CHR(13)+CHR(10)+;
  'ÄÀËÜÍÅÉØÈÉ ÏÐÈÅÌ ÏÎÑÛËÎÊ ÑÒÀÍÅÒ ÍÅÂÎÇÌÎÆÅÍ!'+CHR(13)+CHR(10)+'ÏÐÎÄÎËÆÈÒÜ?'+;
  CHR(13)+CHR(10),4+32,'') = 7
  RETURN 
 ENDIF 
  
 IF !fso.FolderExists(pbase)
  =MESSAGEBOX(CHR(13)+CHR(10)+'ÄÈÐÅÊÒÎÐÈß '+UPPER(ALLTRIM(PBASE))+ 'ÎÒÑÓÒÑÒÂÓÅÒ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(pbase+'\'+gcperiod)
  =MESSAGEBOX(CHR(13)+CHR(10)+'ÄÈÐÅÊÒÎÐÈß '+UPPER(ALLTRIM(pbase+'\'+gcperiod))+ 'ÎÒÑÓÒÑÒÂÓÅÒ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 

 IF !fso.FileExists(pbase+'\'+gcperiod+'\aisoms.dbf')
  =MESSAGEBOX(CHR(13)+CHR(10)+'ÎÒÑÓÒÑÒÂÓÅÒ ÔÀÉË '+UPPER(ALLTRIM(pbase+'\'+gcperiod+'\aisoms.dbf'))+ '!'+;
   CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(pBase+'\'+gcPeriod+'\aisoms', 'aisoms', 'shar') > 0
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\sprlpuxx', "sprlpu", "shar", "mcod") > 0
  IF USED('sprlpu')
   USE IN sprlpu
  ENDIF 
  IF USED('aisoms')
   USE IN aisoms
  ENDIF 
  RETURN
 ENDIF 

 CREATE CURSOR svppl ;
  (lpu_id n(6), mcod c(7), RecId i, date_in d, date_in2 d, date_out d, ;
   spos c(1), spos2 c(1), s_pol c(6), s_pol2 c(6), n_pol c(16), n_pol2 c(16), ;
   tip_d c(1), tip_d2 c(1), q c(2), fam c(25), im c(20), ot c(20), dr d, w n(1), ;
   ans_r c(3))
 SELECT svppl
 INDEX on n_pol2 TAG n_pol2
 SET ORDER TO n_pol2
 
 SELECT aisoms
 SCAN 
  IF EMPTY(bname)
   LOOP 
  ENDIF 
  m.mcod = mcod 
  IF !fso.FolderExists(pbase+'\'+gcperiod+'\'+m.mcod)
   LOOP 
  ENDIF 
  IF !fso.FileExists(pbase+'\'+gcperiod+'\'+m.mcod+'\people.dbf')
   LOOP 
  ENDIF 
  IF OpenFile(pBase+'\'+gcPeriod+'\'+m.mcod+'\people.dbf', 'people', 'shar') > 0
   IF USED('people')
    USE IN people
    SELECT aisoms
   ENDIF 
   LOOP 
  ENDIF 
  IF RECCOUNT('people')=0
   USE IN people 
   SELECT aisoms
   LOOP 
  ENDIF 
  
  WAIT m.mcod+'...' WINDOW NOWAIT 
  
  SELECT people

  SCAN 
   IF c_err != 'OK'
    LOOP 
   ENDIF 

   SCATTER MEMVAR 
   m.lpu_id = IIF(SEEK(m.mcod, 'sprlpu'), sprlpu.lpu_id, 0)
   
   IF SEEK(m.n_pol2, 'svppl')
    DO CASE 
     CASE spos = '1' AND svppl.spos = '1' && Îáà ïî òåððèòîðèè
      IF date_in>svppl.date_in
       m.dmcod = svppl.mcod
       m.drecid = svppl.recid
       IF fso.FolderExists(pbase+'\'+gcperiod+'\'+m.dmcod)
        IF fso.FileExists(pbase+'\'+gcperiod+'\'+m.dmcod+'\people.dbf')
         IF OpenFile(pbase+'\'+gcperiod+'\'+m.dmcod+'\people', 'ppeople', 'shar', 'recid')>0
          IF USED('ppeople')
           USE IN ppeople
          ENDIF 
         ELSE 
          REPLACE c_err WITH 'F9', lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in
*          REPLACE c_err WITH 'FA', lpu_id WITH svppl.lpu_id
*          MESSAGEBOX('FA',0+64,'')
          USE IN ppeople
         ENDIF 
        ENDIF 
       ENDIF 
       DELETE IN svppl
       INSERT INTO svppl FROM MEMVAR 
      ELSE 
       REPLACE c_err WITH 'F9', lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in
      ENDIF 

     CASE spos = '1' AND svppl.spos = '2'
      REPLACE c_err WITH 'F9', lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in

     CASE spos = '2' AND svppl.spos = '1'
       m.dmcod = svppl.mcod
       m.drecid = svppl.recid
       IF fso.FolderExists(pbase+'\'+gcperiod+'\'+m.dmcod)
        IF fso.FileExists(pbase+'\'+gcperiod+'\'+m.dmcod+'\people.dbf')
         IF OpenFile(pbase+'\'+gcperiod+'\'+m.dmcod+'\people', 'ppeople', 'shar', 'recid')>0
          IF USED('ppeople')
           USE IN ppeople
          ENDIF 
         ELSE 
          REPLACE c_err WITH 'F9', lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in
*          REPLACE c_err WITH 'FA'
*          MESSAGEBOX('FA',0+64,'')
          USE IN ppeople
         ENDIF 
        ENDIF 
       ENDIF 
      DELETE IN svppl
      INSERT INTO svppl FROM MEMVAR 

     CASE spos = '2' AND spos2 = '2' && Îáà ïî çàÿëâåíèþ
      IF date_in<svppl.date_in
       m.dmcod = svppl.mcod
       m.drecid = svppl.recid
       IF fso.FolderExists(pbase+'\'+gcperiod+'\'+m.dmcod)
        IF fso.FileExists(pbase+'\'+gcperiod+'\'+m.dmcod+'\people.dbf')
         IF OpenFile(pbase+'\'+gcperiod+'\'+m.dmcod+'\people', 'ppeople', 'shar', 'recid')>0
          IF USED('ppeople')
           USE IN ppeople
          ENDIF 
         ELSE 
          REPLACE c_err WITH 'F9', lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in
*          REPLACE c_err WITH 'FA'
*          MESSAGEBOX('FA',0+64,'')
          USE IN ppeople
         ENDIF 
        ENDIF 
       ENDIF 
       DELETE IN svppl
       INSERT INTO svppl FROM MEMVAR 
      ELSE 
       REPLACE c_err WITH 'F9', lpu_id WITH svppl.lpu_id, spos2 WITH svppl.spos, date_in2 WITH svppl.date_in
      ENDIF 
     OTHERWISE 
    ENDCASE 
   ELSE
    INSERT INTO svppl FROM MEMVAR 
   ENDIF 
   
  ENDSCAN 
  USE IN people 
  
  WAIT CLEAR 
  
  SELECT aisoms

 ENDSCAN 
 USE 
 USE IN sprlpu
 
 SELECT svppl 
 IF RECCOUNT('svppl')<=0
  USE IN svppl
  =MESSAGEBOX(CHR(13)+CHR(10)+'ÍÅ ÎÁÍÀÐÓÆÅÍÎ ÍÈ ÎÄÍÎÃÎ ×ÅËÎÂÅÊÀ!'+CHR(13)+CHR(10),0+64,'')
  RETURN 
 ENDIF  
 
 svpeople = pbase+'\'+gcperiod+'\people'
 IF fso.FileExists(svpeople+'.dbf')
  fso.DeleteFile(svpeople+'.dbf')
 ENDIF 
 CREATE TABLE (svpeople) ;
  (lpu_id n(6), mcod c(7), RecId i, date_in d, date_in2 d, date_out d, ;
   spos c(1), spos2 c(1), s_pol c(6), s_pol2 c(6), n_pol c(16), n_pol2 c(16), ;
   tip_d c(1), tip_d2 c(1), q c(2), fam c(25), im c(20), ot c(20), dr d, w n(1), ;
   ans_r c(3))
 
 SELECT svppl 
 SCAN FOR !DELETED()
  SCATTER MEMVAR 
  INSERT INTO people FROM MEMVAR 
 ENDSCAN 
 USE IN svppl
 USE IN people

 =MESSAGEBOX(CHR(13)+CHR(10)+'ÑÂÎÄÍÛÉ ÐÅÃÈÑÒÐ ÑÎÁÐÀÍ!'+CHR(13)+CHR(10),0+64,'')
 
RETURN 
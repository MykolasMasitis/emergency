PROCEDURE Flk2Lpu

 OldEscStatus = SET("Escape")
 SET ESCAPE OFF 
 CLEAR TYPEAHEAD 

 IF OpenFile(pBase+'\&gcPeriod\AisOms', 'AisOms', 'shar')>0
  RETURN .f. 
 ENDIF 
 
 =OpenFile(pbase+'\'+gcperiod+'\'+'nsi'+'\spraboxx', 'sprabo', 'shar')
  
 
 SELECT AisOms
 SCAN
  m.mcod = mcod 
  WAIT m.mcod WINDOW NOWAIT 

  lcPath = pBase+'\'+m.gcperiod+'\'+mcod
  IF !fso.FolderExists(lcPath)
   LOOP 
  ENDIF 

  m.mmy = PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
  m.ctrl = 'ctrl'+m.mmy+'.dbf'
  m.prsp = 'prsp'+m.qcod+m.mmy+'.pdf'
  IF !fso.FileExists(lcPath+'\'+m.ctrl)
   LOOP 
  ENDIF 
  IF !fso.FileExists(lcPath+'\'+m.prsp)
   LOOP 
  ENDIF 
  
  DocName = pbase+'\'+m.gcperiod+'\'+mcod+'\b_flk_' + mcod

  lcLpuID = lpuid
  m.cTO  = IIF(!EMPTY(ALLTRIM(cfrom)), ALLTRIM(cfrom), ;
  IIF(SEEK(lcLpuID, 'sprabo', 'lpu_id'), 'OMS@'+ALLTRIM(sprabo.abn_name), ''))
  m.un_id    = SYS(3)
  m.bansfile = 'b_flk_'  + mcod
  m.tansfile = 't_flk_'  + mcod
  m.d1file   = 'd1_flk_' + mcod
  m.d2file   = 'd2_flk_' + mcod
  m.mmid     = m.un_id+'.OMS@'+m.qmail
  m.csubj    = 'OMS#'+m.gcperiod+'###SMPM'

  poi = fso.CreateTextFile(lcPath + '\' + m.tansfile)

  poi.WriteLine('To: '+m.cTO)
  poi.WriteLine('Message-Id: ' + m.mmid)
  poi.WriteLine('Subject: ' + m.csubj)
  poi.WriteLine('Content-Type: multipart/mixed')
  poi.WriteLine('Resent-Message-Id: '+ALLTRIM(cmessage))
  poi.WriteLine('Attachment: '+m.d1file+' '+m.ctrl)
  poi.WriteLine('Attachment: '+m.d2file+' '+m.prsp)
 
  poi.Close
  
  IF fso.FileExists(lcPath+'\'+m.ctrl) 
   fso.CopyFile(lcPath+'\'+m.ctrl, pAisOms+'\oms\output\'+m.d1file)
  ENDIF 
  IF fso.FileExists(lcPath+'\'+m.prsp) 
   fso.CopyFile(lcPath+'\'+m.prsp, pAisOms+'\oms\output\'+m.d2file)
  ENDIF 

  fso.CopyFile(lcPath+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

  fso.CopyFile(lcPath+'\'+m.tansfile, lcPath+'\'+m.bansfile)
  fso.DeleteFile(lcPath+'\'+m.tansfile)
  
  IF CHRSAW(0) 
   IF INKEY() == 27
    IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
     EXIT 
    ENDIF 
   ENDIF 
  ENDIF 
 
 ENDSCAN 

 WAIT CLEAR 

 USE
 USE IN sprabo
 
 SET ESCAPE &OldEscStatus
RETURN 
 

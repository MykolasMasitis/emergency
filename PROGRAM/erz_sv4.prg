FUNCTION erz_sv4(lcDir, mcod, IsOK)
 PRIVATE lcdir,mcod,isok

 IF !fso.FolderExists(lcDir)
  RETURN .F.
 ENDIF 

 IF !fso.FileExists(lcDir + '\calls.dbf')
  RETURN .F.
 ENDIF 

 IF OpenFile("&lcDir\calls", "calls", "share")>0
  RETURN .F.
 ENDIF 
 
 IF RECCOUNT('calls')<=0
  USE IN calls
  RETURN .F.
 ENDIF 
 
 fso.CopyFile(ptempl+'\sverka4.dbf', lcDir+'\sverka4.dbf', .t.)
 oSettings.CodePage(lcDir+'\sverka4.dbf', 866, .t.)
 
 IF OpenFile("&lcDir\sverka4", "Zapros", "SHARE")>0
  USE IN calls
  RETURN .F.
 ENDIF 

 SELECT calls
 SCAN 
  m.recid  = PADL(recid,6,'0')
  m.sn_pol = ALLTRIM(sn_pol)
  m.tip    = tip
  m.tip_d  = tip
  DO CASE 
   CASE m.tip = '�' && ��������
    m.s_pol = ''
    m.n_pol = m.sn_pol
   CASE m.tip = '�' && ���
    m.s_pol = SUBSTR(m.sn_pol,1,6)
    m.n_pol = SUBSTR(m.sn_pol,7)
   CASE m.tip = '�' && ���
    m.s_pol = ''
    m.n_pol = m.sn_pol
   OTHERWISE 
    m.s_pol = ''
    m.n_pol = m.sn_pol
  ENDCASE 
*  m.date_in  = m.tdat1
*  m.date_out = m.tdat2
  m.date_in  = d_u
  m.date_out = d_u
  m.q = m.qcod
  m.fam = fam
  m.im = im
  m.ot = ot
  m.dr = DTOS(dr)
  m.w = w

  INSERT INTO Zapros FROM MEMVAR 
    
 ENDSCAN 
 USE IN Zapros
 USE IN Calls

 SELECT AisOms
   
 ChVal  = SYS(3)
 ID     = ALLTRIM(ChVal+'.'+m.gcUser+'@'+m.qmail)
  
 TFile  = 'terz_' + mcod
 BFile  = 'berz_' + mcod
 DFile  = 'derz_' + mcod

 iii = 1
 DO WHILE fso.FileExists(pAisOms+'\'+m.gcUser+'\OUTPUT\'+m.bfile)
  m.tfile  = 'terz_' + m.mcod + '_' + PADL(iii,2,'0')
  m.bfile  = 'berz_' + m.mcod + '_' + PADL(iii,2,'0')
  m.dfile  = 'derz_' + m.mcod + '_' + PADL(iii,2,'0')
  iii = iii + 1
 ENDDO 
   
 fso.CopyFile(lcDir+'\sverka4.dbf', PAisOms+'\'+m.gcUser+'\OutPut\'+DFile)
 fso.CopyFile(lcDir+'\sverka4.dbf', pArc+'\ERZ\OUTPUT\'+DFile)
 fso.DeleteFile(lcDir+'\sverka4.dbf')
   
 poi = FCREATE('&PAisOms\&gcUser\OutPut\&TFile')
 IF poi != -1
  =FPUTS(poi,'To: erz@mgf.msk.oms')
  =FPUTS(poi,'Message-Id: &ID')
  =FPUTS(poi,'Subject: ERZ_sverka4n')
  fzz = 'q_' + PADL(MONTH(DATE()),2,'0')+RIGHT(ALLTRIM(STR(YEAR(DATE()))),1)+'.dbf'
  =FPUTS(poi,'Attachment: &DFile &Fzz')
 ENDIF 
 =FCLOSE(poi)
 
 oTFile = fso.GetFile('&PAisOms\&gcUser\OutPut\&TFile')
 fso.CopyFile(PAisOms+'\'+gcUser+'\OutPut\'+TFile, pArc+'\ERZ\OUTPUT\'+BFile)
 oTFile.Move('&PAisOms\&gcUser\OutPut\&BFile')

* REPLACE erz_id WITH m.id
 REPLACE erzid WITH m.id, erzst WITH 1

 IF IsOk==.t.
  MESSAGEBOX('������ ���������!', 0+64, '')
 ENDIF 
RETURN .T.
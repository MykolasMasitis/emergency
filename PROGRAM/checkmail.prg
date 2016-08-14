PROCEDURE CheckMail
PARAMETERS lcUser

IF !IsAisDir() && �������� ������� ����������, OMS, INPUT, OUTPUT
 RETURN 
ENDIF 

oMailDir        = fso.GetFolder(pAisOms+'\&lcUser\input')
MailDirName     = oMailDir.Path
oFilesInMailDir = oMailDir.Files
nFilesInMailDir = oFilesInMailDir.Count

MESSAGEBOX('���������� '+ALLTRIM(STR(nFilesInMailDir))+' ������!', 0+64, lcUser)

IF nFilesInMailDir<=0
 RETURN 
ENDIF 

IF OpenTemplates() != 0
 =CloseTemplates() 
 RETURN 
ENDIF 

WAIT "�������� �����..." WINDOW NOWAIT 
SELECT AisOms
prvorder = ORDER('aisoms')
SET ORDER TO 

m.un_id = SYS(3)

OldEscStatus = SET("Escape")
SET ESCAPE OFF 
CLEAR TYPEAHEAD 

FOR EACH oFileInMailDir IN oFilesInMailDir

 SCATTER MEMVAR BLANK
 m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)

 m.BFullName = oFileInMailDir.Path
 m.bname     = oFileInMailDir.Name
 m.recieved  = oFileInMailDir.DateLastModified
 m.lpuid     = 0
 m.processed = DATETIME()
 
 m.cfrom      = ''
 m.cdate      = ''
 m.cmessage   = ''
 m.resmesid   = ''
 m.csubject   = ''
 m.csubject1  = ''
 m.csubject2  = ''
 m.attachment = ''
 m.bodypart   = ''

 m.attaches   = 0 && ������� �������������� ������ � ����� ��
 DIMENSION dattaches(10,2)
 dattaches = ''

 m.bparts   = 0 && ������� �������������� ������ � ����� ��
 DIMENSION dbparts(10,2)
 dbparts = ''

 DO CASE 
 CASE LOWER(oFileInMailDir.Name) = 'b'

 CFG = FOPEN(m.BFullName)
 =ReadCFGFile()
 = FCLOSE (CFG)

 IF UPPER(RIGHT((ALLTRIM(m.csubject)),4)) != 'SMPM'
  LOOP 
 ENDIF 
* MESSAGEBOX('OK!',0+64,'')

 m.sent = dt2date(m.cdate)
 WAIT m.cfrom WINDOW NOWAIT 
   
 m.llIsSubject = .F.

 m.adresat = PADR(LOWER(SUBSTR(m.cfrom,AT('@',m.cfrom)+1)),27)
 m.lpuid   = IIF(SEEK(m.adresat, "sprabo"), sprabo.object_id, m.lpuid)
 
 IF m.lpuid==0
*  MESSAGEBOX('������� '+UPPER(ALLTRIM(m.adresat))+' �� ������ � ����������� SPRABOXX.DBF!',0+48,lcUser)
  LOOP 
 ENDIF 
 
 m.mcod    = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.mcod, "")

 IF EMPTY(m.mcod)
*  MESSAGEBOX('������� '+UPPER(ALLTRIM(m.adresat))+' �� ������ � ����������� SPRLPUXX.DBF!',0+48,lcUser)
  LOOP 
 ENDIF 

 m.cokr    = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.cokr, "")
 m.moname  = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.name, "")
 m.usr     = IIF(SEEK(m.lpuid, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")

 IF EMPTY(m.usr) AND m.gcUser!='OMS'
*  MESSAGEBOX('��� '+m.mcod+' �� "���������" � ����������� � USRLPU.DBF!',0+48,lcUser)
  LOOP 
 ENDIF 
 
 IF m.usr != m.gcUser AND m.gcUser!='OMS'
  LOOP 
 ENDIF 

 m.previous_id = m.un_id
 m.un_id     = SYS(3)
 DO WHILE m.un_id = m.previous_id
  m.un_id     = SYS(3)
 ENDDO 
 m.previous_id = m.un_id

 m.tansfile  = 'tok_' + m.mcod
 m.bansfile  = 'bok_' + m.mcod
 iii = 1
 DO WHILE fso.FileExists(pAisOms+'\OMS\OUTPUT\'+m.bansfile)
  m.tansfile  = 'tok_' + m.mcod + '_' + PADL(iii,2,'0')
  m.bansfile  = 'bok_' + m.mcod + '_' + PADL(iii,2,'0')
  iii = iii + 1
 ENDDO 

 m.messageid = ALLTRIM(m.un_id+'.'+m.gcUser+'@'+m.qmail)

 && ������������ �� ���-������ � �����? ���� ���, �� - � ����!
 IF m.attaches == 0 AND m.bparts == 0 && ���� � ����� ������ �� ������������!
  TextToWrite="MyComment: � ����� ������ �� ������������"
  fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  fso.DeleteFile(m.BFullName)
  =WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  LOOP 
 ENDIF 
 && ������������ �� ���-������ � �����? ���� ���, �� - � ����!

 && �������� ������������� �������
 IsComplect = .T.
 IF m.attaches>0
  FOR natt = 1 TO m.attaches
   IF !fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    IsComplect = .F.
*    MESSAGEBOX('�������������� � ����� '+m.bname+CHR(13)+CHR(10)+;
     ' ATTACHMENT '+dattaches(natt,1)+ ' �����������!', 0+48, lcUser)
    LOOP 
   ENDIF 
  ENDFOR 
 ENDIF 
 IF IsComplect = .F.
  FOR natt = 1 TO m.attaches
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1), SpamDir+'\'+dattaches(natt,1), .t.)
    fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
   ENDIF 
  ENDFOR 
  TextToWrite="MyComment: ����������� ��� ���������� �������������� ����"
  fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  fso.DeleteFile(m.BFullName)
  =WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  LOOP 
 ENDIF  

 IsComplect = .T.
 IF m.bparts>0
  FOR natt = 1 TO m.bparts
   IF !fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
    IsComplect = .F.
*    MESSAGEBOX('�������������� � ����� '+m.bname+CHR(13)+CHR(10)+;
     ' DODY-PART '+dbparts(natt,1)+ ' �����������!', 0+48, lcUser)
    LOOP 
   ENDIF 
  ENDFOR 
 ENDIF 
 IF IsComplect = .F.
  FOR natt = 1 TO m.bparts
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1), SpamDir+'\'+dbparts(natt,1), .t.)
    fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
   ENDIF 
  ENDFOR 
  TextToWrite="MyComment: ����������� ��� ���������� �������������� ����"
  fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  fso.DeleteFile(m.BFullName)
  =WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  LOOP 
 ENDIF  
 && �������� ������������� �������

 && ���� �� ���� �� ���� zip-�����?
 llIsOneZip = .F.
 FOR nattach = 1 TO m.attaches

  m.dname   = ALLTRIM(dattaches(nattach, 1))
  m.attname = ALLTRIM(dattaches(nattach, 2))
  ffile = fso.GetFile(MailDirName + '\' + m.dname)
  IF ffile.size >= 2
   fhandl = ffile.OpenAsTextStream
   lcHead = fhandl.Read(2)
   fhandl.Close
  ELSE 
   lcHead = ''
  ENDIF 

  IF lcHead == 'PK' && ��� zip-����!
   ZipName = pAisOms+'\'+lcUser+'\input\'+m.dname
   IsSuccess=UnzipOpen(ZipName)
   IF IsSuccess
    llIsOneZip = .t.
    UnzipClose()
    EXIT 
   ENDIF 
   UnzipClose()
  ENDIF 

 ENDFOR 
 && ���� �� ���� �� ���� zip-�����?
 
 IF llIsOneZip == .F.
  TextToWrite="MyComment: ����� �������������� ������ ��� �� ������ zip-������"
  fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  fso.DeleteFile(m.BFullName)
  =WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  IF m.attaches>0
   FOR nattach = 1 TO m.attaches
    m.dname   = ALLTRIM(dattaches(nattach, 1))
    IF !EMPTY(m.dname)
     fso.CopyFile(MailDirName + '\' + m.dname, SpamDir+'\'+m.dname)
     fso.DeleteFile(MailDirName + '\' + m.dname)
    ENDIF 
   ENDFOR 
  ENDIF 
  IF m.bparts > 0
   FOR npart = 1 TO m.bparts
    m.bpname   = ALLTRIM(dbparts(npart, 1))
    IF !EMPTY(m.bpname)
     fso.CopyFile(MailDirName + '\' + m.bpname, SpamDir+'\'+m.bpname, .t.)
     fso.DeleteFile(MailDirName + '\' + m.bpname)
    ENDIF 
   ENDFOR 
  ENDIF 
  LOOP 
 ENDIF 
 && ���� ��� �� ������ zip-������

 poi = fso.CreateTextFile(pAisOms+'\&lcUser\output\'+m.tansfile)
 poi.WriteLine('To: '+m.cfrom)
 poi.WriteLine('Message-Id: ' + m.messageid)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: ' + m.cmessage)

 && ��������� ������������� ������� - ������� 1 �����!
 UnzipOpen(ZipName)
* sspItem  = 'SSP'+'.'+m.mmy
 sspItem  = 'SSP'+m.qcod+'.'+m.mmy

* brspItem = 'BRSP'+STR(m.lpuid,4)+'.'+m.mmy
 brspItem = 'BRSP'+m.mmy+'.dbf'

* shspItem = 'ShSP'+m.qcod+m.mmy+'.pdf'
* shspItem = 'shsmp'+m.mmy+'.pdf'
 shspItem = 'shsp'+m.qcod+m.mmy+'.pdf'
 IF !IsIPComplete()
  LOOP 
 ENDIF 
 UnzipClose()
 && ��������� ������������� ������� - ������� 1 �����!

 && ���� ������� ���������
* IsThisIPDouble = IIF(SEEK(m.mcod, 'AisOms', 'mcod'), .t., .f.)
 IsThisIPDouble = IIF(SEEK(m.mcod, 'AisOms', 'mcod') AND !EMPTY(AisOms.Sent), .t., .f.)
 IF IsThisIPDouble
  frst_time = AisOms.Sent
  IF frst_time > m.sent && �������� ����� ������� ���������� ����� ������������!
   m.csubject = m.csubject1 + '99' +m.csubject2
   m.cerrmessage = [��� ��������� ����� ������ �������]
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pAisOms+'\&lcUser\OUTPUT\'+m.bansfile)
   fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pArc+'\IPS\OUTPUT\'+m.bansfile)
   fso.DeleteFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile)
   DoubleDir = pDouble + '\' + m.mcod
   IF !fso.FolderExists(DoubleDir)
    fso.CreateFolder(DoubleDir)
   ENDIF 

   fso.CopyFile(m.BFullName, DoubleDir+'\'+m.bname)
   fso.DeleteFile(m.BFullName)
   IF m.attaches>0
    FOR nattach = 1 TO m.attaches
     m.dname   = ALLTRIM(dattaches(nattach, 1))
     IF !EMPTY(m.dname)
      fso.CopyFile(MailDirName + '\' + m.dname, pDouble+'\'+m.mcod+'\'+m.dname)
      fso.DeleteFile(MailDirName + '\' + m.dname)
     ENDIF 
    ENDFOR 
   ENDIF 
   IF m.bparts > 0
    FOR npart = 1 TO m.bparts
     m.bpname   = ALLTRIM(dbparts(npart, 1))
     IF !EMPTY(m.bpname)
      fso.CopyFile(MailDirName + '\' + m.bpname, pDouble+'\'+m.mcod+'\'+m.dname)
      fso.DeleteFile(MailDirName + '\' + m.bpname)
     ENDIF 
    ENDFOR 
   ENDIF 

   IF NOT SEEK(m.cmessage, "daisoms")
    INSERT INTO daisoms FROM MEMVAR 
   ENDIF

   LOOP 

  ELSE                  && ������������ ������� ����� ������, ��� �������� �����!

   m.prv_bfile = ALLTRIM(AisOms.bname)
   DoubleDir   = pDouble + '\' + m.mcod
   IF !fso.FolderExists(DoubleDir)
    fso.CreateFolder(DoubleDir)
   ENDIF 
   
   CFG = FOPEN(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.prv_bfile)
   DO WHILE NOT FEOF(CFG)
    READCFG = FGETS (CFG)
    IF UPPER(READCFG) = 'ATTACHMENT'
     m.dbl_attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
     m.dbl_dname      = ALLTRIM(SUBSTR(m.dbl_attachment, 1, AT(" ",m.dbl_attachment)-1)) && �������� d-�����
     m.dbl_attname    = ALLTRIM(SUBSTR(m.dbl_attachment, AT(" ",m.dbl_attachment)+1))    && ����������� �������� �����
     IF !EMPTY(m.dbl_attname)
      fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.dbl_attname, DoubleDir+'\'+m.dbl_dname)
     ENDIF 
    ENDIF 
   ENDDO
   = FCLOSE (CFG)
   fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.prv_bfile, DoubleDir+'\'+m.prv_bfile, .t.)
   
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\*.*')

   m.t_BName      = AisOms.BName
   m.t_Sent       = AisOms.Sent
   m.t_Recieved   = AisOms.Recieved
   m.t_Processed  = AisOms.Processed
   m.t_CMessage   = AisOms.CMessage
   m.t_dname      = AisOms.dname
*   DELETE IN AisOms && !!!
   
   MailView.get_recs.value = MailView.get_recs.value - 1

   IF NOT SEEK(m.t_CMessage, "daisoms")

    INSERT INTO daisoms (LpuId,Mcod,BName,Sent,Recieved,Processed,CFrom,CMessage,dname ) ;
     VALUES ;
     (m.lpuid,m.mcod,m.t_BName,m.t_Sent,m.t_Recieved,m.t_Processed,m.cfrom,m.t_CMessage,m.t_dname)
   ENDIF

   RELEASE m.t_BName,m.t_Sent,m.t_Recieved,m.t_Processed,m.t_CMessage,m.t_dname

  ENDIF 
 ENDIF 
 && ���� ������� ���������

 && ���� ��� ���������� � ����� �������!
 InDirPeriod = pBase + '\' + m.gcPeriod
 IF !fso.FolderExists(InDirPeriod)
  fso.CreateFolder(InDirPeriod)
 ENDIF 
 InDir = pBase + '\' + m.gcPeriod + '\' + m.mcod
 IF !fso.FolderExists(InDir)
  fso.CreateFolder(InDir)
 ENDIF 

 fso.CopyFile(m.BFullName, InDir + '\' + m.bname)
 fso.CopyFile(m.BFullName, pArc + '\IPS\INPUT\' + m.bname)
 fso.DeleteFile(m.BFullName)
 IF m.attaches>0
  FOR nattach = 1 TO m.attaches
   m.ddname   = ALLTRIM(dattaches(nattach, 1))
   m.aattname = ALLTRIM(dattaches(nattach, 2))
   IF !EMPTY(m.dname)
    fso.CopyFile(MailDirName + '\' + m.ddname, InDir+'\'+m.aattname, .t.)
    fso.CopyFile(MailDirName + '\' + m.ddname, parc+'\IPS\INPUT\'+m.ddname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.ddname)
   ENDIF 
  ENDFOR 
 ENDIF 
 IF m.bparts > 0
  FOR npart = 1 TO m.bparts
   m.bpname   = ALLTRIM(dbparts(npart, 1))
   IF !EMPTY(m.bpname)
    fso.CopyFile(MailDirName + '\' + m.bpname, InDir+'\'+m.bpname, .t.)
    fso.CopyFile(MailDirName + '\' + m.bpname, InDir+'\'+m.bpname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.bpname)
   ENDIF 
  ENDFOR 
 ENDIF 

 ZipName = InDir + '\' + m.attname
 ZipDir  = InDir + '\'

 UnzipOpen(ZipName)
 UnzipGotoFileByName(sspItem)
 UnzipFile(ZipDir)
 UnzipGotoFileByName(brspItem)
 UnzipFile(ZipDir)
* UnzipGotoFileByName(ShSPItem)
* UnzipFile(ZipDir)
 UnzipClose()

 m.lcCurDir = pBase + '\' + m.gcPeriod + '\' + m.mcod+'\'
 SET DEFAULT TO (lcCurDir)

 =OpenFile("&sspItem",  "sspfile",  "SHARED")
 =OpenFile("&brspItem",  "brspfile",  "SHARED")
 
 IF !CheckFilesStucture()
  LOOP 
 ENDIF 
 
 m.csubject = m.csubject1 + '01' +m.csubject2
 poi.WriteLine('Subject: '+m.csubject)
 poi.WriteLine('BodyPart: OK' )
 poi.Close

 =CloseItems()

 fso.CopyFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile, pAisOms+'\'+lcUser+'\OUTPUT\'+bansfile)
 fso.CopyFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile, pArc+'\IPS\OUTPUT\'+bansfile)
 fso.DeleteFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile)
 
 lcDir  = m.pBase + '\' + m.gcPeriod + '\' + m.mcod
 fCalls = lcDir + '\Calls'
 Teams = lcDir + '\Teams'
 Errors = lcDir + '\e'+m.mcod

 m.paz      = 0
 m.calls    = 0
 m.s_all    = 0
 m.pazdef   = 0
 m.callsdef = 0
 m.sumdef   = 0

 =CreateFilesStructure()
 =OpenLocalFiles()
 =MakeCalls()
 =MakeTeams()

 m.opaz   = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.paz, 0)
 m.ocalls = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.calls, 0)
 m.os_all = IIF(SEEK(m.mcod, 'aisoms', 'mcod'), aisoms.s_all, 0)
 
 MailView.recs  = MailView.recs  + 1
 MailView.paz   = MailView.paz   + (m.paz - m.opaz)
 MailView.calls = MailView.calls + (m.calls - m.ocalls)
 MailView.s_all = MailView.s_all + (m.s_all - m.os_all)
 
 UPDATE aisoms SET bname=m.bname, dname=m.dname, sent=m.sent, recieved=m.recieved, ;
  processed=m.processed, cfrom=m.cfrom, cmessage=m.cmessage, erzid='', erzst=0, ;
  paz=m.paz, calls=m.calls, s_all=m.s_all, pazdef=0, callsdef=0, sumdef=0 WHERE mcod=m.mcod 

 SELECT AisOms
 
 WAIT CLEAR 
 MailView.Refresh 

 CASE LOWER(oFileInMailDir.Name) = 'r' && ���� ��� r-����

 CFG = FOPEN(m.BFullName)
 =ReadCFGFile()
 =FCLOSE (CFG)

 WAIT m.cfrom WINDOW NOWAIT 
   
 fso.CopyFile(m.BFullName, DaemonDir + '\' + m.bname, .t.)
 fso.DeleteFile(m.BFullName)

 IF m.attaches > 0
  FOR nattach = 1 TO m.attaches
   m.ddname   = ALLTRIM(dattaches(nattach, 1))
   m.aattname = ALLTRIM(dattaches(nattach, 2))
   IF !EMPTY(m.dname) AND fso.FileExists(MailDirName + '\' + m.ddname)
    fso.CopyFile(MailDirName + '\' + m.ddname, DaemonDir+'\'+m.aattname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.ddname)
   ENDIF 
  ENDFOR 
 ENDIF 

 IF m.bparts > 0
  FOR npart = 1 TO m.bparts
   m.bpname   = ALLTRIM(dbparts(npart, 1))
   IF !EMPTY(m.bpname) AND fso.FileExists(MailDirName + '\' + m.bpname)
    fso.CopyFile(MailDirName + '\' + m.bpname, DaemonDir+'\'+m.bpname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.bpname)
   ENDIF 
  ENDFOR 
 ENDIF 

 SELECT AisOms
 
 WAIT CLEAR 
 MailView.Refresh 

 ENDCASE 

 IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('�� ������ �������� ���������?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 

NEXT && ���� �� ������

SET ESCAPE &OldEscStatus

=CloseTemplates() 

SET ORDER TO (prvorder)
MailView.Refresh
MailView.LockScreen=.f.

WAIT CLEAR 

nFilesInMailDir = oFilesInMailDir.Count

MESSAGEBOX('�������� '+ALLTRIM(STR(nFilesInMailDir))+' �������������� ��!', 0+64, lcUser)

RETURN 

FUNCTION CopyToTrash(lcPath, nTip)
 fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pAisOms+'\&lcUser\OUTPUT\'+m.bansfile)
 fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pArc+'\IPS\OUTPUT\'+m.bansfile)
 fso.DeleteFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile)
 TrashDir = pTrash + '\' + m.mcod
 IF !fso.FolderExists(TrashDir)
  fso.CreateFolder(TrashDir)
 ENDIF 
 fso.CopyFile(lcPath + '\' + m.bname, TrashDir+'\'+m.bname)
 fso.DeleteFile(lcPath + '\' + m.bname)

 FOR nattach = 1 TO m.attaches
  IF !EMPTY(ALLTRIM(dattaches(nattach, 1)))
   fso.CopyFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)), TrashDir + '\' + ALLTRIM(dattaches(nattach, 2)))
   fso.DeleteFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)))
  ENDIF 
 ENDFOR 

 IF NOT SEEK(m.cmessage, "taisoms", "cmessage")
  INSERT INTO taisoms FROM MEMVAR 
 ENDIF
RETURN 

FUNCTION ClDir
 IF fso.FileExists(sspItem)
  DELETE FILE &sspItem
 ENDIF 
 IF fso.FileExists(brspItem)
  DELETE FILE &brspItem
 ENDIF 
RETURN 

FUNCTION OpenTemplates
 tn_result = 0
 tn_result = tn_result + OpenFile("&ptempl\SSPqq.mmy", "ssp_et", "SHARED")
 tn_result = tn_result + OpenFile("&ptempl\BRSPmmy.dbf", "brsp_et", "SHARED")
RETURN tn_result

FUNCTION CloseTemplates
 IF USED('ssp_et')
  USE IN ssp_et
 ENDIF 
 IF USED('brsp_et')
  USE IN brsp_et
 ENDIF 
RETURN 

FUNCTION CloseItems
 IF USED('sspfile')
  USE IN sspfile
 ENDIF 
 IF USED('brspfile')
  USE IN brspfile
 ENDIF 
RETURN 

FUNCTION CompFields(NameOfFile)
 FOR nFld = 1 TO fld_1
  IF (tabl_1(nFld,1) == tabl_2(nFld,1)) AND ;
     (tabl_1(nFld,2) == tabl_2(nFld,2)) AND ;
     (tabl_1(nFld,3) == tabl_2(nFld,3))
  ELSE 
   =CloseItems()
*   =ClDir()
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Wrong structure of ] + NameOfFile
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   RETURN 0 
  ENDIF 
 ENDFOR 
RETURN 1

FUNCTION DiffFields(NameOfFile)
 =CloseItems()
* =ClDir()
 m.csubject = m.csubject1 + '08' +m.csubject2
 m.cerrmessage = [Wrong number of fields in ] + NameOfFile
 IF m.llIsSubject = .F.
  m.llIsSubject = .T.
  poi.WriteLine('Subject: '+m.csubject)
 ENDIF 
 poi.WriteLine('BodyPart: ' + m.cerrmessage)
 poi.Close
RETURN 

FUNCTION IsAisDir()
 IF !fso.FolderExists(pAisOms)
  MESSAGEBOX('����������� ���������� '+pAisOms, 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser')
  MESSAGEBOX('����������� ���������� '+pAisOms+'\&lcUser', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\INPUT')
  MESSAGEBOX('����������� ���������� '+pAisOms+'\&lcUser\INPUT', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\OUTPUT')
  MESSAGEBOX('����������� ���������� '+pAisOms+'\&lcUser\OUTPUT', 0+16, '')
  RETURN .F. 
 ENDIF

RETURN .T. 

FUNCTION WriteInBFile(BFullName, TextToWrite)
 CFG = FOPEN(BFullName,12)
 IsMyCommentExists = .F.
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  IF UPPER(READCFG) = 'MYCOMMENT'
   IsMyCommentExists = .T.
   LOOP 
  ENDIF 
 ENDDO
 IF !IsMyCommentExists
  nFileSize = FSEEK(CFG,0,2)
  =FWRITE(CFG, TextToWrite)
 ENDIF 
 = FCLOSE (CFG)
RETURN 

FUNCTION MakeCalls
 tFile = lcDir+'\'+sspItem
 oSettings.CodePage('&tFile', 866, .t.)
 USE &lcDir\&sspItem IN 0 ALIAS lcRFile  EXCLUSIVE 
 SELECT lcRFile
 SCAN 
  SCATTER MEMVAR
  m.s_all = m.s_all + m.tar
  m.calls = m.calls + 1
  m.paz = m.paz + 1
  m.recid_lpu = m.recid
  m.q    = m.qq
  RELEASE m.recid
  INSERT INTO fCalls FROM MEMVAR
 ENDSCAN 
 USE 
 USE IN fCalls 
 fso.DeleteFile(lcDir+'\'+sspItem)
RETURN 

FUNCTION MakeTeams
 tFile = lcDir+'\'+brspItem
 oSettings.CodePage('&tFile', 866, .t.)
 USE &lcDir\&brspItem IN 0 ALIAS lcSFile  EXCLUSIVE 
 SELECT lcSFile
 SCAN 
  SCATTER MEMVAR
  INSERT INTO Teams FROM MEMVAR
 ENDSCAN 
 USE 
 USE IN Teams
 fso.DeleteFile(lcDir+'\'+brspItem)
RETURN 

FUNCTION CreateFilesStructure

 CREATE TABLE (fCalls) ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1, recid_lpu c(7), tip c(1), sn_pol c(25), q c(2), ans_r c(3),;
  sp_id n(6), pst c(3), c_br c(9), profbr c(3), d_u d, t_uz c(5), t_up c(5), n_u c(12), vid_u c(6), cod n(6), tar n(12,2), ;
  res n(2), ds c(6), fam c(25), im c(20), ot c(20), w n(1), dr d, c_okato c(5), v_doc n(2), s_doc c(9), n_doc c(8), ;
  territ n(1), xx c(40), err c(2), lpu_id n(6))
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX ON dr TAG dr
 USE 
 
 CREATE TABLE (Teams) ;
  (c_br c(9), pst c(3), name c(250))
 INDEX ON c_br TAG c_br
 INDEX ON pst TAG pst
 USE 

 CREATE TABLE (Errors) (f c(1), c_err c(3), rid i)
 INDEX FOR UPPER(f)='R' ON rid TAG rrid
 INDEX FOR UPPER(f)='S' ON rid TAG rid
 USE 

RETURN 

FUNCTION OpenLocalFiles
 USE (fCalls) IN 0 ALIAS fCalls SHARED
 USE (Teams) IN 0 ALIAS Teams SHARED
RETURN 

FUNCTION CheckFilesStucture

 fld_1 = AFIELDS(tabl_1, 'sspfile') && �������� r-�����
 fld_2 = AFIELDS(tabl_2, 'ssp_et')
 IF fld_1 == fld_2 && ���-�� ����� ���������!
  FieldsIdent = CompFields('sspfile') && 0 - ���� �������, 1 - ������ ����������
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields('sspfile')
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 

 fld_1 = AFIELDS(tabl_1, 'brspfile') && �������� r-�����
 fld_2 = AFIELDS(tabl_2, 'brsp_et')
 IF fld_1 == fld_2 && ���-�� ����� ���������!
  FieldsIdent = CompFields('brspfile') && 0 - ���� �������, 1 - ������ ����������
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields('brspfile')
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 

RETURN .T. 

FUNCTION IsIPComplete	
 DO CASE 
  CASE !UnzipGotoFileByName(sspItem)
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [����������� ] + sspItem + [ ����]
   UnzipClose()
   m.cmnt = m.cerrmessage
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   =CopyToTrash(m.MailDirName,1)
   RETURN .F. 

  CASE !UnzipGotoFileByName(brspItem)
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [����������� ] + brspItem + [ ����]
   UnzipClose()
   m.cmnt = m.cerrmessage
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   =CopyToTrash(m.MailDirName,1)
   RETURN .F. 

*  CASE !UnzipGotoFileByName(shspItem)
*   m.csubject = m.csubject1 + '08' +m.csubject2
*   m.cerrmessage = [����������� ] + shspItem + [ ����]
*   UnzipClose()
*   m.cmnt = m.cerrmessage
*   IF m.llIsSubject = .F.
*    m.llIsSubject = .T.
*    poi.WriteLine('Subject: '+m.csubject)
*   ENDIF 
*   poi.WriteLine('BodyPart: ' + m.cerrmessage)
*   poi.Close
*   =CopyToTrash(m.MailDirName,1)
*   RETURN .F. 
 ENDCASE
RETURN .T.

FUNCTION ReadCFGFile
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  DO CASE
   CASE UPPER(READCFG) = 'FROM'
    m.cfrom = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'DATE'
    m.cdate = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'MESSAGE'
    m.cmessage = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'RESENT-MESSAGE-ID'
    m.resmesid = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'SUBJECT'
    m.csubject = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    m.csubject1 = LEFT(m.csubject, RAT('#',m.csubject,2))   && ����� subject ��� ����������� ������� ���� ����������
    m.csubject2 = SUBSTR(m.csubject, RAT('#',m.csubject,1)) && ����� subject ��� ����������� ������� ���� ����������
   CASE UPPER(READCFG) = 'ATTACHMENT'
    m.attaches   = m.attaches + 1
    m.attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dattaches(m.attaches,1) = ALLTRIM(SUBSTR(m.attachment, 1, AT(" ",m.attachment)-1)) && �������� d-�����
    dattaches(m.attaches,2) = ALLTRIM(SUBSTR(m.attachment, AT(" ",m.attachment)+1))    && ����������� �������� �����
   CASE UPPER(READCFG) = 'BODYPART'
    m.bparts   = m.bparts + 1
    m.bodypart = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dbparts(m.bparts,1) = ALLTRIM(SUBSTR(m.bodypart, 1, AT(" ",m.bodypart)-1))
  ENDCASE
 ENDDO
RETURN 
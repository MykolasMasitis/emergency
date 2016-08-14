PROCEDURE Exp2Lpu
 IF MESSAGEBOX("œ≈–≈Õ≈—“» ƒ¿ÕÕ€≈ ¬ LPU2SMO?",4+32,"")=7
  RETURN 
 ENDIF 
 m.pLpu = ALLTRIM(m.plpubin)
 IF EMPTY(m.pLpu)
  MESSAGEBOX("Õ≈ ” ¿«¿Õ œ”“‹   LPU2SMO!",0+64,"")
 ENDIF 
 IF !fso.FolderExists(m.pLpu)
  MESSAGEBOX("œ”“‹ "+m.pLpu+" Õ≈ —”Ÿ≈—“¬”≈“!",0+64,"")
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.pLpu+'\lpu2smo.cfg')
  MESSAGEBOX("Õ≈ Õ¿…ƒ≈Õ  ŒÕ‘»√”–¿÷»ŒÕÕ€… ‘¿…À "+CHR(13)+CHR(10)+m.pLpu+'\lpu2smo.cfg!',0+64,"")
  RETURN 
 ENDIF 
 IF OpenFile(m.pLpu+'\lpu2smo.cfg', 'cfg', 'shar')>0
  IF USED('cfg')
   USE IN cfg
  ENDIF 
  RETURN 
 ENDIF 
 SELECT cfg
 m.lpubase = ALLTRIM(pbase)
 USE IN cfg 
 IF EMPTY(m.LpuBase)
  MESSAGEBOX("Õ≈ ” ¿«¿Õ œ”“‹   LPU2SMO\BASE!",0+64,"")
 ENDIF 
 IF !fso.FolderExists(m.LpuBase)
  MESSAGEBOX("œ”“‹ "+m.LpuBase+" Õ≈ —”Ÿ≈—“¬”≈“!",0+64,"")
  RETURN 
 ENDIF 
 m.lpubase = m.lpubase+'\'+m.gcperiod
 IF !fso.FolderExists(m.LpuBase)
  MESSAGEBOX("Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ œ≈–»Œƒ¿ "+CHR(13)+CHR(10)+m.LpuBase+" !",0+64,"")
  RETURN 
 ENDIF 
 IF !fso.FolderExists(m.lpubase+'\0371001')
  fso.CreateFolder(m.lpubase+'\0371001')
 ENDIF 
 
 m.people = m.lpubase+'\0371001\people'
 m.talon  = m.lpubase+'\0371001\talon'
 m.otdel  = m.lpubase+'\0371001\otdel'
 m.doctor = m.lpubase+'\0371001\doctor'
 m.eerror = m.lpubase+'\0371001\e0371001'
 m.merror = m.lpubase+'\0371001\m0371001'
 
 IF fso.FileExists(m.people+'.dbf')
  fso.DeleteFile(m.people+'.dbf')
 ENDIF 
 IF fso.FileExists(m.people+'.cdx')
  fso.DeleteFile(m.people+'.cdx')
 ENDIF 
 IF fso.FileExists(m.talon+'.dbf')
  fso.DeleteFile(m.talon+'.dbf')
 ENDIF 
 IF fso.FileExists(m.talon+'.cdx')
  fso.DeleteFile(m.talon+'.cdx')
 ENDIF 
 IF fso.FileExists(m.otdel+'.dbf')
  fso.DeleteFile(m.otdel+'.dbf')
 ENDIF 
 IF fso.FileExists(m.otdel+'.cdx')
  fso.DeleteFile(m.otdel+'.cdx')
 ENDIF 
 IF fso.FileExists(m.doctor+'.dbf')
  fso.DeleteFile(m.doctor+'.dbf')
 ENDIF 
 IF fso.FileExists(m.doctor+'.cdx')
  fso.DeleteFile(m.doctor+'.cdx')
 ENDIF 
 IF fso.FileExists(m.eerror+'.dbf')
  fso.DeleteFile(m.eerror+'.dbf')
 ENDIF 
 IF fso.FileExists(m.eerror+'.cdx')
  fso.DeleteFile(m.eerror+'.cdx')
 ENDIF 
 IF fso.FileExists(m.merror+'.dbf')
  fso.DeleteFile(m.merror+'.dbf')
 ENDIF 
 IF fso.FileExists(m.merror+'.cdx')
  fso.DeleteFile(m.merror+'.cdx')
 ENDIF 

 CREATE TABLE (People) ;
  (RecId i AUTOINC NEXTVALUE 1 STEP 1,;
   mcod c(7), prmcod c(7), period c(6), d_beg d, d_end d, s_all n(11,2), ;
   tip_p n(1), sn_pol c(25), tipp c(1), enp c(16), qq c(2), ;
   fam c(25), im c(20), ot c(20), w n(1), dr d, d_type c(1), ;
   sv c(3), recid_lpu c(7), IsPr L)
 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON sn_pol TAG sn_pol
 INDEX ON UPPER(PADR(ALLTRIM(fam)+' '+SUBSTR(im,1,1)+SUBSTR(ot,1,1),26))+DTOC(dr) TAG fio
 INDEX on dr TAG dr
 INDEX on s_all TAG s_all
 USE 

 CREATE TABLE (Talon) ;
	(RecId i AUTOINC NEXTVALUE 1 STEP 1, ;
	 mcod c(7), period c(6), sn_pol c(25), c_i c(25), ds c(6), ds_0 c(6),  ;
	 pcod c(10), otd c(8), cod n(6), tip c(1), d_u d, ;
	 k_u n(3), d_type c(1), s_all n(11,2), profil c(3), rslt n(3), prvs n(4), ishod n(3),;
	 codnom c(14), kur n(5,3), ds_2 c(6), ds_3 c(6), det n(1), k2 n(5,3), tipgr c(1), ;
	 vnov_m n(4), novor c(9),  n_u c(14), n_vmp c(17),;
	 ord n(1), date_ord d, lpu_ord n(6), recid_lpu c(7), fil_id n(6), IsPr L, vz l, mp c(1), n_kd n(3))

 INDEX ON RecId TAG recid CANDIDATE 
 INDEX ON recid_lpu TAG recid_lpu
 INDEX ON c_i TAG c_i
 INDEX ON sn_pol TAG sn_pol
 INDEX ON otd TAG otd
 INDEX on pcod TAG pcod
 INDEX ON ds TAG ds
 INDEX ON d_u TAG d_u
 INDEX ON cod TAG cod
 INDEX ON profil TAG profil
 INDEX ON sn_pol+STR(cod,6)+DTOS(d_u) TAG ExpTag
 INDEX ON sn_pol+otd+ds+PADL(cod,6,'0')+DTOC(d_u) TAG unik 
 INDEX ON tip TAG tip
 INDEX ON s_all TAG s_all
 USE 
 
 CREATE TABLE (Otdel) ;
	(recid c(6), mcod c(7), iotd c(8), name c(100), pr_name c(100), cnt_bed n(5), fil_id n(6))
 INDEX ON iotd TAG iotd
 USE 

 CREATE TABLE (Doctor) ;
   (pcod c(10),sn_pol c(25),fam c(25),im c(20),ot c(20),dr d, w n(1),;
    prvs n(4), d_ser d, d_prik d, iotd c(8),;
	lgot_r c(1),c_ogrn c(15),lpu_id n(6), fil_id n(6))
 INDEX ON pcod TAG pcod
 USE 

 CREATE TABLE (eError) (f c(1), c_err c(3), rid i)
 INDEX FOR UPPER(f)='R' ON rid TAG rrid
 INDEX FOR UPPER(f)='S' ON rid TAG rid
 USE 

 CREATE TABLE (mError) ;
  (rid i autoinc, RecId i, cod n(6), k_u n(3), tip c(1), et c(1), ee c(1), usr c(6), d_exp d,;
   e_cod n(6), e_ku n(3), e_tip c(1), err_mee c(3), osn230 c(5), e_period c(6),  ;
   koeff n(4,2), straf n(4,2), docexp c(7), ;
   s_all n(11,2), s_1 n(11,2), s_2 n(11,2), impdata d)
 INDEX ON rid TAG rid 
 INDEX ON RecId TAG recid
 INDEX ON PADL(recid,6,'0')+et TAG id_et
 INDEX ON PADL(recid,6,'0')+et+LEFT(err_mee,2) TAG unik
 USE 
 
 =OpenFile(m.people, 'people', 'shar', 'sn_pol')
 =OpenFile(m.talon, 'talon', 'shar')
 
 =OpenFile(m.lpubase+'\NSI\sprlpuxx', 'sprlpu', 'shar', 'lpu_id')
 =OpenFile(m.pbase+'\'+m.gcperiod+'\0371001\calls', 'calls', 'shar')
 =OpenFile(m.pbase+'\'+m.gcperiod+'\0371001\e0371001', 'error', 'shar', 'rid')
 IF fso.FileExists(m.pbase+'\'+m.gcperiod+'\0371001\ans_sv4.dbf')
  =OpenFile(m.pbase+'\'+m.gcperiod+'\0371001\ans_sv4', 'answers', 'shar', 'recid')
 ENDIF 
 
 SELECT calls
 SET RELATION TO recid INTO error 
 IF USED('answers')
  SET RELATION TO PADL(recid ,6,'0') INTO answers ADDITIVE 
 ENDIF 
 SCAN 
  IF !EMPTY(error.rid)
   LOOP 
  ENDIF 
 
  m.mcod   = '0371001'
  m.period = m.gcperiod
  m.tip_p  = 1
  m.sn_pol = sn_pol
  m.enp    = m.sn_pol
  m.tipp   = tip
  m.qq     = q
  m.fam    = fam
  m.im     = im
  m.ot     = ot
  m.w      = w
  m.dr     = dr
  m.d_type = '0'
  m.sv     = ans_r
  m.recid_lpu = recid_lpu
  
  m.d_beg = d_u
  m.d_end = d_u
  m.s_all = tar

  IF USED('answers')
   m.lpu_id = answers.lpu_id
   m.prmcod = IIF(SEEK(m.lpu_id, 'sprlpu'), sprlpu.mcod, '')
  ENDIF 
  
  IF !SEEK(m.sn_pol, 'people')
   INSERT INTO people FROM MEMVAR 
  ELSE 
   m.os_all = people.s_all
   m.ns_all = m.os_all + m.s_all
   m.od_beg = people.d_beg
   m.od_end = people.d_end
   m.nd_beg = MIN(m.od_beg, m.d_beg)
   m.nd_end = MAX(m.od_end, m.d_end)
   UPDATE people SET s_all=m.ns_all, d_beg=m.nd_beg, d_end=m.nd_end WHERE sn_pol=m.sn_pol
  ENDIF 
  
  m.c_i    = ALLTRIM(pst)+'#'+ALLTRIM(c_br)+'#'+ALLTRIM(n_u)
  m.ds     = ds
  m.pcod   = c_br
  m.cod    = cod
  m.d_u    = d_u
  m.k_u    = 1
  m.profil = profbr
  m.fil_id = sp_id
  m.rslt   = res
  m.otd    = pst
  
  INSERT INTO talon FROM MEMVAR 
  
 ENDSCAN 
 SET RELATION OFF INTO error 
 IF USED('answers')
  SET RELATION OFF INTO answers
  USE IN answers
 ENDIF 
 USE IN calls
 
 USE IN people
 USE IN talon 
 
 MESSAGEBOX('OK!',0+64,'')
 
RETURN 

 

 
 IF 3=2

 List = 1
 Acti Wind Oms2

 If stc_d>0
  Stroka1 = '��������� ����������� �����������  : ' + Allt(qnam)
  Stroka2 = '�������-���������������� ����������: ' + Allt(lp(num_lpu,2))
  Stroka3 = '����-������� N '+Allt(Str(NMonth))+' �� '+DToC(date())+iif(InfTip=1,' (��������)','( ��������������)')+' �� '+Allt(RMon(NMonth,1))+' '+Year+' �.'+' (c '+DToC(Dat1)+ ' �� '+DToc(Dat2)+')'
  Stroka4 = '�� ��������� ����������� ������ �� ��������� ��� '+stroka6
  
  \\&l1O
  \<<PadR(Stroka1,112)>>
  \<<PadR(Stroka2,112)>>
  \
  \<<PadC(Stroka3,112)>>
  \<<PadC(Stroka4,112)>>
  \
   
  For iii = 1 To stc_d
   inOtd = STC(iii,13)
   ggg = ASca(STE,Stc(iii,13))
   \���������: <<iif(Seek(Stc(iii,13),'Otdel'), Allt(Otdel.IOtd)+' ('+Allt(Otdel.Otd)+'), '+Allt(Otdel.Name), Space(40))>>
   \���������� �����-���� �����   : <<iif(ggg>0, STE(ggg+1), 0)>>
   \���������� ����������� �������: <<iif(ggg>0, STE(ggg+2)+STE(ggg+3), 0)>>, � �.�.: �� "������������" ������ : <<iif(ggg>0, STE(ggg+2), 0)>>, �� "�����������" ������ : <<iif(ggg>0, STE(ggg+3), 0)>>
   \+--------------------------------------------------------------------------------------------------------------+
   \� ���  ���� ��   ������������ �����������     �  "�����������" ������  �"����������"������ �������� �����  �� �
   \� ���  ����-X �                                �------------------------+--------------------------�  �������  �
   \�      �      �                                ����.�  �����  �  �����  ����.� ��� ����� �  �����  �   �����   �
   \�------+------+--------------------------------+----+---------+---------+----+-----+-----+---------+-----------�
   Dime STD(7)
   STD = 0
   Do While STC(iii,13) = inOtd And iii<=stc_d
    std(1) = Std(1) + stc(iii,5)
    std(2) = Std(2) + stc(iii,7)
    std(3) = Std(3) + stc(iii,8)
    std(4) = Std(4) + stc(iii,9)
    std(5) = Std(5) + stc(iii,10)
    std(6) = Std(6) + stc(iii,11)
    std(7) = Std(7) + stc(iii,12)
    \�<<Stc(iii,2)>>�<<Stc(iii,3)>>�<<Stc(iii,4)>>�<<PadL(Stc(iii,5),4))>>�<<Trans(Stc(iii,6),'999999.99')>>�
    \\<<Trans(Stc(iii,7),'999999.99')>>�<<PadL(Stc(iii,8),4)>>�<<PadL(Stc(iii,9),5)>>�<<PadL(Stc(iii,10),5)>>�
    \\<<Trans(Stc(iii,11),'999999.99')>>�<<Trans(Stc(iii,12),'99999999.99')>>�
    iii = iii+1
    if iii > stc_d
     Exit
    EndI
   EndD
   iii = iii-1
   \�----------------------------------------------+----+---------+---------+----+-----+-----+---------+-----------�
   \� �����:                                       �<<PadL(Std(1),4))>>�         �<<Trans(Std(2),'999999.99')>>�<<PadL(Std(3),4)>>
   \\�<<PadL(Std(4),5)>>�<<PadL(Std(5),5)>>�<<Trans(Std(6),'999999.99')>>
   \\�<<Trans(Std(7),'99999999.99')>>�
   \+--------------------------------------------------------------------------------------------------------------+
   if iii < stc_d
*   \
   Else
   \����� �� ���������� ����������:
   \ ���������� �����-���� �����         :
   \ ���������� ����������� ������� �����: <<k_obrs>>
   \  � ��� �����:
   \   �� "������������" ������ �������  : <<k_z>>
   \   �� "�����������" ������ �������   : <<k_nz>>
   EndI
  EndF

  If sta_d>0
  \
  \&l0O
  For iii = 1 To sta_d
   i=1
   \                                                                # <<alltrim(str(list))>>
   If list=1
    \                                                                �.N ��� - 2
    \+--------------------------------------------------------------------------+
    \������������������� ��������� ����������� ������  ��� � ����� � C����.   �
    \�------+-----------------------------------------+------+-------+----------�
    \�  1   �                    2                    �   3  �   4   �    5     �
    \�------+-----------------------------------------+------+-------+----------�
   Else
    \                                                                 �.N ��� - 2
    \+--------------------------------------------------------------------------+
    \�  1   �                   2                     �  3   �   4   �    5     �
    \�------+-----------------------------------------+------+-------+----------�
   EndI

   Do Whil i <= iif(list=1, PageLen-8, PageLen) And iii <= sta_d
    x = Len(Allt(Sta(iii,5)))
    If x <= 40
     \�<<Sta(iii,1)>>�<<PadL(Sta(iii,5),41)>>�<<PadL(Sta(iii,2),6)>>�<<Tran(Sta(iii,3), '9999.99')>>�<<Tran(Sta(iii,4), '9999999.99')>>�
     i=i+1
    Else
     \�<<Sta(iii,1)>>�<<PadL(Sta(iii,5),40)>>-�<<PadL(Sta(iii,2),6)>>�<<Tran(Sta(iii,3), '9999.99')>>�<<Tran(Sta(iii,4),'9999999.99')>>�
     \�<<Repl(' ',6)>>�<<PadR(Subs(Allt(Sta(iii,5)),41,40),41)>>�<<RepL(' ',6)>>�<<RepL(' ',7)>>�<<RepL(' ',10)>>�
     i=i+2
    EndI
    iii = iii + 1
    was_time=seco()
    Do Whil Seco()-was_time <= 0.05
    EndD
   EndD
   iii = iii-1

   If iii < sta_d
    \+--------------------------------------------------------------------------+
    \
   Else
    \�------+-----------------------------------------+------+-------+----------�
    \������:�                                         �<<PadL(Allt(Str(TotalUsl)),6)>>�   �   �<<Tran(TotalSumP, '9999999.99'))>>�
    \+--------------------------------------------------------------------------+
   EndI
   list=list+1
  EndF
  EndI
  TotalSum = TotalSumS + TotalSumP
  \
  \����� ������������ � ������: <<trans(totalsum,'99999999.99')>> ���.
  \<<cpr(INT(totalsum))+str(int((totalsum-int(totalsum))*100),2)+' ���.'>>
  \
  \��������� ��������� ����� ���������� ��� ���������� � ������ ���:
  \... <<trans(ZSum,'99999999.99')>> ���.
  \<<cpr(INT(ZSum))+str(int((ZSum-int(ZSum))*100),2)+' ���.'>>
  \
  \ �� ������������� <<allt(lp(num_lpu,2))>>
  \  ������������ �������������                  <<allt(lp(num_lpu,3))>>
  \
  \  ������� ���������                           <<allt(lp(num_lpu,4))>>
  \
  \&l0O
  
 Else
 For iii = 1 To sta_d
  i=1
  \                                                                # <<alltrim(str(list))>>
  If list=1
   \                                                                �.N ��� - 2
   \ ��������� ����������� <<qnam>>
   \ ����������� ���������� : <<allt(lp(num_lpu,2))>>
   \
   \                 ����-������� N <<NMonth>> �� <<date()>> <<iif(InfTip=1,'(��������)','(��������������)')>>
   \                             �� <<allt(RMon(NMonth,1))>> <<Year>> �.
   \                 �� ��������� ����������� ������ �� ��������� ���
   \ ���������� ������������ ��������� : <<kol_obr>>
   \
   \+--------------------------------------------------------------------------+
   \������������������� ��������� ����������� ������  ��� � ����� � C����.   �
   \�------+-----------------------------------------+------+-------+----------�
   \�  1   �                    2                    �   3  �   4   �    5     �
   \�------+-----------------------------------------+------+-------+----------�
  Else
   \                                                                 �.N ��� - 2
   \ ��������� ����������� : <<qnam>>
   \ ����������� ����������: <<allt(lp(num_lpu,2))>>
   \
   \+--------------------------------------------------------------------------+
   \�  1   �                   2                     �  3   �   4   �    5     �
   \�------+-----------------------------------------+------+-------+----------�
  EndI

  Do Whil i <= iif(list=1, PageLen-8, PageLen) And iii <= sta_d
   x = Len(Allt(Sta(iii,5)))
   If x <= 40
    \�<<Sta(iii,1)>>�<<PadL(Sta(iii,5),41)>>�<<PadL(Sta(iii,2),6)>>�<<Tran(Sta(iii,3), '9999.99')>>�<<Tran(Sta(iii,4), '9999999.99')>>�
    i=i+1
   Else
    \�<<Sta(iii,1)>>�<<PadL(Sta(iii,5),40)>>-�<<PadL(Sta(iii,2),6)>>�<<Tran(Sta(iii,3), '9999.99')>>�<<Tran(Sta(iii,4),'9999999.99')>>�
    \�<<Repl(' ',6)>>�<<PadR(Subs(Allt(Sta(iii,5)),41,40),41)>>�<<RepL(' ',6)>>�<<RepL(' ',7)>>�<<RepL(' ',10)>>�
    i=i+2
   EndI
   iii = iii + 1
   was_time=seco()
   Do Whil Seco()-was_time <= 0.05
   EndD
  EndD
  iii = iii-1

  if iii < sta_d
   \+--------------------------------------------------------------------------+
  else
   \�------+-----------------------------------------+------+-------+----------�
   \������:�                                         �<<PadL(Allt(Str(TotalUsl)),6)>>�   �   �<<Tran(TotalSumP, '9999999.99'))>>�
   \+--------------------------------------------------------------------------+
   \
   \ ����� ������������ � ������: <<trans(TotalSumP,'99999999.99')>> ���.
   \   <<cpr(INT(TotalSumP))+str(int((TotalSumP-int(TotalSumP))*100),2)+' ���.'>>
   \
   \ �� ������������� <<allt(lp(num_lpu,2))>>
   \  <<PadR(Allt(lp(num_lpu,32)),40)>>    <<allt(lp(num_lpu,3))>>
   \
   \  <<PadR(Allt(lp(num_lpu,33)),40)>>    <<allt(lp(num_lpu,4))>>
   \
  EndI
  \
  list=list+1
 EndF
 EndI

  Deac Wind Oms2
  Do BeepOk
  set text to
  set text noshow
  TxFile=iif(Polz=1, '&PLocal\oms2.txt', '&PBase\&lp(num_lpu,1)\Oms2.txt')

  If Kol_Lpu = 1
   copy file &TxFile to &PLocal\temp.txt
   on key label Ctrl+P copy file &PLocal\temp.txt to prn
   Keyb [{Ctrl+End}]
   modi comm &TxFile wind Oms2 noed
   on key label Ctrl+P
   dele file &PLocal\temp.txt
  EndI

 Use In Otdel

 Pop Key
*lnSumtarz = 0
*SUM s_all TO lnSumtarz FOR !INLIST(LOWER(d_type),'z','h')
*IF lnSumtarz > 0
*	lcDocZ = '��������� ��������� �����, ���������� ��� ���������� � ������ ��� ' +;
*		ALLTRIM(STR(INT(lnSumtarz))) + ' ���. ' + ALLTRIM(STR(mod(lnSumtarz,1)*100)) + ' ���.'
*ELSE
*	lcDocZ = ''
*endif

*SELECT * FROM cSfile WHERE !INLIST(LOWER(d_type),'z','h') INTO CURSOR cSfilez


*if lcProf = 'p'
*	do case
*		case oSettings.PrintDevice
*			report form sfp.frx preview
*			REPORT FORMAT sfp.frx TO file d:\sfp.txt ascii
*		otherwise
*			PRINTJOB
*				REPORT FORMAT sfp.frx TO PRINT NOCONS
*			ENDPRINTJOB
*	endcase

*else
*	if lcProf = 's'
*		lnNumPage = 0
*		if reccount() > 0
*			do case
*				case oSettings.PrintDevice
*					report form sfs1.frx preview
*				otherwise
*					PRINTJOB
*						REPORT FORMAT sfs1.frx TO PRINT NOCONS
*					ENDPRINTJOB
*			endcase
*			lnNumPage = lnNumPage + _pageno
*		endif
*		use in cSFprof

*		select cSFpol
*		if reccount() > 0
*			do case
*				case oSettings.PrintDevice
*					report form sfs2.frx preview
*				otherwise
*					PRINTJOB
*						REPORT FORMAT sfs2.frx TO PRINT NOCONS
*					ENDPRINTJOB
*			endcase
*			lnNumPage = lnNumPage + _pageno
*		endif
*		use in cSFpol

*		select cSFsvod
*		if reccount() > 0
*			do case
*				case oSettings.PrintDevice
*					report form sfs3.frx preview
*				otherwise
*					PRINTJOB
*						REPORT FORMAT sfs3.frx TO PRINT NOCONS
*					ENDPRINTJOB
*			endcase
*		endif
*		use in cSFsvod
*	else
*		messagebox('���� ����� ����� ���',0,'')	
*	endif
*endif

*loTherm.release

ENDIF 
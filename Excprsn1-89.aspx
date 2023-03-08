<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Web.Mail" %>
<HTML lang="zh-TW">
<HEAD>
   <meta http-equiv="content-type" content="text/html;charset=UTF-8" >
   <!-- <TITLE>¥Ñ¹A¾ä(>°ê¾ä¬d¤Ñ¤z¦a¤ä<)Â²©ö¤K¦r½L</TITLE>-->

</HEAD>
<BODY BGCOLOR=#FFFFFF>
<H2  ALIGN=CENTER>¥Ñ¹A¾ä(>°ê¾ä¬d¤Ñ¤z¦a¤ä<)Â²©ö¤K¦r½L></H2>
<%

    Dim SSEX, SSPNTRM, SSEXY
    Dim YYNUM, MMNUM, DDNUM, HHNUM
    Dim NYNUM, NMNUM, NDNUM, NHNUM
    Dim LLYR, LLDY, LLJG, HLLYR, LLeap, TS3, DG4 As String
    Dim i, k, J
    Dim TK(4), DG(4), TS(4), DS(4), TF(4), DF(4)
 '' Dim validateRequest=false 
 '' SSPNTRM=Request(SPNTRM)
  SSex = Request("Sex")
 YYNUM=Request("YNUM")
 MMNUM=Request("MNUM")
 DDNUM=Request("DNUM")
 HHNUM=Request("HNUM")
    HLLYR = Request("HLYR")
    LLeap = Request("LLeap")
    ''LLYR = Request("LYR")
    LLDY = Request("LDY")
    LLJG = Request("LDJG")
    'YKK=Request("YKK")
    'YKG=Request("YKG")
    Dim HRTPZ =Right(Trim(HLLYR), 6)       ''Trim() ¶·¥h"r/n/"
    Response.Write("´úinpide¶µ" & HRTPZ & "</br>")
    Response.Write(HLLYR)
   
    ''On Error Resume Next
    '' Dim AVTK() = {"¥Ò", "¬Ñ ", "¤Ð", "¨¯ ", "©° ", "¤v ", "¥³ ", "¤B ", "¤þ ", "¤A ", "¥Ò"}
    '' Dim AVDG() = {"¤l", "¥è ", "¦¦ ", "¨» ", "¥Ó ", "¥¼ ", "¤È ", "¤x ", "¨° ", "¥f ", "±G ", "¤¡ ", "¤l"}
    '' Dim AVTK = StrReverse("¤l, ¤¡, ±G, ¥f, ¨°, ¤x, ¤È, ¥¼, ¥Ó, ¨», ¦¦, ¥è,¥Ò, ¤A, ¤þ, ¤B, ¥³, ¤v, ©°, ¨¯,¤Ð, ¬Ñ")
    '''    Response.Write(AVTK)
    For i = 1 To 3
       '' TK(i) = Mid(LLYR, 2 * i - 1, 1)
        TK(i) = Mid(HRTPZ, 2 * i - 1, 1)
        DG(i) = Mid(HRTPZ, 2 * i, 1)
    Next
   
    TS3 = "©R¥D"
    DG4 = HHNUM
    TK(4) = TKFOR(TK(3), DG4)
    DG(4) = HHNUM
    '' Response.Write(TK(4))
    ''Response.Write((-21 mod 10))
    For k = 1 To 4
        TF(k) = TDFIV(TK(k))
        DF(k) = TDFIV(DG(k))
    Next
    For J = 1 To 4
        TS(J) = TDSTR(TF(3), TF(J))
        DS(J) = TDSTR(TF(3), DF(J))
    Next
     
    Dim Row0 = "<CENTER><TABLE Border=1 BGCOLOR=#FFFF00 >" & "<TR>" & "<TD width=30>¦~" & "</TD>" & "<TD width=30>¤ë</TD>" & "<TD width =30>¤é</TD>" & "<TD width =30>®É</TD>" & "<TD width =30>¬y¦~</TD>" & "<TD width =30>¤j¹B</TD>" & "</TR>"
    Dim Row1 = "<TR>" & "<TD width=30>" & TS(1) & "</TD>" & "<TD width=30>" & TS(2) & "</TD>" & "<TD width =30>" & TS3 & "<TD width =30>" & TS(4) & "</TD>" & "<TD width =30></TD>" & "<TD width =30></TD>" & "</TR>"
    Dim Row2 = "<TR>" & "<TD width=30>" & TF(1) & "</TD>" & "<TD width=30>" & TF(2) & "</TD>" & "<TD width =30>" & TF(3) & "<TD width =30>" & TF(4) & "</TD>" & "<TD width =30></TD>" & "<TD width =30></TD>" & "</TR>"
    Dim Row3 = "<TR>" & "<TD width=30>" & TK(1) & "</TD>" & "<TD width=30>" & TK(2) & "</TD>" & "<TD width =30>" & TK(3) & "<TD width =30>" & TK(4) & "</TD>" & "<TD width =30></TD>" & "<TD width =30></TD>" & "</TR>"
    Dim Row4 = "<TR>" & "<TD width=30>" & DG(1) & "</TD>" & "<TD width=30>" & DG(2) & "</TD>" & "<TD width =30>" & DG(3) & "<TD width =30>" & DG(4) & "</TD>" & "<TD width =30></TD>" & "<TD width =30></TD>" & "</TR>"
    Dim Row5 = "<TR>" & "<TD width=30>" & DF(1) & "</TD>" & "<TD width=30>" & DF(2) & "</TD>" & "<TD width =30>" & DF(3) & "<TD width =30>" & DF(4) & "</TD>" & "<TD width =30></TD>" & "<TD width =30></TD>" & "</TR>"
    Dim Row6 = "<TR>" & "<TD width=30>" & DS(1) & "</TD>" & "<TD width=30>" & DS(2) & "</TD>" & "<TD width =30>" & DS(3) & "<TD width =30>" & DS(4) & "</TD>" & "<TD width =30></TD>" & "<TD width =30></TD>" & "</TR>"
    Dim Row8 = "<TR style='writing-mode:bt-rl' width =150>" & "´L­«µÛ§@Åv:<u>§K¶O¤À¨É</u>" & "</TR>" & "</TABLE>" & "</CENTER>"
    '' Dim Row9 = "<CENTER><H5><FONT color=#FF0000>" & "­Y»Ý­n<u>¨ö ½L µû ½× ªí</u>" & "</FONT><HR></H5>"
    Dim Row10 = "<CENTER><H5><FONT color=#FF0000>" & "­Y»Ý­n<u>¨ö ½L µû ½× ªí</u> ½Ð Email:  tech.t1206@gmail.com §Y¥iÁpµ¸,  ÁÂÁÂ" & "</FONT><HR></H5>"
    Dim RowT = Row0 + Row1 + Row2 + Row3 + Row4 + Row5 + Row6 + Row8 + Row10
    Response.Write(RowT)
    '' Response.Write (Row8)
    SSEXY = Trim(((YYNUM - 1911) Mod 2)) + SSEX
    ''Response.Write(SSEXY + LLDY)
    Dim LLKT = SRCHLGNO(TK(2), DG(2), SSEXY, LLDY, TF(3))
    Dim LYRT = SRCHEGNO(TK(1), DG(1))
    ''.Response.Write(LLKT)
   '''''''''''''''''''''' '''''''''
    %>
  <script  language="VB" runat="server"> 
    Function SRCHLGNO(TK2, DG2, SEXY, SDAY, TF3)
        Dim ASTK() = {"¬Ñ", "¥Ò", "¤A", "¤þ", "¤B", "¥³", "¤v", "©°", "¨¯", "¤Ð", "¬Ñ"}
        Dim ASDG() = {"¥è", "¤l", "¤¡", "±G", "¥f", "¨°", "¤x", "¤È", "¥¼", "¥Ó", "¨»", "¦¦", "¥è"}
        Dim AVTK() = {"¥Ò", "¬Ñ", "¤Ð", "¨¯", "©°", "¤v", "¥³", "¤B", "¤þ", "¤A", "¥Ò"}
        Dim AVDG() = {"¤l", "¥è", "¦¦", "¨»", "¥Ó", "¥¼", "¤È", "¤x", "¨°", "¥f", "±G", "¤¡", "¤l"}
   
        Dim i, j, k, m, n, xo, yo, ko, ko1, ko2 As Integer
        '' Dim LK(10), LLUK(10), LR(60), LYER(60)
        Dim LK(10), LLUK(10), AVTS(10), AVDS(10), AVTF(10), AVDF(10)
        Dim TSEXY = SEXY
        Dim ASVTK(), ASVDG()
        Dim TDAYS = Split(SDAY, ":")
        Dim TDAYS1 = CInt(TDAYS(0))
        Dim TDAYS2 = CInt(TDAYS(1))
        If TSEXY = "1¨k" Or TSEXY = "0¤k" Then
            '' ko = Fix(TDAYS2 / 3) + 1
            ko2 = (TDAYS2 / 3)
            If (ko2 - Fix(ko2)) <> 0 Then
                ko = Fix(ko2) + 1
            Else
                ko = Fix(ko2)
            End If
            ASVTK = ASTK
            ASVDG = ASDG
        Else
            '' ko = Fix(TDAYS1 / 3) + 1
            ko1 = (TDAYS1 / 3)
            If (ko1 - Fix(ko1)) <> 0 Then
                ko = Fix(ko1) + 1
            Else
                ko = Fix(ko1)
            End If
            ASVTK = AVTK
            ASVDG = AVDG
          
        End If
        ''  ko = 3
        For i = 1 To 10
            If ASVTK(i) = TK2 Then
                xo = i
                Exit For
            End If
        Next
        For j = 1 To 12
            If ASVDG(j) = DG2 Then
                yo = j
                Exit For
            End If
        Next
        For k = 1 To 10
            LK(k) = ko + (k - 1) * 10
            '' LLUK(k) = Trim(LK(k)) + ASVTK((xo + k) Mod 10) + ASVDG((yo + k) Mod 12)
            AVTF(k) = TDFIV(ASVTK((xo + k) Mod 10))
            AVDF(k) = TDFIV(ASVDG((yo + k) Mod 12))
        
            AVTS(k) = TDSTR(TF3, AVTF(k))
            AVDS(k) = TDSTR(TF3, AVDF(k))
            LLUK(k) = Trim(LK(k))+ AVTS(k) +"["+ ASVTK((xo + k) Mod 10) + ASVDG((yo + k) Mod 12)+"]" + AVDS(k)
        Next
        ' Dim Rw2
        '  For m = 1 To 2
        '' Dim Rw0 ="<CENTER><TABLE Border=1 BGCOLOR=#FFFF00 >"&"<TR>"&"<TD width=30>¤U(¤º)¤W(¥~)¨ö¤T¤ø¼Æ"&"</TD>"&"<TD width =30>ÅÜ¤ø</TD>"&"<TD width =30>¦¨¨ö¦W</TD>"&"<TD width =30>¤§¨ö¦W</TD>"&"</TR>"
        Dim Rw11 = "<CENTER><TABLE Border=1 >" & "<TR >" & "<TD>¤j¹B°Ï:</TD>" & "<TD>" & LLUK(1) & "</TD>" & "<TD >" & LLUK(2) & "</TD>" & "<TD >" & LLUK(3) & "</TD>" & "<TD>" & LLUK(4) & "</TD>" & "<TD >" & LLUK(5) & "</TD>"
        '' Dim Rw2 = "<TR >" & "<TD>" & LLUK(6) & "</TD>" & "<TD >" & LLUK(7) & "</TD>" & "<TD >" & LLUK(8) & "</TD>" & "<TD>" & LLUK(9) & "</TD>" & "<TD >" & LLUK(10) & "</TD>" & "</TR>"
        Dim Rw12 = "<TD>" & LLUK(6) & "</TD>" & "<TD >" & LLUK(7) & "</TD>" & "<TD >" & LLUK(8) & "</TD>" & "<TD>" & LLUK(9) & "</TD>" & "<TD >" & LLUK(10) & "</TD>" & "</TR>"
          
        Dim Rw13 = Rw11 + Rw12 + "</TABLE>" + "</CENTER>"
        Response.Write(Rw13)
        
    End Function
    Function SRCHEGNO(TK1, DG1)
        Dim ATK() = {"¬Ñ", "¥Ò", "¤A", "¤þ", "¤B", "¥³", "¤v", "©°", "¨¯", "¤Ð", "¬Ñ"}
        Dim ADG() = {"¥è", "¤l", "¤¡", "±G", "¥f", "¨°", "¤x", "¤È", "¥¼", "¥Ó", "¨»", "¦¦", "¥è"}
        Dim i, j, k, m, n, xo, yo, ko As Integer
        Dim LR(60), LYER(60)
        ko = 3
        For i = 1 To 10
            If ATK(i) = TK1 Then
                xo = i
                Exit For
            End If
        Next
        For j = 1 To 12
            If ADG(j) = DG1 Then
                yo = j
                Exit For
            End If
        Next
        
        '' Dim Rw2
        For k = 1 To 60
            '' LR(k) = ko + (k - 1) * 10
            LYER(k) = Trim(k) + ATK((xo - 1 + k) Mod 10) + ADG((yo - 1 + k) Mod 12)
            '' Dim Rw1 = "<TD>" & LYER(k) & "</TD>"
            ''  Rw2 = Rw2 + Rw1
        Next
        Dim Rww1, Rww2, Rww3, Rww4, Rww5, Rww6, Rww7
        For m = 1 To 10
            ''LYER(m) = Trim(m) + ATK((xo + m) Mod 10) + ADG((yo + m) Mod 12)
            Rww1 = Rww1 + "<TD>" & LYER(m) & "</TD>"
            Rww2 = Rww2 + "<TD>" & LYER(10 + m) & "</TD>"
            Rww3 = Rww3 + "<TD>" & LYER(20 + m) & "</TD>"
            Rww4 = Rww4 + "<TD>" & LYER(30 + m) & "</TD>"
            Rww5 = Rww5 + "<TD>" & LYER(40 + m) & "</TD>"
            Rww6 = Rww6 + "<TD>" & LYER(50 + m) & "</TD>"
        Next
        '' Dim Rw0 ="<CENTER><TABLE Border=1 BGCOLOR=#FFFF00 >"&"<TR>"&"<TD width=30>¤U(¤º)¤W(¥~)¨ö¤T¤ø¼Æ"&"</TD>"&"<TD width =30>ÅÜ¤ø</TD>"&"<TD width =30>¦¨¨ö¦W</TD>"&"<TD width =30>¤§¨ö¦W</TD>"&"</TR>"
        ''  Dim Rw11 = "<CENTER><TABLE Border=1 >" & "<TR >" & "<TD>¬y¦~°Ï:</TD>" & "<TD>" & LYER(1) & "</TD>" & "<TD >" & LYER(2) & "</TD>" & "<TD >" & LYER(3) & "</TD>" & "<TD>" & LYER(4) & "</TD>" & "<TD >" & LYER(5) & "</TD>"
        '' Dim Rw2 = "<TR >" & "<TD>" & LYER(6) & "</TD>" & "<TD >" & LYER(7) & "</TD>" & "<TD >" & LYER(8) & "</TD>" & "<TD>" & LYER(9) & "</TD>" & "<TD >" & LYER(10) & "</TD>" & "</TR>"
        ''    Dim Rw12 = "<TD>" & LYER(6) & "</TD>" & "<TD >" & LYER(7) & "</TD>" & "<TD >" & LYER(8) & "</TD>" & "<TD>" & LYER(9) & "</TD>" & "<TD >" & LYER(10) & "</TD>" & "</TR>"
        ''   Dim Rw13 = Rw11 + Rw12 + "</TABLE>" + "</CENTER>"
        ''  Response.Write(Rw13)
        '' Dim Rw13 = Rw11 + Rw12 + "</TABLE>" + "</CENTER>"
        ''  Response.Write(Rw13)
        Dim Rw0 = "<CENTER><TABLE Border=1 >" & "<TR >¤»¤Q¦~¥Ò¤l¬y¦~¹B°Ï" & "</TR>"
        ''Dim Rw1 = "<TR>" & "<TD width=30>" & ( Join(Selected)) & "</TD>"& "<TD width =30>" & Selectd6 & "</TD>"& "<TD width =30>" & FGACK & "</TD>"& "<TD width =30>" & FGCCK & "</TD>"& "</TR>"&"</TABLE>"
        ''Dim Rw2 = "<TR>"&"<TD width=76 >" & FGACK & "</TD>"&"<TD>" & "</TD>" & "<TD width=30  >"& (MDN33+1) & "</TD>"&"<TD width=76>"& FGCCK &"</TD>"&"</TR>"&"</TABLE>"
        Dim RwT = Rw0 + "<TR >" + Rww1 + "/<TR >" + "<TR >" + Rww2 + "</TR >" + "<TR >" + Rww3 + "</TR >" + "<TR >" + Rww4 + "</TR >"+  "<TR >" + Rww5 + "</TR >" +"<TR >" + Rww6 + "</TR >" + "</TABLE>"
        Response.Write(RwT)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
    End Function
     Function TKFOR(DTK3, HDG4)
        Dim ASTK() = {"¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ"}
        Dim ASDG() = {"¥è","¤l","¤¡","±G","¥f","¨°","¤x","¤È","¥¼","¥Ó","¨»","¦¦","¥è"}
     ' '  Dim AVTK() = {"¥Ò","¬Ñ","¤Ð","¨¯","©°","¤v","¥³","¤B","¤þ","¤A","¥Ò"}
     ' '   Dim AVDG() = {"¤l","¥è","¦¦","¨»","¥Ó","¥¼","¤È","¤x","¨°","¥f","±G","¤¡","¤l"}
     ' ' Dim i, j, k, m, n, xo, yo, ko, ko1, ko2 As Integer
       ' ' Dim LK(12), LLUK(12), LR(60), LYER(60)
     ' ' Dim NAMT0(12),ZNAMT1(12),ZNAMT2(12),ZNAMT3(12),ZNAMT4(12) NAMT5(12),ZNAMT6(12),ZNAMT7(12),ZNAMT8(12),ZNAMT9(12),ZNAMT10(10)

Dim ZNAMTT(10,12)
Dim SKCOL()= {"0","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ"}
Dim SKROW()= {"0","¤l","¤¡","±G","¥f","¨°","¤x","¤È","¥¼","¥Ó","¨»","¦¦","¥è"}
  ' 'Dim  ZNAMT0()= {"0","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ"}
Dim  ZNAMT1()= {"0","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ","¥Ò","¤A"}
Dim  ZNAMT2()= {"0","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B"}
Dim  ZNAMT3()= {"0","¥³","¤v","©°","¨¯","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v"}
Dim  ZNAMT4()= {"0","©°","¨¯","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯"}
Dim  ZNAMT5()= {"0","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ"}
Dim  ZNAMT6()= {"0","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ","¥Ò","¤A"}
Dim  ZNAMT7()= {"0","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B"}
Dim  ZNAMT8()= {"0","¥³","¤v","©°","¨¯","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v"}
Dim  ZNAMT9()= {"0","©°","¨¯","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯"}
Dim  ZNAMT10()= {"0","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ"}
 ' '//Dim ZNAMTT(10,12)   // x[5][12] = 3.0;
 ' '     var ZNAMTT = new Array(8);
  Dim i, j, k, m, n As Integer
  For i = 0 To 12
       ZNAMTT(1,i) = ZNAMT1(i) 
       ZNAMTT(2,i) = ZNAMT2(i) 
       ZNAMTT(3,i) = ZNAMT3(i) 
       ZNAMTT(4,i) = ZNAMT4(i) 
       ZNAMTT(5,i) = ZNAMT5(i) 
       ZNAMTT(6,i) = ZNAMT6(i)
       ZNAMTT(7,i) = ZNAMT7(i) 
       ZNAMTT(8,i) = ZNAMT8(i) 
       ZNAMTT(9,i) = ZNAMT9(i) 
       ZNAMTT(10,i) = ZNAMT10(i) 
          
       Next
        
        For j = 0 To 10
        If SKCOL(j) = DTK3 Then
        m = j
         Exit For
          End If
        Next

        For k= 0 To 12
        If SKROW(k) = HDG4 Then
        n = k
         Exit For
          End If
        Next
     ' '' ' // return (ZNAMTT[m][n]);
      ' '   // return (m+n);
     ' '// } // End Function 
       Return (ZNAMTT(m,n))
        
End Function

Function TDSTR(FTK3,FDG4)
        Dim ASTK() = {"¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ"}
        Dim ASDG() = {"¥è","¤l","¤¡","±G","¥f","¨°","¤x","¤È","¥¼","¥Ó","¨»","¦¦","¥è"}
     ' '  Dim AVTK() = {"¥Ò","¬Ñ","¤Ð","¨¯","©°","¤v","¥³","¤B","¤þ","¤A","¥Ò"}
     ' '   Dim AVDG() = {"¤l","¥è","¦¦","¨»","¥Ó","¥¼","¤È","¤x","¨°","¥f","±G","¤¡","¤l"}
     ' ' Dim i, j, k, m, n, xo, yo, ko, ko1, ko2 As Integer
       ' ' Dim LK(12), LLUK(12), LR(60), LYER(60)
     ' ' Dim NAMT0(12),ZNAMT1(12),ZNAMT2(12),ZNAMT3(12),ZNAMT4(12) NAMT5(12),ZNAMT6(12),ZNAMT7(12),ZNAMT8(12),ZNAMT9(12),ZNAMT10(10)

Dim ZNAMTT(10,12)
Dim SKCOL()= {"0","+¤ì","-¤ì","+¤õ","-¤õ","+¤g","-¤g","+ª÷","-ª÷","+¤ô","	-¤ô"}
Dim SKROW()= {"0","+¤ì","-¤ì","+¤õ","-¤õ","+¤g","-¤g","+ª÷","-ª÷","+¤ô","	-¤ô"}
  ' 'Dim  ZNAMT10()= {"0","¶Ë©x","­¹¯«","¥¿°]","°¾°]","¥¿©x","°¾©x","¥¿¦L","°¾¦L","§T°]","¤ñªÓ"}
Dim  ZNAMT1()= {"0","¤ñªÓ","§T°]","­¹¯«","¶Ë©x","	°¾°]","¥¿°]","°¾©x","¥¿©x","°¾¦L","¥¿¦L"}
Dim  ZNAMT2()= {"0","§T°]","¤ñªÓ","¶Ë©x","­¹¯«","¥¿°]","°¾°]","¥¿©x","°¾©x","¥¿¦L","°¾¦L"}
Dim  ZNAMT3()= {"0","°¾¦L","¥¿¦L","¤ñªÓ","§T°]","­¹¯«","¶Ë©x","°¾°]","¥¿°]","°¾©x","¥¿©x"}
Dim  ZNAMT4()= {"0","¥¿¦L","°¾¦L","§T°]","¤ñªÓ","¶Ë©x","­¹¯«","¥¿°]","°¾°]","¥¿©x","°¾©x"}
Dim  ZNAMT5()= {"0","°¾©x","¥¿©x","°¾¦L","¥¿¦L","¤ñªÓ","§T°]","­¹¯«","¶Ë©x","°¾°]","¥¿°]"}
Dim  ZNAMT6()= {"0","¥¿©x","°¾©x","¥¿¦L","°¾¦L","§T°]","¤ñªÓ","¶Ë©x","­¹¯«","¥¿°]","°¾°]"}
Dim  ZNAMT7()= {"0","°¾°]","¥¿°]","°¾©x","¥¿©x","°¾¦L","¥¿¦L","¤ñªÓ","§T°]","­¹¯«","¶Ë©x"}
Dim  ZNAMT8()= {"0","¥¿°]","°¾°]","¥¿©x","°¾©x","¥¿¦L","°¾¦L","§T°]","¤ñªÓ","¶Ë©x","­¹¯«"}
Dim  ZNAMT9()= {"0","­¹¯«","¶Ë©x","°¾°]","¥¿°]","°¾©x","¥¿©x","°¾¦L","¥¿¦L","¤ñªÓ","§T°]"}
Dim  ZNAMT10()= {"0","¶Ë©x","­¹¯«","¥¿°]","°¾°]","¥¿©x","°¾©x","¥¿¦L","°¾¦L","§T°]","¤ñªÓ"}


 ' '//Dim ZNAMTT(10,12)   // x[5][12] = 3.0;
 ' '     var ZNAMTT = new Array(8);
  Dim i, j, k, m, n As Integer
  For i = 0 To 10
       ZNAMTT(1,i) = ZNAMT1(i) 
       ZNAMTT(2,i) = ZNAMT2(i) 
       ZNAMTT(3,i) = ZNAMT3(i) 
       ZNAMTT(4,i) = ZNAMT4(i) 
       ZNAMTT(5,i) = ZNAMT5(i) 
       ZNAMTT(6,i) = ZNAMT6(i)
       ZNAMTT(7,i) = ZNAMT7(i) 
       ZNAMTT(8,i) = ZNAMT8(i) 
       ZNAMTT(9,i) = ZNAMT9(i) 
       ZNAMTT(10,i) = ZNAMT10(i) 
          
       Next
        
        For j = 0 To 10        

        If Trim(SKCOL(j)) = Trim(FTK3) Then
        m = j
         Exit For
          End If
        Next

        For k= 0 To 10
        If Trim(SKROW(k)) = Trim(FDG4) Then
        n = k
         Exit For
          End If
        Next
     ' '' ' // return (ZNAMTT[m][n]);
      ' '   // return (m+n);
     ' '// } // End Function 
       Return (ZNAMTT(m,n))
        
End Function
           
  Function TDFIV(DGTK)
        Dim ASTK() = {"¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ"}
        Dim ASDG() = {"¥è","¤l","¤¡","±G","¥f","¨°","¤x","¤È","¥¼","¥Ó","¨»","¦¦","¥è"}
        Dim mad11,mad22 As String
          
      ' '  Dim AVTK() = {"¥Ò","¬Ñ","¤Ð","¨¯","©°","¤v","¥³","¤B","¤þ","¤A","¥Ò"}
     ' '   Dim AVDG() = {"¤l","¥è","¦¦","¨»","¥Ó","¥¼","¤È","¤x","¨°","¥f","±G","¤¡","¤l"}
     ' ' Dim i, j, k, m, n, xo, yo, ko, ko1, ko2 As Integer
     ' ' Dim LK(12), LLUK(12), LR(60), LYER(60)
     ' ' Dim NAMT0(12),ZNAMT1(12),ZNAMT2(12),ZNAMT3(12),ZNAMT4(12) NAMT5(12),ZNAMT6(12),ZNAMT7(12),ZNAMT8(12),ZNAMT9(12),ZNAMT10(10)
    ' ' Dim ZNAMTT(10,12)
Dim TKDG()= {"0","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°","¨¯","¤Ð","¬Ñ","¤l","¤¡","±G","¥f","¨°","¤x","¤È","¥¼","¥Ó","¨»","¦¦","¥è"}
Dim SXFIV()= {"0","+¤ì","-¤ì","+¤õ","-¤õ","+¤g","-¤g","+ª÷","-ª÷","+¤ô","	-¤ô","-¤ô","-¤g","+¤ì","-¤ì","+¤g","+¤õ","-¤õ","-¤g","+ª÷","-ª÷","+¤g","+¤ô"}
 ''Dim  ZNAMT0()= {"0","¤Ð","¬Ñ","¥Ò","¤A","¤þ","¤B","¥³","¤v","©°",
''//Dim ZNAMTT(10,12)   // x[5][12] = 3.0;
 ' //    var ZNAMTT = new Array(8);
  Dim i, j, k, m, n As Integer
  For j = 0 To 22
        If Trim(TKDG(j)) = Trim(DGTK) Then
            mad22= SXFIV(j)
            Exit For
        End If
    Next

    Return mad22
   End Function
   
  </script>
 
</BODY>
    
</HTML>
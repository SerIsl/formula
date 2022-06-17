Function Add(num1, num2)
	Add = num1+num2
End Function


Function getVal(val1)
	Dim x
	x=Split(Format.Objects(val1).Value, " ")
	getVal=CDbl(x(0))
End Function
' ==========================================
' Stringten belirli formatta sayi almak
' ==========================================

Function getNum(num1)

	Dim x, y
	x = Trim(num1)

	If InStr(1, x, "x", 1) > 0 Then
	 
		y = Left(x,InStrRev(x, " ")-1)
		getNum = CDbl(Right(y, Len(y)-InStr(1, y, "x", 1)))
	Else
		y = Left(x,InStrRev(x, " ")-1)
		getNum = Round(CDbl(y)/230,2)
	End If

End Function

Sub getName
	Dim x

	x=Split(Format.BaseName,"_")

	If Ubound(x)>4 Then 
	Value = x(1)
	fModel = x(2)
	EAN = x(Ubound(x)-1)
		If Ubound(x)=6 Then
			Code = x(Ubound(x)-2)
		End If
	ElseIf Ubound(x)=4 AND IsNumeric(x(Ubound(x)-1)) AND Len(x(Ubound(x)-1))>11 OR Ubound(x)=2 Then 
	Value = x(1)
	fModel = "Gerek Yok!"
	EAN = x(Ubound(x)-1)
	Else
	EAN = x(Ubound(x)-1)
	fModel = x(2)
	Value = x(1)
	End If
End Sub

Function Dimension(x)
	Dim a, koli, b, koli_tipi, c, d
	
	koli = Split(Format.BaseName, "_")
	
	If x = "Gerek Yok!" Then
		c = koli(0)
	Else
		c = x
	End If

	If IsNumeric(Mid(c, 3, 1)) Then
		a = Mid(c, 3, 1)
	End If
	
    d = Mid(c, 1, 1)
	 
	
	koli_tipi = Split(koli(UBound(koli)), "-")(1)

    Select Case d
        case "F"
            If koli_tipi = "SHRINK" Then
                Select Case a
                    Case 4
                        b = "540x670x880"
                    Case 5
                        b = "545x695x880"
                    case 6 
                        b = "645x695x880"
                    case 9 
                        b = "950x680x890"
                End Select 

            ElseIf koli_tipi = "ABOX" Then

                Select Case a
                    Case 4
                        b = "515x610x883"
                    Case 5
                        b = "515x670x883"
                    case 6 
                        b = "615x670x883"
                    case 9
                        b = "905*650*810"
                End Select 
            End If
        case "H"
            If koli_tipi = "MOON" Then
                Select Case a
                    case 3
                        b = "365x570x155"
                    case 6
                        b = "635x570x155"
                    case 7
                        b = "795x570x155"
                    case 9
                        b = "915x570x155"
                End Select
            ElseIf koli_tipi = "STANDART" Then
                b = "640x595x130"
            End If 
        case "B"
            b = "632x654x634"
    End Select
	

	Value = b
End Function

Sub getName1
	Dim x

	x=Split(Format.BaseName,"_")

	If Ubound(x)>3 Then 
	Value = x(0)
	fModel = x(1)
	EAN = x(Ubound(x)-1)
		If Ubound(x)=5 OR Ubound(x)=4 Then
			Code = x(Ubound(x)-2)
		End If
	ElseIf Ubound(x)=3 AND IsNumeric(x(Ubound(x)-1)) AND Len(x(Ubound(x)-1))>11 OR Ubound(x)=2 Then 
	Value = x(0)
	fModel = "Gerek Yok!"
	EAN = x(Ubound(x)-1)
	Else
	EAN = x(Ubound(x)-1)
	fModel = x(1)
	Value = x(0)
	End If

    
End Sub


'Sub Armco
'Dim x, y, z, x1,y1

'x = Format.Objects("Barkod 6").Value
'x1 = len(x)
'y = "X2"
'y1= "X2D"
'z = "TDF"

'If x1 = 25 Then 
'	Value = "GC-"&Left(x,7)&"("&Mid(x,8,2)&")"
'ElseIf InStr(x,y1)>0 AND InStr(x,z)>0 Then
'	Value = "GC-"&Left(x,10)&"("&Mid(x,InStr(x,z),3)&")"
'ElseIf InStr(x,y1)>0 Then
'	Value = "GC-"&Left(x,10)&"("&Mid(x,11,2)&")"				
'ElseIf InStr(x,y)>0 AND InStr(x,z)>0 Then
'	Value = "GC-"&Left(x,8)&"("&Mid(x,InStr(x,z),3)&")"
'ElseIf InStr(x,y)>0 Then
'	Value = "GC-"&Left(x,8)&"("&Mid(x,9,2)&")"
'Else
'	Value = "GC-"&Left(x,7)&"("&Mid(x,InStr(x,z),3)&")"
'End If

'End Sub



Sub libertyName

Dim x

x=Split(Format.BaseName,"_")

If InStr(x(0), "-")>0 Then
fModel = x(1)
Value = x(0)
EAN = x(Ubound(x)-1) 
ElseIf Ubound(x)>2 AND InStr(x(0), "-")=0 Then 
Value = StrReverse(Replace(StrReverse(x(0)), " ", "-",1,1))
fModel =x(1)
EAN = x(Ubound(x)-1)
Else
fModel ="Gerek Yok!"
Value = x(0)
End If

End Sub

Sub bossName

Dim x

x=Split(Format.BaseName,"_")

Value = Replace(x(0),"-", "/")
fModel = x(1)
EAN = x(Ubound(x)-1) 

End Sub

Sub ac
	getName1
End Sub


Function range(x , y)
Dim arr(), i
Redim arr(y-1-x)

For i = 0 to y-1-x
arr(i) = x
x = x + 1
Next
range = arr
End Function

Function LenA(x)   
        LenA = UBound(x)+1
End Function




Sub WattTopla1 (x)
    Dim  a, result, c, d

    result = 0


    
    a = Split(x, vbCr)
    'MsgBox UBound(a)

    For Each ab in a
        c = Split(ab, " ")
        'MsgBox c(1)
        If InStr(1, c(0), "x", 1) > 0 Then
            
            d = Split(c(0), "x",-1, 1)
            result = result + CInt(d(0))*CInt(d(1)) 
        Else
            result = result + CInt(c(0))
        End If
    'MsgBox c(0)
    Next

  Value = result & " W"

End Sub

Sub Amper

Value = Replace(Round(CInt(Split(TotalWatt, " ")(0))/230, 2), ",", ".") & "A"

End Sub

Function goster(dizi)
    For Each i in dizi

        MsgBox i
 
    Next
End Function


Sub WattTopla2
    Dim  a, result, c, d, MOWA

    result = 0
	

	If Len(mOvenWatt) > 0 Then 
	 a = Split(TopWatt, vbCr)
    'MsgBox UBound(a)
    b = Split(OvenWatt, vbCr)
	 MOWA = Split(mOvenWatt, vbCr)
    ReDim e(Ubound(a)+Ubound(b)+UBound(MOWA)+2)
	'For Each q in range(0, UBound(e)+1)
   '    MsgBox q& ":" & e(q) 
   'Next
	 
	For Each i in range(0, Ubound(b)+Ubound(a)+Ubound(MOWA)+3)
        If i < UBound(a)+1 Then
            e(i) = a(i)
        ElseIf i < Ubound(a)+1+Ubound(b)+1 Then
            e(i) = b(i-UBound(a)-1) 
		  Else
				 e(i) = MOWA(i-Ubound(a)-Ubound(b)-2)
        End If
		'MsgBox  i &" : " & e(i)
    Next
	
    Else
    a = Split(TopWatt, vbCr)
    'MsgBox UBound(a)
    b = Split(OvenWatt, vbCr)
    ReDim e(Ubound(a)+Ubound(b)+1)

        'MsgBox Ubound(a)


    For Each i in range(0, Ubound(b)+Ubound(a)+2)
        If i < UBound(a)+1 Then
            e(i) = a(i)
        Else
            e(i) = b(i-UBound(a)-1) 
        End If

    Next
    'For Each q in a
    '    MsgBox q
    'Next
	End If

    For Each ab in e
        c = Split(ab, " ")
        'MsgBox c(0)
        If InStr(1, c(0), "x", 1) > 0 Then
            
            d = Split(c(0), "x",-1, 1)
            result = result + CInt(d(0))*CInt(d(1)) 
        Else
            result = result + CInt(c(0))
        End If
    'MsgBox result
    Next

  Value = result & " W"

End Sub

Function kWattTopla2 (x, y, z)
    Dim  a, result, c, d, e()

    result = 0

    a = Split(x, vbCr) 
    'MsgBox UBound(a)
    b = Split(y, vbCr)
    'MsgBox UBound(b)

    If z = 1 Then
		If LenA(b) > 1 Then
            	    ReDim e(LenA(a)+LenA(b)-2)
		Else
		    ReDim e(LenA(a)+LenA(b)-1)
		End If


        For Each i in range(0, LenA(e))
            If i < LenA(a) Then
                e(i) = a(i)
            Else
                e(i) = b(i-LenA(a)) 
            End If

        Next

        'For Each q in e

        '    MsgBox q
        'Next

        For Each ab in e
            c = Split(ab, " ")
            'MsgBox c(1)
            If InStr(1, c(0), "x", 1) > 0 Then
                
                d = Split(c(0), "x",-1, 1)
                result = result + CInt(d(0))*CDbl(d(1)) 
            Else
                result = result + CDbl(c(0))
            End If
        'MsgBox c(0)
        Next

    Else
        ReDim e(Ubound(a)+Ubound(b)+1)

        'MsgBox Ubound(a)


        For Each i in range(0, Ubound(b)+Ubound(a)+2)
            If i < UBound(a)+1 Then
            e(i) = a(i)
            Else
            e(i) = b(i-UBound(a)-1) 
            End If

        Next

        'For Each q in e

        '    MsgBox q
        'Next

        For Each ab in e
            c = Split(ab, " ")
            'MsgBox c(1)
            If InStr(1, c(0), "x", 1) > 0 Then
                
                d = Split(c(0), "x",-1, 1)
                result = result + CInt(d(0))*CDbl(d(1)) 
            Else
                result = result + CDbl(c(0))
            End If
        'MsgBox c(0)
        Next
    
    End If
    Value = result

End Function

Function kWattTopla1 (x)
    Dim  a, result, c, d

    result = 0


    
    a = Split(x, vbCr)
    'MsgBox UBound(a)

    For Each ab in a
        c = Split(ab, " ")
        'MsgBox c(1)
        If InStr(1, c(0), "x", 1) > 0 Then
            
            d = Split(c(0), "x",-1, 1)
            result = result + CInt(d(0))*CDbl(d(1)) 
        Else
            result = result + CDbl(c(0))
        End If
    'MsgBox c(0)
    Next

  Value  = Round(result, 2) 

End Function

Function wattHesapla(x)



    Dim dizi1, dizi2, TotalDizi, tmpD

    Select Case x
        case 0 
	    If Len(OvenkWatt) > 0 Then
	        TotalDizi = Split(OvenkWatt, vbCr)
            Else 
            	TotalDizi = Split(TopkWatt, vbCr)
	    End If	
            wattHesapla = TotalDizi
        case 1
            dizi1 = Split(TopkWatt, vbCr) 
            dizi2 = Split(OvenkWatt, vbCr)
		If LenA(dizi2) > 1 Then
            	    ReDim TotalDizi(LenA(dizi1)+LenA(dizi2)-2)
		Else
		    ReDim TotalDizi(LenA(dizi1)+LenA(dizi2)-1)
		End If

            For Each i in range(0, LenA(TotalDizi))
                If  i < LenA(dizi1) Then
                    TotalDizi(i) = dizi1(i)
                Else
                    TotalDizi(i) = dizi2(i-lenA(dizi1))
                End If
            Next
            wattHesapla = TotalDizi
        case 2 
            dizi1 = Split(TopkWatt, vbCr) 
            dizi2 = Split(OvenkWatt, vbCr)
            ReDim TotalDizi(LenA(dizi1)+LenA(dizi2)-1)

            For Each i in range(0, LenA(TotalDizi))
                If  i < LenA(dizi1) Then
                    TotalDizi(i) = dizi1(i)
                Else
                    TotalDizi(i) = dizi2(i-lenA(dizi1))
                End If
            Next
            wattHesapla = TotalDizi

        case 3
            dizi1 = Split(TopkWatt25, vbCr) 
            dizi2 = Split(OvenkWatt, vbCr)
		If LenA(dizi2) > 1 Then
            	    ReDim TotalDizi(LenA(dizi1)+LenA(dizi2)-2)
		Else
		    ReDim TotalDizi(LenA(dizi1)+LenA(dizi2)-1)
		End If

            For Each i in range(0, LenA(TotalDizi))
                If  i < LenA(dizi1) Then
                    TotalDizi(i) = dizi1(i)
                Else
                    TotalDizi(i) = dizi2(i-lenA(dizi1))
                End If
            Next
            wattHesapla = TotalDizi    
        
        case 4 
	    If Len(OvenkWatt) > 0 Then
	        TotalDizi = Split(OvenkWatt, vbCr)
            Else 
            	TotalDizi = Split(TopkWatt25, vbCr)
	    End If	
            wattHesapla = TotalDizi

    End Select 

End Function    

Function gasCons(x, y)
Dim c, result, GasType
Dim G30_30, G20_20, G20_13, G30_30kW, a, b
G30_30 = Array(276, 247, 218, 182, 124, 65, 160, 102, 255)
G30_30kW = Array("3,80","3,40","3,00", "2,50", "1,70", "0,90", "2,20", "1,40", "3,50", "3,60", "3,10", "0,95", "2,50h")
G20_20 = Array(0.362, 0.324, 0.275, 0.234, 0.168, 0.085, 0.218, 0.140, , 0.418, , , 0.238)
G20_13 = Array(0.345, 0.309, 0.273, 0.227, 0.155, 0.082, 0.200, 0.127, , 0.327)
G20_25 = Array(0.362, 0.324, , 0.237, 0.160, , 0.209, 0.141, , , 0.290, 0.094)
G25_25 = Array(0.362, 0.324, 0.311, 0.258, 0.176, 0.093, 0.231, 0.160)
E_Bek_kW = Array("3,00","2,11", "1,33", "0,94", "2,20", "1,40", "3,50", "2,50")
E_Bek = Array(218, 153, 97, 68, 160, 102, 255, 182)
result = 0
a = wattHesapla(x)
'MsgBox LenA(a)
'goster(a)

	GasType = y

'MsgBox GasType
Select Case GasType
    case "G30 - 30 mbar"
        For Each ab in range(0, LenA(a))
            c=Split(a(ab), " ")
            If InStr(1, c(0), "x", 1) > 0 Then
                arr = Split(c(0), "x",-1, 1)
                'MsgBox arr(1)
                For Each sayi in range(0, LenA(G30_30kW))
                    If arr(1) = G30_30kW(sayi) Then
                        result = result + CInt(arr(0))*G30_30(sayi)
                        'MsgBox result
                    End If
                Next
            End If
            For Each ax in range(0, LenA(G30_30kW))
                If c(0) = G30_30kW(ax) Then
                    result = result + G30_30(ax)
                End If
            Next
        Next
    
    case "G20 - 20 mbar"
        For Each ab in range(0, LenA(a))
            c=Split(a(ab), " ")
            If InStr(1, c(0), "x", 1) > 0 Then
                arr = Split(c(0), "x",-1, 1)
                'MsgBox arr(1)
                For Each sayi in range(0, LenA(G30_30kW))
                    If arr(1) = G30_30kW(sayi) Then
                        result = result + CInt(arr(0))*G20_20(sayi)
                    End If
                Next
            End If
            For Each ax in range(0, LenA(G30_30kW))
                If ab > 3 and c(0) = "2,50" Then 
                    c(0) = c(0) + "h"
                    
                End If
                If c(0) = G30_30kW(ax) Then
                        'MsgBox  c(0)
                        result = result + G20_20(ax)
                End If
            Next
            'MsgBox result & "  " & CStr(ab) & ". değer"
        Next
    
    case "G20 - 13 mbar"
        For Each ab in range(0, LenA(a))
            c=Split(a(ab), " ")
            If InStr(1, c(0), "x", 1) > 0 Then
                arr = Split(c(0), "x",-1, 1)
                'MsgBox arr(1)
                For Each sayi in range(0, LenA(G30_30kW))
                    If arr(1) = G30_30kW(sayi) Then
                        result = result + CInt(arr(0))*G20_13(sayi)
                    End If
                Next
            End If
            For Each ax in range(0, LenA(G30_30kW))
                If c(0) = G30_30kW(ax) Then
                    result = result + G20_13(ax)
                End If
            Next
        Next

    case "G20 - 25 mbar"
        For Each ab in range(0, LenA(a))
            c=Split(a(ab), " ")
            If InStr(1, c(0), "x", 1) > 0 Then
                arr = Split(c(0), "x",-1, 1)
                'MsgBox arr(1)
                For Each sayi in range(0, LenA(G30_30kW))
                    If arr(1) = G30_30kW(sayi) Then
                        result = result + CInt(arr(0))*G20_25(sayi)
                    End If
                Next
            End If
            For Each ax in range(0, LenA(G30_30kW))
                If c(0) = G30_30kW(ax) Then
                    result = result + G20_25(ax)
                End If
            Next
        Next

    case "G25 - 25 mbar"
        For Each ab in range(0, LenA(a))
            c=Split(a(ab), " ")
            If InStr(1, c(0), "x", 1) > 0 Then
                arr = Split(c(0), "x",-1, 1)
                'MsgBox arr(1)
                For Each sayi in range(0, LenA(G30_30kW))
                    If arr(1) = G30_30kW(sayi) Then
                        result = result + CInt(arr(0))*G25_25(sayi)
                    End If
                Next
            End If
            For Each ax in range(0, LenA(G30_30kW))
                If c(0) = G30_30kW(ax) Then
                    result = result + G25_25(ax)
                End If
            Next
        Next

    case "E-Bek"
        For Each ab in range(0, LenA(a))
            c=Split(a(ab), " ")
            If InStr(1, c(0), "x", 1) > 0 Then
                arr = Split(c(0), "x",-1, 1)
                'MsgBox arr(1)
                For Each sayi in range(0, LenA(E_Bek_kW))
                    If arr(1) = E_Bek_kW(sayi) Then
                        result = result + CInt(arr(0))*E_Bek(sayi)
                    End If
                Next
            End If
            For Each ax in range(0, LenA(E_Bek_kW))
                If c(0) = E_Bek_kW(ax) Then
                    result = result + E_Bek(ax)
                End If
            Next
        Next
End Select
Value = result
End Function

Sub seriNo(SeriNoSon, adet, basamak)
Dim fark, serinoEsas, barkot
barkot = CInt(Right(barCode, 4))
fark = SeriNoSon - adet
serinoEsas = barkot + fark
Select Case basamak
case 4
	If serinoEsas < 1000 And serinoEsas >99 Then
	
	Value = "0" & serinoEsas

	ElseIf serinoEsas < 100 And serinoEsas > 9 Then
	Value = "00" & serinoEsas

	ElseIf serinoEsas < 10 Then
	Value = "000" & serinoEsas
	Else
	Value = serinoEsas
	End If

case 5
	If serinoEsas < 10000 And serinoEsas >999 Then
	
	Value = "0" & serinoEsas
   ElseIf serinoEsas < 1000 And serinoEsas >99 Then
	
	Value = "00" & serinoEsas

	ElseIf serinoEsas < 100 And serinoEsas > 9 Then
	Value = "000" & serinoEsas

	ElseIf serinoEsas < 10 Then
	Value = "0000" & serinoEsas
	Else 
	Value = serinoEsas
	End If

End Select

End Sub


Sub ColorName

Dim ColorCode, ColorName, fModel

ColorCode = Array("W", "B", "G", "K", "S", "M", "Y", "R", "C", "D", "P", "A", "V", "L", "H", "E", "U", "I", "F", "J", "N", "O", "T", "X", "x", "v", "a", "f", "h", "b", "k", "s","e", "j", "k", "g", "c", "t", "Z")

ColorName = Array("WHITE", "BLACK", "GREY (METALLIC)", "BROWN", "INOX BRIGHT", "INOX MATT", "GREEN", "RED", "BEIGE", "MATT BLACK", "PINK", "ANTRACIT", "BLUE", "WOODEN DESIGN","HALF INOX", "COOKTOP MATT BLACK", "COOKTOP BLACK", "COOKTOP MIRRORED BLACK", "BRIGHT BLACK", "LIGHT YELLOW", "YELLOW", "ORANGE", "TURQUAZ", "DARK RED", "BRIGHT RED", "OCEAN BLUE", "MATT ANTRACIT", "FULL INOX", "HALF INOX", "DARK BLUE", "BEIGE PAINT", "SILVER GREY PAINT", "COOKTOP GREY ENAMEL", "COOKTOP INOX", "COOKTOP AND COMMAND PANEL INOX", "GOLD","CUPPER", "TDF", "NO INFO")

fModel = Right(Split(Format.BaseName, "_")(1), 1) 

For Each sayi in range(0, LenA(ColorCode))

	If ColorCode(sayi) = fModel Then
		Value = ColorName(sayi)

	End If

Next

End Sub


Sub firinolcusu(modele)
Dim a, a1, a2


a = Split(modele, "(")
a1 = Replace(a(1), ")", "")
a2 = Trim(a(0))
If  Len(a1) = 10  Then
	If CInt(Mid(a1, 3, 1)) = 6 Then
		Value = "60x60"
	ElseIf CInt(Mid(a1, 3, 1)) = 5 Then
		Value = "50x60"
	Else
		Value = "50x50"
	End If
Else
	If CInt(Mid(a2, 3, 1)) = 6 Then
		Value = "60x60"
	ElseIf CInt(Mid(a2, 3, 1)) = 5 Then
		Value = "50x60"
	Else
		Value = "50x50"
	End If

End If
End Sub

Sub BarkodNo

Dim x

	x=Split(Format.BaseName,"_")
    Value = x(0)

End Sub

Sub firinolcusu2(modele)

	If CInt(Mid(modele, 3, 1)) = 6 Then
		Value = "60x60"
	ElseIf CInt(Mid(modele, 3, 1)) = 5 Then
		Value = "50x60"
	Else
		Value = "50x50"
	End If

End Sub

Sub firinolcusu3(modele)
Dim a, a1, a2


a = Split(modele, "(")
a1 = Replace(a(1), ")", "")

	If CInt(Mid(a1, 3, 1)) = 6 Then
		Value = "60x60"
	ElseIf CInt(Mid(a1, 3, 1)) = 5 Then
		Value = "50x60"
	Else
		Value = "50x50"
	End If

End Sub
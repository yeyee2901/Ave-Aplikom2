Function dgratus(angka As Double) As String
    Dim Huruf(0 To 9) As String
    Dim ax(0 To 3) As Double
    Huruf(0) = ""
    Huruf(1) = "satu "
    Huruf(2) = "dua "
    Huruf(3) = "tiga "
    Huruf(4) = "empat "
    Huruf(5) = "lima "
    Huruf(6) = "enam "
    Huruf(7) = "tujuh "
    Huruf(8) = "delapan "
    Huruf(9) = "sembilan "
    
    Temp = ""
    panjang = Len(Trim(Str(angka)))
    
    nilai = Right("000", 3 - panjang) + Trim(Str(angka))
    For y = 3 To 1 Step -1
        ax(y) = Mid(nilai, y, 1)
    Next y
    
    Select Case ax(1)
        Case Is = 1
          Temp = "seratus "
        Case Is > 1
          Temp = Huruf(Val(ax(1))) + "" + "ratus "
        Case Else
          Temp = ""
    End Select
 
    Select Case ax(2)
      Case Is = 0
          Temp = Temp + Huruf(Val(ax(3)))
      Case Is = 1
          Select Case ax(3)
            Case Is = 1
              Temp = Temp + "sebelas"
            Case Is = 0
              Temp = Temp + "sepuluh"
            Case Else
              Temp = Temp + Huruf(Val(ax(3))) + " belas"
          End Select
      Case Is > 1
          Temp = Temp + Huruf(Val(ax(2))) + "puluh"
          Temp = Temp + " " + Huruf(Val(ax(3)))
      End Select
    dgratus = Temp
End Function


Function dghuruf(angka As Double) As String
    Dim ratusan(0 To 6) As String
    Dim sebut(0 To 4) As String
    sebut(1) = " ribu "
    sebut(2) = " juta "
    sebut(3) = " milyar "
    sebut(4) = " trilyun "
    panjang = Len(Trim(Str(angka)))
    kali = Int(panjang / 3)
    If Int(panjang / 3) * 3 <> panjang Then
        kali = kali + 1
        sisa = panjang - Int(panjang / 3) * 3
        nilai = Right("000", 3 - sisa) + Trim(Str(angka))
    Else
        nilai = Trim(Str(angka))
    End If
    
    For x = 0 To kali
       ratusan(kali - x) = Mid(nilai, x * 3 + 1, 3)
    Next x
    
    For y = kali To 1 Step -1
        If y = 2 And Val(ratusan(y)) = 1 Then
            Temp = Temp + "seribu "
        Else
            If Val(ratusan(y)) = 0 Then
                Temp = Temp
            Else
                Temp = Temp + dgratus(Val(ratusan(y)))
                Temp = Temp + sebut(y - 1)
            End If
        End If
    Next y
    dghuruf = Temp
End Function

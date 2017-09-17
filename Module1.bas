Attribute VB_Name = "Module1"
Sub recherche()

'variables
    
    Dim tableau As Variant
    
    Dim tableau2 As Variant
    
    tableau2 = Array("MT", "A7", "MBG", "MY", "GR", "DC", "GL", "MZ", "MAG", "GU") '(RMT, RA7, RMBG, RMY, RGR, RDC, RGL, RMZ, RMAG, RGU)

    extrait = "value=go>go"
    
    MT = "value=mt>mt"
    
    A7 = "value=a7>a7"
    
    MBG = "value=mbg>mbg"
    
    MY = "value=my>my"
    
    GR = "value=gr>gr"
    
    DC = "value=dc>dc"
    
    GL = "value=gl>gl"
    
    MZ = "value=mz>mz"
    
    MAG = "value=mag>mag"
    
    GU = "value=gu>gu"
    
    rue = "name=clients-street" ' valeur de 25
    
    codePostal = "name=clients-zip"
    
    coupure = "siz"
    
    coupure2 = "value="
    
    dcol = Cells(1, Cells.Columns.Count).End(xlToLeft).Column

'boucle pour concatener
'-----------------------------------------------------------------

    For i = 1 To dcol
    
    hexVal = hexVal + Cells(5, i)
    
    Next

    y = Len(hexVal)
    
    
'--------------------------
'variables boucles
    x = 1
    
    a = 1
    
    u = 2
    
'trouver les strings afin d'etablir le point de rencontre avec les divers elements
    
    While a <> 0
    
' trouver l'element associe a la recherche (debut, string total, extrait)

    a = InStr(a + 1, hexVal, extrait)
    
        If a <> 0 Then
        
            RMT = InStr(a, hexVal, MT)
             
            RA7 = InStr(a, hexVal, A7)
         
            RMBG = InStr(a, hexVal, MBG)
    
            RMY = InStr(a, hexVal, MY)
            
            RGR = InStr(a, hexVal, GR)
            
            RDC = InStr(a, hexVal, DC)
            
            RGL = InStr(a, hexVal, GL)
            
            RMZ = InStr(a, hexVal, MZ)
                        
            RMAG = InStr(a, hexVal, MAG)
            
            RGU = InStr(a, hexVal, GU)
        
         
         tableau = Array(RMT, RA7, RMBG, RMY, RGR, RDC, RGL, RMZ, RMAG, RGU)
                
         minValue = tableau(0)
         
         For Each rep In tableau
         
                If rep < minValue And rep <> a And rep <> 0 Then
                
                minValue = rep
                
                End If
            
            pos = WhereInArray(tableau, minValue)
                
            r = tableau2(pos)
            
         Next
        
                If minValue <> 0 Then
                
                    RV = InStr(minValue, hexVal, rue)
                    CP = InStr(minValue, hexVal, codePostal)
' donne l'adresse
                    resultat = Mid(hexVal, RV + 25, 20)
                    resultatCp = Mid(hexVal, CP + 16, 9)
'reduite l'adresse et retire le siz
                    resultatIntermediaire = InStr(1, resultat, coupure)
                    
                    resultatIntermediaireCp = InStr(1, resultatCp, coupure2)
                    
                        If resultatIntermediaire <> 0 Then
                        
                            resultatGauche = Left(resultat, resultatIntermediaire - 1) 'elimine
                            
                        Else
                        
                            resultatGauche = resultat
                        
                        End If
                        
                        If resultatIntermediaireCp <> 0 Then
                        
                            resultatCpFinal = Right(resultatCp, resultatIntermediaireCp + 2)
                            
                        Else
                        
                            resultatCpFinal = resultatCp
                        
                        End If
                        
                        
                
                    nomDuRep = r
                    
                        
                        If Worksheets("Feuil1").Cells(u, 1) <> "" Then
                            
                            u = u + 1
                            
                            Worksheets("Feuil1").Cells(u, 1) = nomDuRep
                            
                            Worksheets("Feuil1").Cells(u, 2) = resultatGauche
                            
                            Worksheets("Feuil1").Cells(u, 3) = resultatCpFinal
                            
                            Else
                            
                            Worksheets("Feuil1").Cells(u, 1) = nomDuRep
                            
                            Worksheets("Feuil1").Cells(u, 2) = resultatGauche
                            
                            Worksheets("Feuil1").Cells(u, 3) = resultatCpFinal
                        
                    
                        End If
                    
                        
                        
                
                  End If
            

        End If
        
    Wend
    
End Sub


Function WhereInArray(arr1 As Variant, vFind As Variant) As Variant
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check where a value is in an array
Dim i As Long
For i = LBound(arr1) To UBound(arr1)
    If arr1(i) = vFind Then
        WhereInArray = i
        Exit Function
    End If
Next i
'if you get here, vFind was not in the array. Set to null
WhereInArray = Null
End Function

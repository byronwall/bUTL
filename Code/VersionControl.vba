Attribute VB_Name = "VersionControl"
Sub SaveCodeModules()

'This code Exports all VBA modules
Dim i%, sName$

With ThisWorkbook.VBProject
    For i% = .VBComponents.count To 1 Step -1
        
        sName$ = .VBComponents(i%).CodeModule.name
        .VBComponents(i%).Export ThisWorkbook.path & "\Code\" & sName$ & ".vba"

    Next i
End With

End Sub

Sub ImportCodeModules()

    With ThisWorkbook.VBProject
        For i% = .VBComponents.count To 1 Step -1
    
            ModuleName = .VBComponents(i%).CodeModule.name
            
            If .VBComponents(i%).Type <> vbext_ct_Document Then
    
                If ModuleName <> "VersionControl" Then
                    
                        .VBComponents.Remove .VBComponents(ModuleName)
                        .VBComponents.Import "C:\Users\eltron\Desktop\byron work\bUTL\Code\" & ModuleName & ".vba"
                   
                End If
            End If
        Next i
    End With

End Sub

Sub PUSH()
  Dim ws As Worksheet
    Dim i As Integer
    Dim userMail As String, phoneNumber As String, policy As String, dialPlan As String, cli As String, locationId As String
    Dim cmd As String, confirmation As Integer
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ' Définir la feuille active
    Set ws = ThisWorkbook.Sheets("SDA")
    
    ' Détecter la ligne active correctement
    On Error Resume Next
    i = Windows(ThisWorkbook.Name).ActiveCell.Row
    On Error GoTo 0
    
    If i = 0 Then
        MsgBox "Impossible de détecter la ligne active.", vbExclamation
        Exit Sub
    End If
    
    ' Forcer le recalcul des formules pour éviter un problème d'affichage des valeurs
    ws.Cells(i, 5).Calculate
    ws.Cells(i, 6).Calculate
    ws.Cells(i, 7).Calculate
    ws.Cells(i, 8).Calculate
    ws.Cells(i, 9).Calculate
    ws.Cells(i, 10).Calculate
    ws.Cells(i, 11).Calculate
    ws.Cells(i, 12).Calculate
    
    ' Vérifier si l'e-mail est renseigné (colonne 5)
    If Trim(ws.Cells(i, 5).Value) = "" Then
        MsgBox "Cette ligne ne contient pas de données valides.", vbExclamation
        Exit Sub
    End If
    
    ' Récupérer les valeurs nécessaires
    userMail = Trim(ws.Cells(i, 5).Value) ' Colonne Mail
    phoneNumber = Trim(ws.Cells(i, 7).Value) ' SDA BRUTE (colonne corrigée)
    cli = Trim(ws.Cells(i, 8).Value) ' CallingLineIdentity (colonne corrigée)
    locationId = Trim(ws.Cells(i, 9).Value) ' LocationId (colonne corrigée)
    policy = Trim(ws.Cells(i, 10).Value) ' Policy (colonne corrigée)
    dialPlan = Trim(ws.Cells(i, 11).Value) ' DialPlan (colonne corrigée)
    
    ' Debugging pour voir les valeurs dans la fenêtre immédiate
    Debug.Print "Ligne: " & i
    Debug.Print "Email: " & userMail
    Debug.Print "PhoneNumber: " & phoneNumber
    Debug.Print "LocationId: " & locationId
    Debug.Print "Policy: " & policy
    Debug.Print "DialPlan: " & dialPlan
    Debug.Print "CLI: " & cli
    
    ' Demander confirmation avant d'exécuter la commande PowerShell
    confirmation = MsgBox("Confirmez-vous l'exécution de la mise à jour avec les informations suivantes ?" & vbCrLf & _
                          "Ligne: " & i & vbCrLf & _
                          "Email: " & userMail & vbCrLf & _
                          "PhoneNumber: " & phoneNumber & vbCrLf & _
                          "LocationId: " & locationId & vbCrLf & _
                          "Policy: " & policy & vbCrLf & _
                          "DialPlan: " & dialPlan & vbCrLf & _
                          "CLI: " & cli, vbYesNo + vbQuestion, "Confirmation")
    
    If confirmation = vbNo Then Exit Sub
    
    ' Construire la commande PowerShell
    cmd = "powershell -ExecutionPolicy Bypass -Command " & _
          "Connect-MicrosoftTeams; " & _
          "Set-CsPhoneNumberAssignment -Identity '" & userMail & "' -PhoneNumber '" & phoneNumber & "' -PhoneNumberType DirectRouting -LocationId '" & locationId & "'; " & _
          "Grant-CsOnlineVoiceRoutingPolicy -Identity '" & userMail & "' -PolicyName '" & policy & "'; " & _
          "Grant-CsTenantDialPlan -Identity '" & userMail & "' -PolicyName '" & dialPlan & "'; " & _
          "Grant-CsCallingLineIdentity -Identity '" & userMail & "' -Policyname '" & cli & "'"
  Debug.Print cmd
    ' Exécuter la commande PowerShell
    wsh.Run cmd, 1, True
    'Shell cmd, vbNormalFocus
    
    ' Ajouter la date et l'heure dans la colonne M
    ws.Cells(i, 13).Value = Now
    
    MsgBox "Mise à jour de la ligne terminée !", vbInformation
End Sub



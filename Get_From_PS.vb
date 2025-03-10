Sub GET_FROM_PS()
   Dim ws As Worksheet
    Dim i As Integer
    Dim userMail As String, userPrincipalName As String, phoneNumber As String, policy As String, dialPlan As String, cli As String
    Dim cmd As String, confirmation As Integer
    Dim tempFile As String, fileNum As Integer, outputLine As String
    Dim shellOutput As Object
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
    
    ' Récupérer l'adresse e-mail
    userMail = Trim(ws.Cells(i, 5).Value)
    
    If userMail = "" Then
        MsgBox "Aucun email détecté sur cette ligne.", vbExclamation
        Exit Sub
    End If
    
    ' Demander confirmation avant d'exécuter la commande PowerShell
    confirmation = MsgBox("Confirmez-vous la récupération des informations pour l'utilisateur suivant ?" & vbCrLf & _
                          "Ligne: " & i & vbCrLf & _
                          "Email: " & userMail, vbYesNo + vbQuestion, "Confirmation")
    
    If confirmation = vbNo Then Exit Sub

    ' Créer un fichier temporaire pour stocker la sortie PowerShell
    tempFile = Environ("TEMP") & "\PS_Output.txt"
    
    ' Construire la commande PowerShell pour récupérer les informations de l'utilisateur
    cmd = "powershell -ExecutionPolicy Bypass -NoProfile -Command " & _
          """Connect-MicrosoftTeams; " & _
          "$user = Get-CsOnlineUser -Identity '" & userMail & "'; " & _
          "$output = @(); " & _
          "$output += 'userPrincipalName=' + $user.UserPrincipalName; " & _
          "$output += 'PhoneNumber=' + $user.LineURI; " & _
          "$output += 'Policy=' + $user.OnlineVoiceRoutingPolicy; " & _
          "$output += 'DialPlan=' + $user.TenantDialPlan; " & _
          "$output += 'CLI=' + $user.CallingLineIdentity; " & _
          "$output | Out-File -FilePath '" & tempFile & "';"""

    ' Exécuter la commande PowerShell
    wsh.Run cmd, 1, True
    
    ' Attendre que le fichier soit rempli
    Application.Wait (Now + TimeValue("00:00:15"))
    
    ' Vérifier si le fichier a bien été créé
    If Dir(tempFile) = "" Then
        MsgBox "Erreur lors de l'exécution de PowerShell. Vérifiez la connexion à Teams.", vbExclamation
        Exit Sub
    End If
    
    ' Lire les résultats depuis le fichier temporaire
    fileNum = FreeFile
    Open tempFile For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, outputLine
        outputLine = Replace(outputLine, "ÿþ", "") ' Supprimer les caractères indésirables
        If InStr(outputLine, "PhoneNumber=") > 0 Then phoneNumber = Replace(outputLine, "PhoneNumber=", "")
        If InStr(outputLine, "userPrincipalName=") > 0 Then userPrincipalName = Replace(outputLine, "userPrincipalName=", "")
        If InStr(outputLine, "Policy=") > 0 Then policy = Replace(outputLine, "Policy=", "")
        If InStr(outputLine, "DialPlan=") > 0 Then dialPlan = Replace(outputLine, "DialPlan=", "")
        If InStr(outputLine, "CLI=") > 0 Then cli = Replace(outputLine, "CLI=", "")
    Loop
    Close #fileNum
    
    Debug.Print userPrincipalName
    
    ' Supprimer le fichier temporaire
    Kill tempFile
    
    ' Extraire uniquement ce qu'il y a après le "+" pour PhoneNumber
    If InStr(phoneNumber, "+") > 0 Then
        phoneNumber = Mid(phoneNumber, InStr(phoneNumber, "+") + 1)
    End If
    
    phoneNumber = VBA.Trim(phoneNumber)
    userPrincipalName = VBA.Trim(userPrincipalName)
    
    ' Vérifier si des valeurs ont été récupérées
    If phoneNumber = "" And policy = "" And dialPlan = "" And cli = "" Then
        MsgBox "Aucune donnée trouvée pour cet utilisateur.", vbExclamation
        Exit Sub
    End If
    
    ' Vérifier s'il y a déjà des données et demander confirmation avant d'écraser
    If ws.Cells(i, 7).Value <> "" Or ws.Cells(i, 10).Value <> "" Or ws.Cells(i, 11).Value <> "" Or ws.Cells(i, 8).Value <> "" Then
        confirmation = MsgBox("Des données existent déjà. Voulez-vous les écraser ?" & vbCrLf & _
                              "Anciennes valeurs :" & vbCrLf & _
                              "Usermail: " & ws.Cells(i, 5).Value & vbCrLf & _
                              "PhoneNumber: " & ws.Cells(i, 7).Value & vbCrLf & _
                              "Policy: " & ws.Cells(i, 10).Value & vbCrLf & _
                              "DialPlan: " & ws.Cells(i, 11).Value & vbCrLf & _
                              "CLI: " & ws.Cells(i, 8).Value & vbCrLf & _
                              "Nouvelles valeurs :" & vbCrLf & _
                              "Usermail: " & userPrincipalName & vbCrLf & _
                              "PhoneNumber: " & phoneNumber & vbCrLf & _
                              "Policy: " & policy & vbCrLf & _
                              "DialPlan: " & dialPlan & vbCrLf & _
                              "CLI: " & cli, vbYesNo + vbQuestion, "Confirmation")
        If confirmation = vbNo Then Exit Sub
    End If
    
    ' Mise à jour des cellules avec les nouvelles valeurs
    ws.Cells(i, 7).Value = phoneNumber
    ws.Cells(i, 10).Value = policy
    ws.Cells(i, 11).Value = dialPlan
    ws.Cells(i, 8).Value = cli
    
    ' Ajouter la date et l'heure dans la colonne L
    ws.Cells(i, 12).Value = Now
    
    MsgBox "Mise à jour terminée avec succès !", vbInformation
End Sub


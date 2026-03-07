Attribute VB_Name = "SailingXML"
'''
' Auteur : Mttwt9 (GitHub) # FFV 1377942G
' Licence : GNU GPL v3 (voir fichier LICENSE)
' Détail : Génération d'un fichier XML (XRR) pour SCORE (FFVoile) à partir des données d'une feuille Excel.
' Encodage : Western European - Windows 1252 (CP1252)
' Version : 2026-03-07 -- XRR v2.1
'''
Sub CreateSailingXML()
    Dim xmlDoc As Object
    Dim rootNode As Object
    Dim personNode As Object
    Dim boatNode As Object
    Dim eventNode As Object
    Dim teamNode As Object
    Dim crewNode As Object
    Dim ws As Worksheet
    Dim firstRow As Integer: firstRow = 3 ' Première ligne de données
    Dim lastRow As Integer
    Dim i As Integer
    Dim eventCOID As String
    Dim eventEPID As String
    Dim filePath As String
    Dim desktopPath As String
    Dim boatID As String
    Dim teamID As Variant
    Dim skipperID As String
    Dim crewID As String
    Dim tempTeams As Object ' Collection temporaire des Teams
    Dim tempBoats As Object ' Collection temporaire des Boats
    Dim shell As Object


    ' Définition des constantes pour référencer les champs avec les numéros des colonnes
    ' Colonnes pour les attributs de Team
    Dim COL_COID As Integer: COL_COID = 1 ' ID de la competition
    Dim COL_EPID As Integer: COL_EPID = 2 ' ID de l'épreuve
    Dim COL_EP_CODEBAT As Integer: COL_EP_CODEBAT = 3 ' Code bateau de l'épreuve
    Dim COL_REMARKS As Integer: COL_REMARKS = 4 ' Remarques pour l'import
    Dim COL_NOC_TEAM As Integer: COL_NOC_TEAM = 5
    
    ' Colonnes pour les attributs de Boat
    Dim COL_NUM_VOILE As Integer: COL_NUM_VOILE = 6
    Dim COL_BOW_NUMBER As Integer: COL_BOW_NUMBER = 7
    Dim COL_NOM_VOILE As Integer: COL_NOM_VOILE = 8
    Dim COL_MODELE_VOILE As Integer: COL_MODELE_VOILE = 9
    Dim COL_OSIRS_GUEST As Integer: COL_OSIRS_GUEST = 10
    Dim COL_CAT_TEAM As Integer: COL_CAT_TEAM = 11
    Dim COL_RANKING_TEAM As Integer: COL_RANKING_TEAM = 12

    ' Colonnes pour les attributs du barreur
    Dim COL_NUM_LICENCE_BARREUR As Integer: COL_NUM_LICENCE_BARREUR = 13
    Dim COL_NOM_BARREUR As Integer: COL_NOM_BARREUR = 14
    Dim COL_PRENOM_BARREUR As Integer: COL_PRENOM_BARREUR = 15
    Dim COL_GENRE_BARREUR As Integer: COL_GENRE_BARREUR = 16
    Dim COL_NOC_BARREUR As Integer: COL_NOC_BARREUR = 17
    Dim COL_CLUB_BARREUR As Integer: COL_CLUB_BARREUR = 18
    Dim COL_DATE_NAISSANCE_BARREUR As Integer: COL_DATE_NAISSANCE_BARREUR = 19
    Dim COL_CLASS_ID_BARREUR As Integer: COL_CLASS_ID_BARREUR = 20
    Dim COL_WS_ID_BARREUR As Integer: COL_WS_ID_BARREUR = 21

    ' Colonnes pour les attributs de l'équipier
    Dim COL_NUM_LICENCE_EQUIP As Integer: COL_NUM_LICENCE_EQUIP = 22
    Dim COL_NOM_EQUIP As Integer: COL_NOM_EQUIP = 23
    Dim COL_PRENOM_EQUIP As Integer: COL_PRENOM_EQUIP = 24
    Dim COL_GENRE_EQUIP As Integer: COL_GENRE_EQUIP = 25
    Dim COL_NOC_EQUIP As Integer: COL_NOC_EQUIP = 26
    Dim COL_CLUB_EQUIP As Integer: COL_CLUB_EQUIP = 27
    Dim COL_DATE_NAISSANCE_EQUIP As Integer: COL_DATE_NAISSANCE_EQUIP = 28
    Dim COL_CLASS_ID_EQUIP As Integer: COL_CLASS_ID_EQUIP = 29
    Dim COL_WS_ID_EQUIP As Integer: COL_WS_ID_EQUIP = 30
    
   

    ' Création de l'objet document XML et ajout de la 1re ligne
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.appendChild xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")

    ' Création du nœud racine
    Set rootNode = xmlDoc.createElement("SailingXRR")
    xmlDoc.appendChild rootNode

    ' Définition de la feuille de calcul
    Set ws = ActiveSheet ' Utilise la feuille actuellement sélectionnée
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Création des collections temporaires pour stocker les Teams et Boats
    Set tempTeams = CreateObject("Scripting.Dictionary")
    Set tempBoats = CreateObject("Scripting.Dictionary")

    ' Boucle sur les lignes (chaque ligne contient un bateau et 1 ou 2 coureurs)
    For i = firstRow To lastRow
        ' Définition de l'eventCOID et eventEPID depuis la ligne courante
        eventCOID = ws.Cells(i, COL_COID).Value
        eventEPID = ws.Cells(i, COL_EPID).Value

        ' Génération des identifiants
        boatID = eventCOID & "_B" & i
        teamID = eventCOID & "_T" & i
        skipperID = eventCOID & "_P" & i & "_1"
        crewID = eventCOID & "_P" & i & "_2"

        ' Création du skipper
        Set personNode = xmlDoc.createElement("Person")
        personNode.setAttribute "PersonID", skipperID
        personNode.setAttribute "FamilyName", Trim(ws.Cells(i, COL_NOM_BARREUR).Value)
        personNode.setAttribute "GivenName", Trim(ws.Cells(i, COL_PRENOM_BARREUR).Value)
        personNode.setAttribute "NOC", Trim(ws.Cells(i, COL_NOC_BARREUR).Value)
        personNode.setAttribute "Gender", Trim(ws.Cells(i, COL_GENRE_BARREUR).Value)
        personNode.setAttribute "BirthDate", Format(ws.Cells(i, COL_DATE_NAISSANCE_BARREUR).Value, "yyyy-mm-dd")
        personNode.setAttribute "FFVLicenseNumber", Trim(ws.Cells(i, COL_NUM_LICENCE_BARREUR).Value)
        personNode.setAttribute "ClubName", Trim(ws.Cells(i, COL_CLUB_BARREUR).Value)
        personNode.setAttribute "IFPersonID", Trim(ws.Cells(i, COL_WS_ID_BARREUR).Value)
        personNode.setAttribute "ClassPersonID", Trim(ws.Cells(i, COL_CLASS_ID_BARREUR).Value)
        rootNode.appendChild personNode

        ' Création de l'équipier s'il existe
        If Trim(ws.Cells(i, COL_NOM_EQUIP).Value) <> "" Or Trim(ws.Cells(i, COL_NUM_LICENCE_EQUIP).Value) <> "" Then
            Set personNode = xmlDoc.createElement("Person")
            personNode.setAttribute "PersonID", crewID
            personNode.setAttribute "FamilyName", Trim(ws.Cells(i, COL_NOM_EQUIP).Value)
            personNode.setAttribute "GivenName", Trim(ws.Cells(i, COL_PRENOM_EQUIP).Value)
            personNode.setAttribute "NOC", Trim(ws.Cells(i, COL_NOC_EQUIP).Value)
            personNode.setAttribute "Gender", Trim(ws.Cells(i, COL_GENRE_EQUIP).Value)
            personNode.setAttribute "BirthDate", Format(ws.Cells(i, COL_DATE_NAISSANCE_EQUIP).Value, "yyyy-mm-dd")
            personNode.setAttribute "FFVLicenseNumber", Trim(ws.Cells(i, COL_NUM_LICENCE_EQUIP).Value)
            personNode.setAttribute "ClubName", Trim(ws.Cells(i, COL_CLUB_EQUIP).Value)
            personNode.setAttribute "IFPersonID", Trim(ws.Cells(i, COL_WS_ID_EQUIP).Value)
            personNode.setAttribute "ClassPersonID", Trim(ws.Cells(i, COL_CLASS_ID_EQUIP).Value)
            rootNode.appendChild personNode
        End If

        ' Création du bateau et mise en tampon pour sortir d'abord toutes les Persons puis les Boats et enfin l'Event qui contiendra les Teams
        Set boatNode = xmlDoc.createElement("Boat")
        boatNode.setAttribute "BoatID", boatID
        If Trim(ws.Cells(i, COL_NUM_VOILE).Value) <> "" Then
            boatNode.setAttribute "SailNumber", Trim(ws.Cells(i, COL_NUM_VOILE).Value)
        Else
            boatNode.setAttribute "SailNumber", Trim(ws.Cells(i, COL_NOM_BARREUR).Value) & "_" & i
        End If
        boatNode.setAttribute "BoatName", ""
        boatNode.setAttribute "BowNumber", ""
        boatNode.setAttribute "BoatModel", Trim(ws.Cells(i, COL_MODELE_VOILE).Value)
        If ws.Cells(i, COL_OSIRS_GUEST).Value = 1 Then
            boatNode.setAttribute "OsirisGuest", "true"
        End If
        tempBoats.Add boatID, boatNode

        ' Création de l'équipe et mise en tampon
        Set teamNode = xmlDoc.createElement("Team")
        teamNode.setAttribute "TeamID", teamID
        teamNode.setAttribute "BoatID", boatID
        teamNode.setAttribute "NOC", Trim(ws.Cells(i, COL_NOC_TEAM).Value)
        teamNode.setAttribute "Cat", Trim(ws.Cells(i, COL_CAT_TEAM).Value)
        
        ' Ajout des attributs optionnels s'ils sont renseignés
        If Trim(ws.Cells(i, COL_RANKING_TEAM).Value) <> "" Then
            teamNode.setAttribute "Ranking", Trim(ws.Cells(i, COL_RANKING_TEAM).Value)
        End If
        If Trim(ws.Cells(i, COL_EPID).Value) <> "" Then
            teamNode.setAttribute "EpID", Trim(ws.Cells(i, COL_EPID).Value)
        End If
        If Trim(ws.Cells(i, COL_EP_CODEBAT).Value) <> "" Then
            teamNode.setAttribute "EpBoatCode", Trim(ws.Cells(i, COL_EP_CODEBAT).Value)
        End If
        If Trim(ws.Cells(i, COL_REMARKS).Value) <> "" Then
            teamNode.setAttribute "Remarks", Trim(ws.Cells(i, COL_REMARKS).Value)
        End If

        ' Ajout du skipper dans l'équipe
        Set crewNode = xmlDoc.createElement("Crew")
        crewNode.setAttribute "PersonID", skipperID
        crewNode.setAttribute "Position", "S"
        teamNode.appendChild crewNode

        ' Ajout de l'équipier dans l'équipe s'il existe
        If Trim(ws.Cells(i, COL_NOM_EQUIP).Value) <> "" Then
            Set crewNode = xmlDoc.createElement("Crew")
            crewNode.setAttribute "PersonID", crewID
            crewNode.setAttribute "Position", "C"
            teamNode.appendChild crewNode
        End If

        ' Stocker l'équipe temporairement
        tempTeams.Add teamID, teamNode
    Next i

    ' Ajout des Boats après les Persons
    Dim boatIDKey As Variant
    For Each boatIDKey In tempBoats.Keys
        rootNode.appendChild tempBoats(boatIDKey)
    Next boatIDKey

    ' Ajout d'un unique nœud Event contenant toutes les Teams, après les Boats
    Set eventNode = xmlDoc.createElement("Event")
    eventNode.setAttribute "CoID", eventCOID
    rootNode.appendChild eventNode
    
    ' Ajout de toutes les Teams dans le nœud Event
    For Each teamID In tempTeams.Keys
        eventNode.appendChild tempTeams(teamID)
    Next teamID

    ' Définition du chemin du fichier avec la date du jour
    Set shell = CreateObject("WScript.Shell")
    desktopPath = shell.SpecialFolders("Desktop")

    filePath = desktopPath & "\SailingXRR_" & Format(Date, "yyyy-mm-dd") & ".xml"
    ' Sauvegarde du fichier XML
    xmlDoc.Save filePath

    MsgBox "Fichier XML (XRR) créé avec succès à " & filePath
End Sub

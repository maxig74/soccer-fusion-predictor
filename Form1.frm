VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Pronostico Calcio 1.6 Fusion automatico di Massimiliano G."
   ClientHeight    =   10545
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   18990
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   18990
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLabelBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   11760
      ScaleHeight     =   105
      ScaleWidth      =   1905
      TabIndex        =   6
      Top             =   120
      Width           =   1935
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   960
      ScaleHeight     =   1425
      ScaleWidth      =   4305
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Timer tmrRefresh 
      Left            =   7200
      Top             =   8040
   End
   Begin MSComctlLib.ListView lvPronostico 
      Height          =   2055
      Left            =   8280
      TabIndex        =   3
      Top             =   5640
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvPartite 
      Height          =   2055
      Left            =   600
      TabIndex        =   2
      Top             =   5640
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvClassifica 
      Height          =   4695
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblSchedina 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   13080
      TabIndex        =   4
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Label lblTitolo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' DICHIARAZIONI API PER IL MENU POPUP (DA METTERE IN CIMA AL FORM O IN FONDO)
' ==============================================================================
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal prcRect As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const TPM_RETURNCMD = &H100
Private Const MF_STRING = &H0

' ********************
' * DICHIARAZIONI API PER FLICKERING FIX *
' ********************
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Const WM_CLOSE = &H10

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0 ' larghezza schermo in pixel
Private Const SM_CYSCREEN = 1 ' altezza schermo in pixel
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long

Private Const WM_SETREDRAW = &HB

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Private Const MF_BYCOMMAND As Long = &H0&
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_RESTORE As Long = &HF120&

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE As Long = (-16)
Private Const WS_THICKFRAME As Long = &H40000

' --------------------
' ListView message constants (più leggibili)
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SCROLL As Long = (LVM_FIRST + 20)
Private Const LVM_GETTOPINDEX As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE As Long = (LVM_FIRST + 40)
' --------------------

' ********************
' * COSTANTI DI COLORE E TESTO *
' ********************
Private Const COLORE_SFONDO_FORM As Long = &H3C3C3C
Private Const COLORE_SFONDO_LISTVIEW As Long = &H2B2B2B
Private Const COLORE_TESTO_PRIMARIO As Long = &HFFFFFF
Private Const COLORE_ACCENTO As Long = &H3498DB

Private Const COLORE_CHAMPIONS As Long = &HFFEF00
Private Const COLORE_EUROPA As Long = &H8000FF
Private Const COLORE_RETROCESSIONE As Long = &H4545FF
Private Const COLORE_SALVEZZA As Long = &HFFFFFF

Private Const COLORE_MI_POSITIVA As Long = &H70B048
Private Const COLORE_MI_NEGATIVA As Long = &H4747FF
Private Const COLORE_MI_ZERO As Long = &HA5FF
Private Const TESTO_TITOLO As String = "PRONOSTICO CALCIO FUSION AUTOMATICO"

' Layout / Resize
Private Const ALTEZZA_TITOLO As Long = 400
Private Const MARGINE_VERTICALE As Long = 120
Private Const LARGHEZZA_SCROLLBAR_VERTICALE As Long = 350
Private Const SCALA_MINIMA As Single = 0.75
Private Const LARGHEZZA_MINIMA_SQUADRA As Long = 1000

' Logica classifica
Private Const POSIZIONI_CHAMPIONS As Integer = 4
Private Const POSIZIONI_EUROPA As Integer = 6
Private Const SQUADRE_RETROCESSIONE As Integer = 3

' INDICI LISTVIEW CLASSIFICA
Private Const INDICE_POSIZIONE As Integer = 0
Private Const INDICE_SQUADRA As Integer = 1
Private Const INDICE_PUNTI As Integer = 2
Private Const INDICE_GIOCATE As Integer = 3
Private Const INDICE_V As Integer = 4
Private Const INDICE_N As Integer = 5
Private Const INDICE_P As Integer = 6
Private Const INDICE_GF As Integer = 7
Private Const INDICE_GS As Integer = 8
Private Const INDICE_DR As Integer = 9
Private Const INDICE_MI As Integer = 10

Private Const INDICE_COLONNA_ELASTICA_CLASS As Integer = 2
Private Const INDICE_COLONNA_ELASTICA_PARTITE As Integer = 4
Private Const INDICE_COLONNA_ELASTICA_PRONOSTICO As Integer = 4
Private Const LUNGHEZZA_MAX_NOME As Long = 14

' Variabili form
Private BaseWidth As Long
Private BaseHeight As Long
Private LarghezzaTotaleColonneFisseClassifica As Long
Private LarghezzaTotaleColonneFissePartite As Long
Private LarghezzaTotaleColonneFissePronostico As Long
Private BaseFontSizeLV As Single
Private BaseFontSizeTitolo As Single
Private LastScaledFontSizeLV As Single
Private LastScaledFontSizeTitolo As Single
Private LastScaledFontSizeSchedina As Single
Private BaseLeft As Long
Private BaseTop As Long

Public UltimaChiamata As String
Private UltimaPercentuale As Integer
Private UltimoTempoStimato As String
Public ModelloGiaCalibrato As Boolean
Public EsciSubito As Boolean
Public InCalibrazione As Boolean
Private ChiusuraInCorso As Boolean

' ==============================================================================
' EVENTO TASTO DESTRO SULLA SCHEDINA
' ==============================================================================
Private Sub lblSchedina_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Intercetta il Tasto Destro del mouse (Button = 2 in VB6)
    If Button = 2 Then
        Dim hMenu As Long
        Dim Pt As POINTAPI
        Dim Scelta As Long
        
        ' Crea un menu popup al volo tramite API di Windows
        hMenu = CreatePopupMenu()
        
        ' Aggiunge la voce "Copia Schedina" con ID uguale a 1
        Call AppendMenu(hMenu, MF_STRING, 1, "Copia Schedina")
        
        ' Prende la posizione esatta del mouse sullo schermo
        Call GetCursorPos(Pt)
        
        ' Mostra il menu a tendina e intercetta il click dell'utente
        Scelta = TrackPopupMenu(hMenu, TPM_RETURNCMD, Pt.x, Pt.y, 0, Me.hwnd, 0&)
        
        ' Distrugge il menu dalla memoria per non sprecare risorse
        Call DestroyMenu(hMenu)
        
        ' Se l'utente ha cliccato su "Copia Schedina" (ID = 1) chiama la Sub dedicata
        If Scelta = 1 Then
            Call EseguiCopiaSchedina
        End If
    End If
End Sub

Private Sub EseguiCopiaSchedina()
    ' Verifica che la schedina contenga testo valido
    If Trim$(lblSchedina.Caption) <> "" And lblSchedina.Caption <> "Label1" Then
        
        Dim Righe() As String
        Dim RigaModificata As String
        Dim TestoFinaleSpedito As String
        Dim i As Long
        Dim IndicePartitaListView As Long
        Dim RigaCorrente As String
        Dim PosQuadra As Long
        Dim StringaGol As String
        
        ' Variabili per l'estrazione sicura del testo dalla colonna motivazione
        Dim TestoMotivazione As String
        Dim PosInizioGol As Long
        Dim PosFinePunto As Long
        Dim PunteggioEstratto As String
        
        ' Divide il testo visibile della schedina riga per riga
        Righe = Split(lblSchedina.Caption, vbCrLf)
        
        ' Inizializza l'intestazione con la versione del software preso dal titolo
        TestoFinaleSpedito = "Generato con: " & Me.Caption & vbCrLf & _
                             "--------------------------------------------------" & vbCrLf & vbCrLf
        
        ' Contatore per tracciare le righe di partita effettive presenti nella ListView
        IndicePartitaListView = 1
        
        ' Cicla tutte le righe della schedina testuale
        For i = LBound(Righe) To UBound(Righe)
            RigaCorrente = Trim$(Righe(i))
            
            ' Controllo 1: Salta le righe vuote o l'intestazione della giornata
            If Len(RigaCorrente) = 0 Or Left$(UCase(RigaCorrente), 8) = "SCHEDINA" Then
                TestoFinaleSpedito = TestoFinaleSpedito & Righe(i) & vbCrLf
            Else
                ' Imposta un fallback di sicurezza iniziale
                StringaGol = " (0-0)"
                
                ' Controllo di sicurezza per non superare il numero di elementi presenti nella ListView
                If IndicePartitaListView <= lvPronostico.ListItems.Count Then
                    ' Legge la motivazione estesa memorizzata nella colonna 4 (SubItems(3)) della ListView
                    TestoMotivazione = lvPronostico.ListItems(IndicePartitaListView).SubItems(3)
                    
                    ' FIX CRITICO: Cerca "Punteggio più probabile: " (la stringa esatta della versione 1.6)
                    PosInizioGol = InStr(TestoMotivazione, "Punteggio più probabile: ")
                    If PosInizioGol > 0 Then
                        ' Si sposta subito dopo la parola cercata (25 caratteri)
                        PosInizioGol = PosInizioGol + 25
                        
                        ' Cerca il punto fermo che delimita la fine del punteggio (es. dopo "2-1.")
                        PosFinePunto = InStr(PosInizioGol, TestoMotivazione, ".")
                        
                        If PosFinePunto > PosInizioGol Then
                            ' Estrae solo la stringa dei gol (es. "2-1") pulendola da spazi
                            PunteggioEstratto = Mid$(TestoMotivazione, PosInizioGol, PosFinePunto - PosInizioGol)
                            StringaGol = " (" & Trim$(PunteggioEstratto) & ")"
                        End If
                    End If
                End If
                
                ' Cerca dove si trova l'inizio delle parentesi quadre del blocco combo [
                PosQuadra = InStr(Righe(i), "[")
                
                If PosQuadra > 0 Then
                    ' Inietta i gol estratti reali mantenendo la spaziatura ordinata prima di [
                    RigaModificata = Left$(Righe(i), PosQuadra - 1) & StringaGol & " " & Mid$(Righe(i), PosQuadra)
                Else
                    ' Se non ci sono quadre, appende semplicemente i gol alla fine del rigo
                    RigaModificata = Righe(i) & StringaGol
                End If
                
                TestoFinaleSpedito = TestoFinaleSpedito & RigaModificata & vbCrLf
                
                ' Passa alla partita successiva della ListView
                IndicePartitaListView = IndicePartitaListView + 1
            End If
        Next i
        
        ' Svuota gli appunti e copia il testo finale completo di tutto
        Clipboard.Clear
        Clipboard.SetText TestoFinaleSpedito
        
        ' Notifica visiva finale
        MsgBox "Schedina (completa di Versione e Gol Probabili Reali) copiata negli appunti!", vbInformation, "Copia Schedina"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Se la chiusura è già stata autorizzata, lascia scaricare la form
    If ChiusuraInCorso Then Exit Sub
    
    ' Attiva il flag di stop immediato per interrompere i calcoli
    EsciSubito = True
    tmrRefresh.Enabled = False
    
    ' SE STA CALIBRANDO: blocchiamo l'Unload per un millisecondo per evitare crash.
    ' Ci penserà il ciclo For a interrompersi e a chiudere l'applicazione.
    If InCalibrazione Then
        Cancel = True
        Exit Sub
    End If
    
    ' SE NON STA CALIBRANDO: Chiudiamo normalmente al 1° Click
    ChiusuraInCorso = True
    Call ModuloPronostico.TerminaApplicazione
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Pulizia finale dei riferimenti grafici
    On Error Resume Next
    Set lvClassifica.SelectedItem = Nothing
    Set lvPartite.SelectedItem = Nothing
    Set lvPronostico.SelectedItem = Nothing
    
    lvClassifica.ListItems.Clear
    lvPartite.ListItems.Clear
    lvPronostico.ListItems.Clear
End Sub

Private Sub DisabilitaMassimizza()
    Dim hMenu As Long
    hMenu = GetSystemMenu(Me.hwnd, 0)
    If hMenu <> 0 Then
        RemoveMenu hMenu, SC_MAXIMIZE, MF_BYCOMMAND   ' blocca Massimizza
        RemoveMenu hMenu, SC_RESTORE, MF_BYCOMMAND    ' blocca Ripristina
        DrawMenuBar Me.hwnd
    End If
End Sub

Private Sub AbilitaMassimizza()
    Call GetSystemMenu(Me.hwnd, 1)
    DrawMenuBar Me.hwnd
End Sub

' ************************************************************
' * FUNZIONI DI LOGICA E AGGIORNAMENTO DATI
' ************************************************************

Private Sub PopolaClassificaIniziale()
UltimaChiamata = "PopolaClassificaIniziale"


    ' Carica risultati, calibra modello, ordina e aggiorna UI
    Call ModuloPronostico.CaricaEAnalizzaRisultati
    Call ModuloPronostico.OrdinaClassifica
    Call AggiornaListView

    'Call AggiornaPronostici
    
    ' Imposta il layout dopo il caricamento dati
    Form_Resize
End Sub

Private Sub Form_Activate()
Form_Resize
End Sub

' ************************************************************
' * AGGIORNAMENTO LISTVIEW (COLORI CLASSIFICA E MI)
' ************************************************************
Private Function ColorePosizione(Pos As Long, UltimaPosizione As Long) As Long
    Select Case Pos
        Case 1 To POSIZIONI_CHAMPIONS
            ColorePosizione = COLORE_CHAMPIONS
        Case POSIZIONI_CHAMPIONS + 1 To POSIZIONI_EUROPA
            ColorePosizione = COLORE_EUROPA
        Case UltimaPosizione - SQUADRE_RETROCESSIONE + 1 To UltimaPosizione
            ColorePosizione = COLORE_RETROCESSIONE
        Case Else
            ColorePosizione = COLORE_SALVEZZA
    End Select
End Function

Private Sub AggiornaListView()
If Not Me.Visible Or Me.hwnd = 0 Or EsciSubito Then Exit Sub
    Dim i As Long
    Dim ListItem As ListItem
    Dim PartitaObj As Partita

    ' DISABILITA RIDISEGNO
    Call SendMessage(lvClassifica.hwnd, WM_SETREDRAW, 0, 0)
    Call SendMessage(lvPartite.hwnd, WM_SETREDRAW, 0, 0)
    
    ' 1) CLASSIFICA
    lvClassifica.ListItems.Clear
    Dim UltimaPosizione As Long: UltimaPosizione = ModuloPronostico.ContaSquadre
    For i = 1 To UltimaPosizione
        Set ListItem = lvClassifica.ListItems.Add(, , i)
        ListItem.Bold = True
        ListItem.ForeColor = ColorePosizione(i, UltimaPosizione)
        ListItem.SubItems(INDICE_SQUADRA) = ModuloPronostico.Squadre(i).Nome
        With ModuloPronostico.Squadre(i).Stats
            ListItem.SubItems(INDICE_PUNTI) = .Punti
            ListItem.SubItems(INDICE_GIOCATE) = .Giocate
            ListItem.SubItems(INDICE_V) = .Vinte
            ListItem.SubItems(INDICE_N) = .Pareggiate
            ListItem.SubItems(INDICE_P) = .Perse
            ListItem.SubItems(INDICE_GF) = .GoalFatti
            ListItem.SubItems(INDICE_GS) = .GoalSubiti
            ListItem.SubItems(INDICE_DR) = .DifferenzaReti
            
            Dim MediaInglese As Long
            ListItem.SubItems(INDICE_MI) = .MediaInglese
            
            With lvClassifica.ListItems(i).ListSubItems(INDICE_MI)
                If ModuloPronostico.Squadre(i).Stats.MediaInglese > 0 Then
                    .ForeColor = COLORE_MI_POSITIVA
                    .Bold = True
                ElseIf ModuloPronostico.Squadre(i).Stats.MediaInglese < 0 Then
                    .ForeColor = COLORE_MI_NEGATIVA
                    .Bold = True
                Else
                    .ForeColor = RGB(255, 128, 0) ' Arancione puro se è esattamente 0
                    .Bold = False
                End If
            End With
            
            ListItem.SubItems(INDICE_MI + 1) = .VinteC
            ListItem.SubItems(INDICE_MI + 2) = .PareggiateC
            ListItem.SubItems(INDICE_MI + 3) = .PerseC
            ListItem.SubItems(INDICE_MI + 4) = .GoalFattiC
            ListItem.SubItems(INDICE_MI + 5) = .GoalSubitiC
            ListItem.SubItems(INDICE_MI + 6) = .VinteT
            ListItem.SubItems(INDICE_MI + 7) = .PareggiateT
            ListItem.SubItems(INDICE_MI + 8) = .PerseT
            ListItem.SubItems(INDICE_MI + 9) = .GoalFattiT
            ListItem.SubItems(INDICE_MI + 10) = .GoalSubitiT
        End With
    Next i
    
    ' 2) PARTITE (ordine decrescente: ultima giornata sopra)
    lvPartite.ListItems.Clear
    If ModuloPronostico.ContaPartite = 0 Then GoTo FineAggiornamento
    
    Dim UltimoIndiceTotale As Long
    UltimoIndiceTotale = ModuloPronostico.ContaPartite - 1
    
    For i = UltimoIndiceTotale To 0 Step -1
        Set PartitaObj = ModuloPronostico.PartiteGiocate(i)
        With PartitaObj
            Set ListItem = lvPartite.ListItems.Add(, , .Giornata)
            ListItem.SubItems(1) = .Casa
            If .GoalCasa < 0 Or .GoalOspite < 0 Then
                ListItem.SubItems(2) = "-"
            Else
                ListItem.SubItems(2) = CStr(.GoalCasa) & " - " & CStr(.GoalOspite)
            End If
            ListItem.SubItems(3) = .Ospite
            ListItem.ForeColor = COLORE_TESTO_PRIMARIO
        End With
        Set PartitaObj = Nothing
    Next i

FineAggiornamento:
    ' 3) POSIZIONA SULLA GIORNATA SUCCESSIVA A QUELLE GIOCATE
    On Error Resume Next

    Dim GiornataDaMostrare As Long
    Dim iPart As Long, iItem As Long
    Dim offset As Long, Trovata As Boolean
    GiornataDaMostrare = 0

    For iPart = 0 To ModuloPronostico.ContaPartite - 1
        With ModuloPronostico.PartiteGiocate(iPart)
            If .GoalCasa < 0 Or .GoalOspite < 0 Then
                GiornataDaMostrare = .Giornata
                Exit For
            End If
        End With
    Next iPart

    If GiornataDaMostrare = 0 And ModuloPronostico.ContaPartite > 0 Then
        GiornataDaMostrare = ModuloPronostico.PartiteGiocate(ModuloPronostico.ContaPartite - 1).Giornata
    End If
    ' Cerca nella ListView (decrescente)
    For iItem = 1 To lvPartite.ListItems.Count
        If Val(lvPartite.ListItems(iItem).Text) = GiornataDaMostrare Then
            lvPartite.ListItems(iItem).Selected = True
            lvPartite.ListItems(iItem).EnsureVisible
            
            For offset = 1 To 9
                If iItem + offset <= lvPartite.ListItems.Count Then
                    lvPartite.ListItems(iItem + offset).EnsureVisible
                End If
            Next offset

            Trovata = True
            Exit For
        End If
    Next iItem

    If Not Trovata And lvPartite.ListItems.Count > 0 Then
        lvPartite.ListItems(1).Selected = True
        lvPartite.ListItems(1).EnsureVisible
    End If
    
    On Error GoTo 0

    ' 4) RIABILITA RIDISEGNO
    Call SendMessage(lvClassifica.hwnd, WM_SETREDRAW, 1, 0)
    Call SendMessage(lvPartite.hwnd, WM_SETREDRAW, 1, 0)
    
    lvClassifica.Refresh
    lvPartite.Refresh

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Se sta calibrando ? blocca TUTTI i tasti F
    If InCalibrazione Then
        If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
            KeyCode = 0
            Exit Sub
        End If
    End If

    ' --- F5: aggiorna la schermata ---
    If KeyCode = vbKeyF5 Then
        AggiornaListView
        KeyCode = 0
        Exit Sub
    End If
    
    ' --- F6: ripristina dimensioni originali ---
    If KeyCode = vbKeyF6 Then

        If Me.WindowState = vbMaximized Then
            Me.WindowState = vbNormal
        End If

        Me.Width = BaseWidth
        Me.Height = BaseHeight

        Me.Left = (Screen.Width - Me.Width) \ 2
        Me.Top = (Screen.Height - Me.Height) \ 2

        Form_Resize

        KeyCode = 0
        Exit Sub
    End If

    ' --- F8: calibrazione ---
       ' --- F8: calibrazione ---
    ' --- F8: calibrazione o visualizzazione parametri se già calibrato ---
    If KeyCode = vbKeyF8 Then

        If Not ModelloGiaCalibrato Then

            Dim risposta As VbMsgBoxResult
            risposta = MsgBox("Vuoi avviare la calibrazione iniziale?", vbYesNo + vbQuestion, "Calibrazione")

            If risposta = vbYes Then
                DisabilitaMassimizza
                InCalibrazione = True

                picProgress.Visible = True
                picLabelBox.Visible = True
                Form_Resize

                EsciSubito = False

                ' Avvia il calcolo
                CalibraModelloCompleto

                ' Se l'utente ha premuto X durante il calcolo, usciamo subito
                If EsciSubito Then
                    InCalibrazione = False
                    KeyCode = 0
                    Exit Sub
                End If

                DisegnaProgress 100, "Finito", picProgress, lblProgress


Dim startattesa As Long
startattesa = GetTickCount()
Do While GetTickCount() < startattesa + 1000
    DoEvents
    If EsciSubito Then
        InCalibrazione = False
        KeyCode = 0
        Exit Sub
        End If
                    
                    
                'Dim t As Single
                't = Timer
                'Do While Timer < t + 1
                 '   DoEvents
                 '   If EsciSubito Then
                 '       InCalibrazione = False
                 '       KeyCode = 0
                 '       Exit Sub
                 '   End If
                Loop

                picProgress.Visible = False
                picLabelBox.Visible = False

                AbilitaMassimizza

                lvPartite.Enabled = True
                lvClassifica.Enabled = True
                lvPronostico.Enabled = True
                lblSchedina.Enabled = True

                InCalibrazione = False

                OrdinaClassifica
                AggiornaListView
                AggiornaPronostici

                ' Stampa i parametri subito dopo la prima calibrazione
                ModuloPronostico.StampaParametriCorrenti
                ModelloGiaCalibrato = True
            End If
        Else
            ' --- MODIFICA F8: SE È GIÀ CALIBRATO, FA COMODO RIVEDERE I DATI ---
            MsgBox "Il modello è già calibrato. Ecco i parametri correnti ottimizzati:", vbInformation, "Calibrazione"
            
            ' Questa riga mancava nel blocco Else! Ora mostra subito la finestra dei parametri
            ModuloPronostico.StampaParametriCorrenti
            
            ModuloPronostico.CaricaEAnalizzaRisultati
            OrdinaClassifica
            AggiornaListView
            AggiornaPronostici

            KeyCode = 0
            Exit Sub
        End If
        KeyCode = 0
    End If

    ' --- NUOVO TASTO F9: VISUALIZZA PARAMETRI IN QUALSIASI MOMENTO ---
    If KeyCode = vbKeyF9 Then
        If ModelloGiaCalibrato Then
            ' Mostra direttamente la finestra con i dati calcolati
            ModuloPronostico.StampaParametriCorrenti
        Else
            ' Avviso se provi a premerlo prima di aver calibrato il modello
            MsgBox "Il modello non è ancora stato calibrato. Premi F8 per avviarla.", vbExclamation, "Parametri non disponibili"
        End If
        KeyCode = 0
        Exit Sub
    End If


End Sub

Private Sub AggiornaPronostici()
    Dim PartiteVar As Variant
    Dim PronosticiVar As Variant
    Dim PronosticoObj As PronosticoPartita
    Dim p As Long
    Dim ListItem As ListItem

    Dim UltimaGiornataGiocata As Long
    If ModuloPronostico.ContaPartite > 0 Then
        Dim MaxG As Long: MaxG = 0
        Dim i As Long
        For i = 0 To ModuloPronostico.ContaPartite - 1
            If ModuloPronostico.PartiteGiocate(i).Giornata > MaxG Then
                MaxG = ModuloPronostico.PartiteGiocate(i).Giornata
            End If
        Next i
        UltimaGiornataGiocata = MaxG
    Else
        UltimaGiornataGiocata = 0
    End If
    
    Dim GiornataCorrente As Long: GiornataCorrente = UltimaGiornataGiocata
    PartiteVar = ModuloPronostico.OttieniProssimePartite(GiornataCorrente)

    Call SendMessage(lvPronostico.hwnd, WM_SETREDRAW, 0, 0)
    lvPronostico.ListItems.Clear
    lblSchedina.Caption = ""
    
    If IsEmpty(PartiteVar) Then GoTo FinePronostici

    PronosticiVar = ModuloPronostico.PronosticaPartite(PartiteVar)
    If Not IsArray(PronosticiVar) Then GoTo FinePronostici

    Dim Lb As Long: Lb = LBound(PronosticiVar)
    Dim Ub As Long: Ub = UBound(PronosticiVar)

    If Ub >= Lb Then
        Dim SchedinaCompleta As String
        Dim RigaPartita As String
        Dim TitoloRiga As String
        Dim GiornataPronostico As Long

        ' Variabili di allineamento per la Label grafica (Senza colonna Gol)
        Dim AlignCasa As String
        Dim AlignOspite As String
        Dim AlignSegno As String

        If IsArray(PartiteVar) Then
            If UBound(PartiteVar) >= 0 Then
                GiornataPronostico = PartiteVar(0).Giornata
            Else
                GiornataPronostico = GiornataCorrente
            End If
        Else
            GiornataPronostico = GiornataCorrente
        End If

        TitoloRiga = "SCHEDINA GIORNATA " & GiornataPronostico & " (Pronostico)"
        SchedinaCompleta = TitoloRiga & vbCrLf & vbCrLf

        For p = Lb To Ub
            Set PronosticoObj = PronosticiVar(p)
            With PronosticoObj
                Set ListItem = lvPronostico.ListItems.Add(, , .Casa)
                Dim SegnoFinale As String: SegnoFinale = .PronosticoComplesso
                ListItem.SubItems(1) = .Ospite
                ListItem.SubItems(2) = SegnoFinale
                ListItem.SubItems(3) = .Motivazione
                
                ' === QUINTA COLONNA NELLA LISTVIEW ===
                ListItem.SubItems(4) = .PronosticoGG_NG & " (" & Format$(.ConfidenzaGG_NG, "0.0") & "%)"

                ' === COLORAZIONE DELLA RIGA ===
                Select Case SegnoFinale
                    Case "1", "2"
                        ListItem.ForeColor = RGB(60, 179, 113)
                        ListItem.Bold = True
                    Case "X"
                        ListItem.ForeColor = RGB(255, 165, 0)
                        ListItem.Bold = True
                    Case "1X", "X2", "12"
                        ListItem.ForeColor = RGB(135, 206, 250)
                        ListItem.Bold = False
                    Case "1X2"
                        ListItem.ForeColor = RGB(255, 99, 71)
                        ListItem.Bold = False
                    Case Else
                        ListItem.ForeColor = COLORE_TESTO_PRIMARIO
                        ListItem.Bold = False
                End Select
                
                ' === COLORAZIONE SPECIFICA GOL/NOGOL ===
                If .ConfidenzaGG_NG >= 54# Then
                    ListItem.ListSubItems(4).ForeColor = RGB(50, 205, 50)
                    ListItem.ListSubItems(4).Bold = True
                ElseIf .ConfidenzaGG_NG >= 48# Then
                    ListItem.ListSubItems(4).ForeColor = RGB(255, 165, 0)
                    ListItem.ListSubItems(4).Bold = True
                Else
                    ListItem.ListSubItems(4).ForeColor = RGB(220, 20, 60)
                    ListItem.ListSubItems(4).Bold = False
                End If

                ' ==============================================================================
                ' COMPOSIZIONE ALLINEATA, COMPATTA E SENZA GOL PROBABILI A VIDEO (lblSchedina)
                ' ==============================================================================
                ' Spazio fisso ridotto a 10 caratteri per tenere tutto vicino e stretto
                AlignCasa = Left$(.Casa & Space(10), 10)
                AlignOspite = Left$(.Ospite & Space(10), 10)
                AlignSegno = Left$(SegnoFinale & Space(3), 3)

                ' Compone la riga unendo i pezzi in modo compatto ed eliminando il vecchio blocco AlignGol
                RigaPartita = AlignCasa & "- " & AlignOspite & " " & AlignSegno & "[" & .PronosticoGG_NG & "]"

                SchedinaCompleta = SchedinaCompleta & RigaPartita & vbCrLf
            End With
            Set PronosticoObj = Nothing
        Next p

        lblSchedina.Caption = SchedinaCompleta
    End If

FinePronostici:
    Call SendMessage(lvPronostico.hwnd, WM_SETREDRAW, 1, 0)
    lvPronostico.Refresh
End Sub

' ************************************************************
' * GESTIONE CLICK E LEGGENDE
' ************************************************************

Private Sub lvClassifica_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ItemHit As ListItem
    Dim Legenda As String

    If Button = vbLeftButton Then
        Set ItemHit = lvClassifica.HitTest(x, y)
        If Not ItemHit Is Nothing Then
            Dim LarghezzaPosCol As Long: LarghezzaPosCol = lvClassifica.ColumnHeaders(1).Width
            Dim LarghezzaMICol As Long: LarghezzaMICol = lvClassifica.ColumnHeaders(INDICE_MI + 1).Width
            Dim ColonnaCliccata As String

            If x >= 0 And x < LarghezzaPosCol Then
                ColonnaCliccata = "POSIZIONE"
            Else
                Dim PosizioneInizioMI As Long: PosizioneInizioMI = 0
                Dim i As Integer
                For i = 1 To INDICE_MI
                    PosizioneInizioMI = PosizioneInizioMI + lvClassifica.ColumnHeaders(i).Width
                Next i
                If x >= PosizioneInizioMI And x < PosizioneInizioMI + LarghezzaMICol Then
                    ColonnaCliccata = "MI"
                End If
            End If

            Select Case ColonnaCliccata
                Case "POSIZIONE"
                    Dim Posizione As String: Posizione = ItemHit.Text
                    Dim Squadra As String: Squadra = ItemHit.SubItems(INDICE_SQUADRA)
                    Dim PosizioneInt As Long: PosizioneInt = CLng(Posizione)
                    Dim UltimaPosizione As Long: UltimaPosizione = lvClassifica.ListItems.Count
                    Legenda = "Posizione " & Posizione & " " & UCase(Squadra) & vbCrLf & vbCrLf
                    Select Case PosizioneInt
                        Case 1 To POSIZIONI_CHAMPIONS: Legenda = Legenda & "CHAMPIONS LEAGUE"
                        Case POSIZIONI_CHAMPIONS + 1 To POSIZIONI_EUROPA: Legenda = Legenda & "EUROPA/CONFERENCE LEAGUE"
                        Case UltimaPosizione - SQUADRE_RETROCESSIONE + 1 To UltimaPosizione: Legenda = Legenda & "ZONA RETROCESSIONE"
                        Case Else: Legenda = Legenda & "ZONA SALVEZZA"
                    End Select
                    MsgBox Legenda, vbInformation, "Legenda Classifica"
                Case "MI"
                    Dim MI_Valore As Long: MI_Valore = CLng(ItemHit.ListSubItems(INDICE_MI).Text)
                    Dim DescrizioneMI As String
                    If MI_Valore > 0 Then
                        DescrizioneMI = "SOVRAPERFORMANCE. La squadra ha più punti di quanti ne servono per l'obiettivo (3 punti/partita)."
                    ElseIf MI_Valore < 0 Then
                        DescrizioneMI = "SOTTOPERFORMANCE. La squadra è in ritardo rispetto all'obiettivo dei 3 punti/partita."
                    Else
                        DescrizioneMI = "OBIETTIVO RAGGIUNTO. La squadra ha esattamente 3 punti in media a partita giocata."
                    End If
                    Legenda = "MEDIA INGLESE (" & MI_Valore & "):" & vbCrLf & vbCrLf & DescrizioneMI
                    MsgBox Legenda, vbInformation, "Legenda Media Inglese"
            End Select
        End If
    End If
End Sub

Private Sub lvPartite_KeyDown(KeyCode As Integer, Shift As Integer)

If InCalibrazione Then
    KeyCode = 0
    Exit Sub
End If

    If KeyCode <> vbKeyReturn Then Exit Sub

    Dim itm As ListItem
    Dim Ris As String
    Dim Casa As String, Ospite As String
    Dim Giornata As String
    Dim NuovoRis As String

    If lvPartite.SelectedItem Is Nothing Then Exit Sub
    Set itm = lvPartite.SelectedItem

    Dim SelectedIndex As Long
    SelectedIndex = itm.Index

    Giornata = Trim$(itm.Text)
    Casa = Trim$(itm.SubItems(1))
    Ris = itm.SubItems(2)
    Ospite = Trim$(itm.SubItems(3))

    NuovoRis = InputBox( _
        "Modifica risultato per " & Casa & " - " & Ospite & vbCrLf & _
        "Inserisci nel formato: golCasa-golOspite (es. 2-1). Lascia vuoto per '-'.", _
        "Modifica risultato - Giornata " & Giornata, Ris)

    If StrPtr(NuovoRis) = 0 Then Exit Sub
    If Len(Trim$(NuovoRis)) = 0 Then NuovoRis = "-"

    If NuovoRis <> "-" Then
        If InStr(NuovoRis, "-") = 0 Then
            MsgBox "Formato non valido. Usa ad es. 2-1 oppure '-'.", vbExclamation
            Exit Sub
        Else
            Dim parts() As String
            parts = Split(NuovoRis, "-")
            If UBound(parts) <> 1 Or _
               Not (IsNumeric(Trim$(parts(0))) And IsNumeric(Trim$(parts(1)))) Then
                MsgBox "Formato non valido. Usa ad es. 2-1 oppure '-'.", vbExclamation
                Exit Sub
            End If
        End If
    End If

    itm.SubItems(2) = NuovoRis

    ModuloPronostico.AggiornaRisultatoNelFile _
        App.Path & "\Risultati.txt", CLng(Giornata), Casa, Ospite, NuovoRis

    ModuloPronostico.CaricaEAnalizzaRisultati

    If ModelloGiaCalibrato Then
        ModuloPronostico.ApplicaParametri ModuloPronostico.ParametriCalibrati
    End If

    OrdinaClassifica
    AggiornaListView
    AggiornaPronostici

    tmrRefresh.Enabled = False
    tmrRefresh.Interval = 5
    tmrRefresh.Enabled = True

    On Error Resume Next
    If SelectedIndex > 0 And SelectedIndex <= lvPartite.ListItems.Count Then
        lvPartite.ListItems(SelectedIndex).EnsureVisible
        lvPartite.ListItems(SelectedIndex).Selected = True
    Else
        If lvPartite.ListItems.Count > 0 Then
            lvPartite.ListItems(1).EnsureVisible
            lvPartite.ListItems(1).Selected = True
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub lvPronostico_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ItemHit As ListItem
    Dim Legenda As String

    If Button = vbLeftButton Then
        Set ItemHit = lvPronostico.HitTest(x, y)
        If Not ItemHit Is Nothing Then
            Dim LarghezzaCasaCol As Long: LarghezzaCasaCol = lvPronostico.ColumnHeaders(1).Width
            If x >= 0 And x < LarghezzaCasaCol Then
                Dim Pronostico As String: Pronostico = ItemHit.SubItems(2)
                Legenda = "INFORMAZIONI SUL PRONOSTICO:" & vbCrLf & vbCrLf & _
                          "Partita: " & ItemHit.Text & " - " & ItemHit.SubItems(1) & vbCrLf & _
                          "Segno Pronosticato: " & UCase(Pronostico) & vbCrLf & vbCrLf
                Select Case Pronostico
                    Case "1", "2": Legenda = Legenda & "SEGNO FISSO(VERDE): Alta sicurezza. Il pronostico prevede una vittoria."
                    Case "X": Legenda = Legenda & "SEGNO FISSO(ARANCIONE): Alta sicurezza. Il pronostico prevede un pareggio."
                    Case "1X", "X2", "12": Legenda = Legenda & "DOPPIA CHANCE(CELESTE): Sicurezza media. Due esiti su tre sono considerati probabili."
                    Case "1X2": Legenda = Legenda & "TRIPLA(ROSSO): Bassa sicurezza / Rischio elevato. Partita estremamente incerta o bilanciata."
                    Case Else: Legenda = Legenda & "NESSUN PRONOSTICO: Dati insufficienti o errore nell'analisi."
                End Select
                MsgBox Legenda, vbInformation, "Legenda Pronostico"
            End If
        End If
    End If
End Sub

' ************************************************************
' * INIZIALIZZAZIONE LISTVIEW
' ************************************************************

Private Sub SetupClassificaLV()
    Dim Col As ColumnHeader
    Dim i As Integer
    With lvClassifica
        .View = lvwReport: .FullRowSelect = True: .GridLines = True
        .BackColor = COLORE_SFONDO_LISTVIEW: .ForeColor = COLORE_TESTO_PRIMARIO
        .Font.Name = "Tahoma": .Font.Bold = True: .Font.Size = BaseFontSizeLV
        .ColumnHeaders.Clear

        .ColumnHeaders.Add 1, , "Pos.", 700, lvwColumnLeft: .ColumnHeaders(1).Tag = 700
        .ColumnHeaders.Add 2, , "Squadra", 2500: .ColumnHeaders(2).Tag = 2500
        .ColumnHeaders.Add 3, , "Punti", 1600, lvwColumnCenter: .ColumnHeaders(3).Tag = 1600
        .ColumnHeaders.Add 4, , "G", 800, lvwColumnCenter: .ColumnHeaders(4).Tag = 800
        .ColumnHeaders.Add 5, , "V", 800, lvwColumnCenter: .ColumnHeaders(5).Tag = 800
        .ColumnHeaders.Add 6, , "N", 800, lvwColumnCenter: .ColumnHeaders(6).Tag = 800
        .ColumnHeaders.Add 7, , "P", 800, lvwColumnCenter: .ColumnHeaders(7).Tag = 800
        .ColumnHeaders.Add 8, , "GF", 800, lvwColumnCenter: .ColumnHeaders(8).Tag = 800
        .ColumnHeaders.Add 9, , "GS", 800, lvwColumnCenter: .ColumnHeaders(9).Tag = 800
        .ColumnHeaders.Add 10, , "DR", 800, lvwColumnCenter: .ColumnHeaders(10).Tag = 800
        .ColumnHeaders.Add 11, , "MI", 800, lvwColumnCenter: .ColumnHeaders(11).Tag = 800

        .ColumnHeaders.Add 12, , "VC", 800, lvwColumnCenter: .ColumnHeaders(12).Tag = 800
        .ColumnHeaders.Add 13, , "NC", 800, lvwColumnCenter: .ColumnHeaders(13).Tag = 800
        .ColumnHeaders.Add 14, , "PC", 800, lvwColumnCenter: .ColumnHeaders(14).Tag = 800
        .ColumnHeaders.Add 15, , "GFC", 800, lvwColumnCenter: .ColumnHeaders(15).Tag = 800
        .ColumnHeaders.Add 16, , "GSC", 800, lvwColumnCenter: .ColumnHeaders(16).Tag = 800
        .ColumnHeaders.Add 17, , "VT", 800, lvwColumnCenter: .ColumnHeaders(17).Tag = 800
        .ColumnHeaders.Add 18, , "NT", 800, lvwColumnCenter: .ColumnHeaders(18).Tag = 800
        .ColumnHeaders.Add 19, , "PT", 800, lvwColumnCenter: .ColumnHeaders(19).Tag = 800
        .ColumnHeaders.Add 20, , "GFT", 800, lvwColumnCenter: .ColumnHeaders(20).Tag = 800
        .ColumnHeaders.Add 21, , "GST", 800, lvwColumnCenter: .ColumnHeaders(21).Tag = 800

        LarghezzaTotaleColonneFisseClassifica = 0
        For i = 1 To .ColumnHeaders.Count
            If i <> INDICE_COLONNA_ELASTICA_CLASS Then
                LarghezzaTotaleColonneFisseClassifica = LarghezzaTotaleColonneFisseClassifica + .ColumnHeaders(i).Tag
            End If
        Next i
    End With
End Sub

Private Sub SetupPartiteLV()
    Dim i As Integer
    With lvPartite
        .View = lvwReport: .FullRowSelect = True: .GridLines = True
        .BackColor = COLORE_SFONDO_LISTVIEW: .ForeColor = COLORE_TESTO_PRIMARIO
        .Font.Name = "Tahoma": .Font.Bold = True: .Font.Size = BaseFontSizeLV
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Giornata", 1500, lvwColumnLeft: .ColumnHeaders(1).Tag = 1500
        .ColumnHeaders.Add 2, , "Casa", 2000, lvwColumnRight: .ColumnHeaders(2).Tag = 2000
        .ColumnHeaders.Add 3, , "Risultato", 1500, lvwColumnCenter: .ColumnHeaders(3).Tag = 1500
        .ColumnHeaders.Add 4, , "Ospite", 2000, lvwColumnLeft: .ColumnHeaders(4).Tag = 2000

        LarghezzaTotaleColonneFissePartite = 0
        For i = 1 To .ColumnHeaders.Count
            If i <> INDICE_COLONNA_ELASTICA_PARTITE Then
                LarghezzaTotaleColonneFissePartite = LarghezzaTotaleColonneFissePartite + .ColumnHeaders(i).Tag
            End If
        Next i
    End With
End Sub

Private Sub SetupPronosticoLV()
    Dim i As Integer
    With lvPronostico
        .View = lvwReport: .FullRowSelect = True: .GridLines = True
        .BackColor = COLORE_SFONDO_LISTVIEW: .ForeColor = COLORE_TESTO_PRIMARIO
        .Font.Name = "Tahoma": .Font.Bold = True: .Font.Size = BaseFontSizeLV
        .ColumnHeaders.Clear
        
        ' Struttura a 5 colonne con Gol/NoGol alla fine
        .ColumnHeaders.Add 1, , "Casa", 1400, lvwColumnLeft: .ColumnHeaders(1).Tag = 1400
        .ColumnHeaders.Add 2, , "Ospite", 1400, lvwColumnLeft: .ColumnHeaders(2).Tag = 1400
        .ColumnHeaders.Add 3, , "1X2", 600, lvwColumnCenter: .ColumnHeaders(3).Tag = 600
        .ColumnHeaders.Add 4, , "Motivazione (Analisi Statistica e Forza Interna)", 5500, lvwColumnLeft
        .ColumnHeaders.Add 5, , "Gol/NoGol", 3100, lvwColumnCenter: .ColumnHeaders(5).Tag = 3100
        
        LarghezzaTotaleColonneFissePronostico = 0
        For i = 1 To .ColumnHeaders.Count
            If i <> INDICE_COLONNA_ELASTICA_PRONOSTICO Then
                LarghezzaTotaleColonneFissePronostico = LarghezzaTotaleColonneFissePronostico + .ColumnHeaders(i).Tag
            End If
        Next i
    End With
End Sub


' ************************************************************
' * EVENTI FORM
' ************************************************************

Private Sub Form_Load()
Dim style As Long
style = GetWindowLong(Me.hwnd, GWL_STYLE)

' Rimuove il bordo ridimensionabile
style = style And Not WS_THICKFRAME

Call SetWindowLong(Me.hwnd, GWL_STYLE, style)

    ' Il modello parte SEMPRE come non calibrato
    ModelloGiaCalibrato = False
    ModuloPronostico.ParametriCalibratiValidi = False

    ' Base per scaling
    BaseWidth = Me.ScaleWidth
    BaseHeight = Me.ScaleHeight

    ' Imposto dimensioni BASE molto più grandi (così scalano bene)
    picProgress.Width = 6000
    picProgress.Height = 300

    picLabelBox.Width = 7000
    picLabelBox.Height = 300

    lblProgress.Width = 5800
    lblProgress.Height = 300

    ' Salvo dimensioni base per scaling (LARGHEZZA|ALTEZZA)
    picProgress.Tag = "6000|300"
    picLabelBox.Tag = "7000|300"
    lblProgress.Tag = "7000|300"

    ' Font base della label
    lblProgress.Font.Size = 12
    lblProgress.Font.Bold = True
    lblProgress.AutoSize = False
    lblProgress.WordWrap = False

    ' Invisibili all’avvio
    picProgress.Visible = False
    picLabelBox.Visible = False

    ' Posizionamento form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Me.BackColor = COLORE_SFONDO_FORM

    ' Font base
    BaseFontSizeTitolo = 14
    BaseFontSizeLV = 9.75

    LastScaledFontSizeLV = 0
    LastScaledFontSizeTitolo = 0
    LastScaledFontSizeSchedina = 0

    ' Setup ListView
    Call SetupClassificaLV
    Call SetupPartiteLV
    Call SetupPronosticoLV

    ' Titolo
    With lblTitolo
        .Caption = TESTO_TITOLO
        .BackStyle = 0
        .Font.Name = "Arial Black"
        .Font.Bold = True
        .Alignment = 2
        .ForeColor = COLORE_ACCENTO
    End With

    ' Schedina
    With lblSchedina
        .BackStyle = 1
        .BackColor = COLORE_SFONDO_LISTVIEW
        .ForeColor = COLORE_TESTO_PRIMARIO
        .BorderStyle = 1
        .Font.Name = "Courier New"
        .Font.Size = 10
        .Alignment = 0
    End With

    ' Dimensioni base
    BaseWidth = Me.Width
    BaseHeight = Me.Height
    BaseLeft = Me.Left
    BaseTop = Me.Top

    ' Carica squadre
    Call ModuloPronostico.CaricaSquadreValide

    ' Popola classifica iniziale
    Call PopolaClassificaIniziale


    Me.KeyPreview = True

End Sub

Public Sub DisegnaProgress(ByVal percentuale As Integer, ByVal tempoStimato As String, pic As PictureBox, lbl As Label)

    Dim w As Single, h As Single, wFill As Single
    Dim testo As String
    Dim rStart As Integer, gStart As Integer, bStart As Integer
    Dim rEnd As Integer, gEnd As Integer, bEnd As Integer
    Dim r As Integer, G As Integer, b As Integer
    Dim t As Double
    Dim baseW As Single, fattore As Single
    Dim parts() As String

    ' --- dimensioni attuali ---
    w = pic.ScaleWidth
    h = pic.ScaleHeight
    wFill = (percentuale / 100) * w

    ' --- base scaling ---
    If Len(pic.Tag) > 0 And InStr(pic.Tag, "|") > 0 Then
        parts = Split(pic.Tag, "|")
        baseW = Val(parts(0))
        If baseW > 0 Then
            fattore = w / baseW
        Else
            fattore = 1
        End If
    Else
        fattore = 1
    End If

    ' --- colori barra ---
    rStart = 180: gStart = 255: bStart = 180
    rEnd = 0: gEnd = 180: bEnd = 0

    t = percentuale / 100#
    r = rStart + (rEnd - rStart) * t
    G = gStart + (gEnd - gStart) * t
    b = bStart + (bEnd - bStart) * t

    ' --- pulizia e riempimento ---
    pic.Cls
    pic.FillStyle = vbSolid
    pic.FillColor = RGB(r, G, b)
    pic.Line (0, 0)-(wFill, h), pic.FillColor, BF

    ' --- font percentuale ---
    pic.Font.Bold = True
    pic.Font.Size = 12 * fattore
    If pic.Font.Size < 12 Then pic.Font.Size = 12
    If pic.Font.Size > 48 Then pic.Font.Size = 48

    ' --- testo percentuale ---
    testo = CStr(percentuale) & "%"

    pic.ForeColor = vbBlack
    pic.CurrentX = (w / 2) - (pic.TextWidth(testo) / 2)
    pic.CurrentY = (h / 2) - (pic.TextHeight(testo) / 2)
    pic.Print testo

    pic.Refresh

    ' --- testo sotto ---
    lbl.Caption = "Calibrazione in corso...  Tempo stimato: " & tempoStimato
    lbl.Refresh

    ' --- salva stato per il resize ---
    UltimaPercentuale = percentuale
    UltimoTempoStimato = tempoStimato

End Sub

Public Sub CalibraModelloCompleto()

    ' --- NON RICALIBRARE SE GIÀ FATTO ---
    If ModelloGiaCalibrato Then Exit Sub

    ' Variabili per il calcolo del tempo stabili con GetTickCount
    Dim startTick As Long
    Dim elapsedSeconds As Single
    Dim tempoStimato As Double
    Dim nuovoTempo As Double

    Dim iter As Long
    Dim maxIter As Long
    maxIter = 2000

    Dim bestErr As Double
    Dim errCalcolato As Double
    Dim p As ParametriModello
    Dim best As ParametriModello

    bestErr = 1E+30

    ' --- INIZIALIZZA MODELLO UNA SOLA VOLTA ---
    ModuloPronostico.InizializzaParametri
    ModuloPronostico.CaricaEAnalizzaRisultati

    ' Memorizza il tick di partenza in millisecondi
    startTick = GetTickCount()
    tempoStimato = 0

    For iter = 1 To maxIter

        ' --- INTERCETTA X ALL'INIZIO DEL CICLO (Chiusura Istantanea 1° Click) ---
        If EsciSubito Then
            InCalibrazione = False
            ChiusuraInCorso = True
            Call ModuloPronostico.TerminaApplicazione
            Call PostMessage(Me.hwnd, WM_CLOSE, 0&, 0&)
            Exit Sub
        End If

        ' --- GENERA PARAMETRI E VALUTA ---
        p = ModuloPronostico.GeneraParametriRandom()
        ModuloPronostico.ApplicaParametri p
        errCalcolato = ModuloPronostico.ValutaModello()

        If errCalcolato < bestErr Then
            bestErr = errCalcolato
            best = p
        End If

        ' --- CALCOLO TEMPO TRASCORSO (Convertito in secondi) ---
        elapsedSeconds = CSng(GetTickCount() - startTick) / 1000!

        If iter > 10 Then
            nuovoTempo = (elapsedSeconds / iter) * (maxIter - iter)
        Else
            nuovoTempo = 0
        End If

        tempoStimato = tempoStimato * 0.8 + nuovoTempo * 0.2

        ' --- AGGIORNAMENTO BARRA E UI (PASSO 10 - VELOCE) ---
        If (iter Mod 10 = 0) Or (iter = maxIter) Then

            Dim secTot As Long
            Dim min As Long
            Dim sec As Long
            Dim testoTempo As String

            secTot = CLng(tempoStimato)
            min = secTot \ 60
            sec = secTot Mod 60

            If min > 0 Then
                testoTempo = min & " min " & sec & " sec"
            Else
                testoTempo = sec & " sec"
            End If

            ' Aggiorna visivamente la barra
            DisegnaProgress _
                CInt(iter * 100 / maxIter), _
                testoTempo, _
                picProgress, _
                lblProgress

            ' Cede il controllo al sistema operativo
            DoEvents

            ' --- INTERCETTA X SUBITO DOPO IL DOEVENTS (Chiusura Istantanea 1° Click) ---
            If EsciSubito Then
                InCalibrazione = False
                ChiusuraInCorso = True
                Call ModuloPronostico.TerminaApplicazione
                Call PostMessage(Me.hwnd, WM_CLOSE, 0&, 0&)
                Exit Sub
            End If

        End If

    Next iter

    ' --- APPLICA I MIGLIORI PARAMETRI ---
    ModuloPronostico.ApplicaParametri best

    ' --- SALVA PARAMETRI CALIBRATI ---
    ModuloPronostico.ParametriCalibrati = best
    ModuloPronostico.ParametriCalibratiValidi = True

    ' --- SEGNA CHE IL MODELLO È CALIBRATO ---
    ModelloGiaCalibrato = True

End Sub


Private Sub Form_Resize()

    On Error GoTo GestioneErrore
    If Me.WindowState = vbMinimized Then Exit Sub

    Static InResize As Boolean
    If InResize Then Exit Sub
    InResize = True

    ' =====================================================
    ' SE SIAMO IN CALIBRAZIONE: NON TOCCO LA UI,
    ' CENTRO SOLO LA BARRA DI PROGRESSO E ESCO
    ' =====================================================
If InCalibrazione Then

    ' NON toccare WindowState, mai.
    ' Centra solo la barra e basta.

    If picProgress.Visible Then
        picProgress.Left = (Me.ScaleWidth - picProgress.Width) \ 2
        picProgress.Top = (Me.ScaleHeight \ 2) - (picProgress.Height \ 2)

        picLabelBox.Left = (Me.ScaleWidth - picLabelBox.Width) \ 2
        picLabelBox.Top = picProgress.Top + picProgress.Height + 80

        lblProgress.Left = (picLabelBox.ScaleWidth - lblProgress.Width) \ 2
        lblProgress.Top = (picLabelBox.ScaleHeight - lblProgress.Height) \ 2
    End If

    InResize = False
    Exit Sub
End If


    Const WM_SETREDRAW As Long = &HB
    Dim FattoreScala As Single
    Dim NuovaFontSizeLV As Single
    Dim NuovaFontSizeTitolo As Single
    Dim NuovaFontSizeSchedina As Single
    Dim Col As ColumnHeader
    Dim i As Integer

    ' Blocca ridisegno ListView
    SendMessage lvClassifica.hwnd, WM_SETREDRAW, 0, 0
    SendMessage lvPartite.hwnd, WM_SETREDRAW, 0, 0
    'SendMessage lvPronostico.hwnd, WM_SETREDRAW, 0, 0

    ' Base scaling
    If BaseWidth = 0 Then BaseWidth = Me.ScaleWidth

    ' Fattore scala
    FattoreScala = Me.ScaleWidth / BaseWidth
    If FattoreScala < SCALA_MINIMA Then FattoreScala = SCALA_MINIMA

    ' ============================
    ' SCALING BARRA + LABEL
    ' ============================
    Dim pW As Long, pH As Long
    Dim bW As Long, bH As Long
    Dim lW As Long, lH As Long
    Dim parts() As String

    parts = Split(picProgress.Tag, "|")
    pW = CLng(parts(0)): pH = CLng(parts(1))

    parts = Split(picLabelBox.Tag, "|")
    bW = CLng(parts(0)): bH = CLng(parts(1))

    parts = Split(lblProgress.Tag, "|")
    lW = CLng(parts(0)): lH = CLng(parts(1))

    ' SCALING SERIO
    picProgress.Width = CLng(pW * FattoreScala)
    picProgress.Height = CLng(pH * FattoreScala)

    picLabelBox.Width = CLng(bW * FattoreScala)
    picLabelBox.Height = CLng(bH * FattoreScala)

    lblProgress.Width = CLng(lW * FattoreScala)
    lblProgress.Height = CLng(lH * FattoreScala)

    ' FONT SERIO (PIÙ GRANDE)
    With lblProgress.Font
        .Bold = True
        .Size = 10 * FattoreScala
        If .Size < 12 Then .Size = 12
        If .Size > 48 Then .Size = 48
    End With

    lblProgress.AutoSize = False
    lblProgress.WordWrap = False

    ' ============================
    ' FONT ALTRI OGGETTI
    ' ============================
    NuovaFontSizeTitolo = BaseFontSizeTitolo * FattoreScala
    If NuovaFontSizeTitolo < 10 Then NuovaFontSizeTitolo = 10

    NuovaFontSizeLV = BaseFontSizeLV * FattoreScala
    If NuovaFontSizeLV < 8 Then NuovaFontSizeLV = 8

    NuovaFontSizeSchedina = 10 * FattoreScala
    If NuovaFontSizeSchedina < 8 Then NuovaFontSizeSchedina = 8

    If Abs(NuovaFontSizeTitolo - LastScaledFontSizeTitolo) >= 0.5 Then
        lblTitolo.Font.Size = NuovaFontSizeTitolo
        LastScaledFontSizeTitolo = NuovaFontSizeTitolo
    End If

    If Abs(NuovaFontSizeLV - LastScaledFontSizeLV) >= 0.5 Then
        lvClassifica.Font.Size = NuovaFontSizeLV
        lvPartite.Font.Size = NuovaFontSizeLV
        lvPronostico.Font.Size = NuovaFontSizeLV
        LastScaledFontSizeLV = NuovaFontSizeLV
    End If

    If Abs(NuovaFontSizeSchedina - LastScaledFontSizeSchedina) >= 0.5 Then
        lblSchedina.Font.Size = NuovaFontSizeSchedina
        LastScaledFontSizeSchedina = NuovaFontSizeSchedina
    End If

    ' ============================
    ' POSIZIONAMENTO TITOLO
    ' ============================
    lblTitolo.Width = Me.ScaleWidth - 2 * MARGINE_VERTICALE
    lblTitolo.Left = (Me.ScaleWidth - lblTitolo.Width) / 2
    lblTitolo.Height = ALTEZZA_TITOLO * FattoreScala
    lblTitolo.Top = MARGINE_VERTICALE * FattoreScala

    ' ============================
    ' CALCOLI SPAZI
    ' ============================
    Dim TopClassifica As Long: TopClassifica = lblTitolo.Top + lblTitolo.Height + MARGINE_VERTICALE * 2
    Dim AltezzaDisponibileTotale As Long: AltezzaDisponibileTotale = Me.ScaleHeight - TopClassifica - MARGINE_VERTICALE

    Dim AltezzaLVClassifica As Long: AltezzaLVClassifica = CLng(AltezzaDisponibileTotale * 0.37)
    Dim AltezzaLVPartiteSchedina As Long: AltezzaLVPartiteSchedina = CLng(AltezzaDisponibileTotale * 0.3)
    Dim AltezzaLVPronostico As Long: AltezzaLVPronostico = AltezzaDisponibileTotale - AltezzaLVClassifica - AltezzaLVPartiteSchedina - 2 * MARGINE_VERTICALE

    ' ============================
    ' LVCLASSIFICA
    ' ============================
    With lvClassifica
        .Top = TopClassifica
        .Width = Me.ScaleWidth - 2 * MARGINE_VERTICALE
        .Left = (Me.ScaleWidth - .Width) / 2
        .Height = AltezzaLVClassifica

        Dim LarghezzaListView As Long: LarghezzaListView = .Width
        Dim LarghezzaColFisseScalata As Long
        Dim LarghezzaResidua As Long

        LarghezzaColFisseScalata = 0
        For i = 1 To .ColumnHeaders.Count
            If i <> INDICE_COLONNA_ELASTICA_CLASS Then
                LarghezzaColFisseScalata = LarghezzaColFisseScalata + CLng(CDbl(.ColumnHeaders(i).Tag) * FattoreScala)
            End If
        Next i

        LarghezzaResidua = LarghezzaListView - LarghezzaColFisseScalata - LARGHEZZA_SCROLLBAR_VERTICALE
        If LarghezzaResidua < CLng(LARGHEZZA_MINIMA_SQUADRA * FattoreScala) Then LarghezzaResidua = CLng(LARGHEZZA_MINIMA_SQUADRA * FattoreScala)

        For i = 1 To .ColumnHeaders.Count
            Set Col = .ColumnHeaders(i)
            If i = INDICE_COLONNA_ELASTICA_CLASS Then
                Col.Width = LarghezzaResidua
            Else
                Col.Width = CLng(CDbl(Col.Tag) * FattoreScala)
            End If
        Next i
    End With

    ' ============================
    ' LVPARTITE + SCHEDINA
    ' ============================
    Dim TopRigaIntermedia As Long: TopRigaIntermedia = lvClassifica.Top + lvClassifica.Height + MARGINE_VERTICALE
    Dim LarghezzaDisponibileRiga As Long: LarghezzaDisponibileRiga = lvClassifica.Width

    Dim LarghezzaLVPartite As Long: LarghezzaLVPartite = CLng(LarghezzaDisponibileRiga * 0.6)
    Dim LarghezzaLBSchedina As Long: LarghezzaLBSchedina = LarghezzaDisponibileRiga - LarghezzaLVPartite - MARGINE_VERTICALE

    With lvPartite
        .Top = TopRigaIntermedia
        .Left = lvClassifica.Left
        .Width = LarghezzaLVPartite
        .Height = AltezzaLVPartiteSchedina

        LarghezzaColFisseScalata = 0
        For i = 1 To .ColumnHeaders.Count
            If i <> INDICE_COLONNA_ELASTICA_PARTITE Then
                LarghezzaColFisseScalata = LarghezzaColFisseScalata + CLng(CDbl(.ColumnHeaders(i).Tag) * FattoreScala)
            End If
        Next i

        LarghezzaResidua = .Width - LarghezzaColFisseScalata - LARGHEZZA_SCROLLBAR_VERTICALE
        If LarghezzaResidua < CLng(1000 * FattoreScala) Then LarghezzaResidua = CLng(1000 * FattoreScala)

        For i = 1 To .ColumnHeaders.Count
            Set Col = .ColumnHeaders(i)
            If i = INDICE_COLONNA_ELASTICA_PARTITE Then
                Col.Width = LarghezzaResidua
            Else
                Col.Width = CLng(CDbl(Col.Tag) * FattoreScala)
            End If
        Next i
    End With

    With lblSchedina
        .Top = TopRigaIntermedia
        .Left = lvPartite.Left + lvPartite.Width + MARGINE_VERTICALE
        .Width = LarghezzaLBSchedina
        .Height = AltezzaLVPartiteSchedina
    End With

    ' ----------------------------------------------------------------------
    ' LVPRONOSTICO (AUTO-FIT CON BORDO AGGANCIATO IN VISTA E STABILITÀ BACK)
    ' ----------------------------------------------------------------------
    Dim TopPronostico As Long: TopPronostico = TopRigaIntermedia + AltezzaLVPartiteSchedina + MARGINE_VERTICALE

    With lvPronostico
        .Top = TopPronostico
        .Left = lvClassifica.Left
        .Height = AltezzaLVPronostico
        
        ' === CORREZIONE STRUTTURALE ALLA LARGHEZZA DELLA GRIGLIA ===
        ' Sottraiamo 90 twips alla larghezza totale del controllo rispetto alla classifica.
        ' Questo costringe la ListView a fermarsi un millimetro prima, lasciando scoperto
        ' il pezzetto di riga finale del bordo della colonna come volevi tu.
        .Width = lvClassifica.Width - 90

        Dim LarghezzaPronostico As Long: LarghezzaPronostico = .Width
        Dim LarghezzaFissaScalata_P As Long: LarghezzaFissaScalata_P = 0

        ' 1) Esegue il calcolo originale saltando la colonna elastica (la 4, Motivazione)
        For i = 1 To .ColumnHeaders.Count
            If i <> INDICE_COLONNA_ELASTICA_PRONOSTICO Then
                LarghezzaFissaScalata_P = LarghezzaFissaScalata_P + CLng(CDbl(.ColumnHeaders(i).Tag) * FattoreScala)
            End If
        Next i

        LarghezzaResidua = LarghezzaPronostico - LarghezzaFissaScalata_P - LARGHEZZA_SCROLLBAR_VERTICALE
        If LarghezzaResidua < CLng(3000 * FattoreScala) Then LarghezzaResidua = CLng(3000 * FattoreScala)

        ' 2) Ridimensiona stabilmente le colonne
        For i = 1 To 5
            Set Col = .ColumnHeaders(i)
            If i = INDICE_COLONNA_ELASTICA_PRONOSTICO Then
                Col.Width = LarghezzaResidua
            Else
                Col.Width = CLng(CDbl(Col.Tag) * FattoreScala)
            End If
        Next i

        ' 3) RE-INSERIMENTO FORZATURA API ELASTICA SULL'ULTIMA COLONNA (Gol/NoGol)
        '    Dichiarate localmente per sicurezza per evitare errori se le hai tolte in cima
        Const LVM_FIRST As Long = &H1000
        Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
        
        ' Forza l'aggiornamento istantaneo dell'ultima colonna in entrambi gli stati della finestra
        Call SendMessage(.hwnd, LVM_FIRST + 30, 4, ByVal LVSCW_AUTOSIZE_USEHEADER)
        
        ' 4) RIPRISTINO DELLA CENTRATURA PERSA DALL'API
        .ColumnHeaders(5).Alignment = lvwColumnCenter
    End With

    ' ============================
    ' POSIZIONAMENTO BARRA + LABEL
    ' ============================
    If picProgress.Visible Then
        picProgress.Left = (Me.ScaleWidth - picProgress.Width) \ 2
        picProgress.Top = (Me.ScaleHeight \ 2) - (picProgress.Height \ 2)

        picLabelBox.Left = (Me.ScaleWidth - picLabelBox.Width) \ 2
        picLabelBox.Top = picProgress.Top + picProgress.Height + CLng(80 * FattoreScala)

        lblProgress.Left = (picLabelBox.ScaleWidth - lblProgress.Width) \ 2
        lblProgress.Top = (picLabelBox.ScaleHeight - lblProgress.Height) \ 2

        ' RIDISEGNA LA BARRA VERDE DOPO IL RESIZE
        If UltimaPercentuale > 0 Then
            Call DisegnaProgress(UltimaPercentuale, UltimoTempoStimato, picProgress, lblProgress)
        End If
    End If

    ' Riabilita ridisegno
    SendMessage lvClassifica.hwnd, WM_SETREDRAW, 1, 0
    SendMessage lvPartite.hwnd, WM_SETREDRAW, 1, 0
    SendMessage lvPronostico.hwnd, WM_SETREDRAW, 1, 0
    Me.Refresh

    ' Debounce
    tmrRefresh.Enabled = False
    tmrRefresh.Interval = 1
    tmrRefresh.Enabled = True

    InResize = False
    Exit Sub

GestioneErrore:
    InResize = False
    SendMessage lvClassifica.hwnd, WM_SETREDRAW, 1, 0
    SendMessage lvPartite.hwnd, WM_SETREDRAW, 1, 0
    SendMessage lvPronostico.hwnd, WM_SETREDRAW, 1, 0
    Resume Next

End Sub


' Blocchi BeforeLabelEdit
Private Sub lvClassifica_BeforeLabelEdit(Cancel As Integer): Cancel = True: End Sub
Private Sub lvPartite_BeforeLabelEdit(Cancel As Integer): Cancel = True: End Sub
Private Sub lvPronostico_BeforeLabelEdit(Cancel As Integer): Cancel = True: End Sub

Private Sub lvPartite_DblClick()
    Dim itm As ListItem
    Dim Ris As String
    Dim Casa As String, Ospite As String
    Dim Giornata As String
    Dim NuovoRis As String
    
If InCalibrazione Then Exit Sub

    If lvPartite.SelectedItem Is Nothing Then Exit Sub
    Set itm = lvPartite.SelectedItem

    Dim SelectedIndex As Long
    SelectedIndex = itm.Index

    Giornata = Trim$(itm.Text)
    Casa = Trim$(itm.SubItems(1))
    Ris = itm.SubItems(2)
    Ospite = Trim$(itm.SubItems(3))

    NuovoRis = InputBox( _
        "Modifica risultato per " & Casa & " - " & Ospite & vbCrLf & _
        "Inserisci nel formato: golCasa-golOspite (es. 2-1). Lascia vuoto per '-'.", _
        "Modifica risultato - Giornata " & Giornata, Ris)

    If StrPtr(NuovoRis) = 0 Then Exit Sub
    If Len(Trim$(NuovoRis)) = 0 Then NuovoRis = "-"

    If NuovoRis <> "-" Then
        If InStr(NuovoRis, "-") = 0 Then
            MsgBox "Formato non valido. Usa ad es. 2-1 oppure '-'.", vbExclamation
            Exit Sub
        Else
            Dim parts() As String
            parts = Split(NuovoRis, "-")
            If UBound(parts) <> 1 Or _
               Not (IsNumeric(Trim$(parts(0))) And IsNumeric(Trim$(parts(1)))) Then
                MsgBox "Formato non valido. Usa ad es. 2-1 oppure '-'.", vbExclamation
                Exit Sub
            End If
        End If
    End If

    itm.SubItems(2) = NuovoRis

    ModuloPronostico.AggiornaRisultatoNelFile _
        App.Path & "\Risultati.txt", CLng(Giornata), Casa, Ospite, NuovoRis

    ModuloPronostico.CaricaEAnalizzaRisultati

    If ModelloGiaCalibrato Then
        ModuloPronostico.ApplicaParametri ModuloPronostico.ParametriCalibrati
    End If

    OrdinaClassifica
    AggiornaListView
    AggiornaPronostici

    tmrRefresh.Enabled = False
    tmrRefresh.Interval = 5
    tmrRefresh.Enabled = True

    On Error Resume Next
    If SelectedIndex > 0 And SelectedIndex <= lvPartite.ListItems.Count Then
        lvPartite.ListItems(SelectedIndex).EnsureVisible
        lvPartite.ListItems(SelectedIndex).Selected = True
    Else
        If lvPartite.ListItems.Count > 0 Then
            lvPartite.ListItems(1).EnsureVisible
            lvPartite.ListItems(1).Selected = True
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub RimuoviPronosticoDaSchedina(ByVal Casa As String, ByVal Ospite As String)
    Dim i As Long
    Dim itm As ListItem

    ' ?? Rimuove dalla ListView dei pronostici
    For i = lvPronostico.ListItems.Count To 1 Step -1
        Set itm = lvPronostico.ListItems(i)
        If itm.Text = Casa And itm.SubItems(1) = Ospite Then
            lvPronostico.ListItems.Remove i
            Exit For
        End If
    Next i

' Rimuove dalla schedina (lblSchedina)
Dim Righe() As String
Dim nuova As String
Dim r As Long

Righe = Split(lblSchedina.Caption, vbCrLf)

For r = LBound(Righe) To UBound(Righe)
    ' salta SOLO la riga che contiene sia Casa che Ospite
    If Not (InStr(1, Righe(r), Casa, vbTextCompare) > 0 And _
            InStr(1, Righe(r), " - ", vbTextCompare) > 0 And _
            InStr(1, Righe(r), Ospite, vbTextCompare) > 0) Then

        nuova = nuova & Righe(r) & vbCrLf
    End If
Next r

lblSchedina.Caption = nuova
End Sub

Public Sub CancellaPronosticoPartita(ByVal G As Long, ByVal Casa As String, ByVal Ospite As String)
    Dim i As Long
    For i = 0 To ContaPartite - 1
        With PartiteGiocate(i)
            If .Giornata = G And .Casa = Casa And .Ospite = Ospite Then
                .GoalCasa = -1
                .GoalOspite = -1
                Exit For
            End If
        End With
    Next i
End Sub

Private Sub AggiornaSoloDatiELista()
    ModuloPronostico.CaricaEAnalizzaRisultati
    OrdinaClassifica

    AggiornaListView   ' ora è veloce e senza flicker

    Dim nextG As Long
    nextG = ProssimaGiornataDaFile()

    SelezionaGiornataInListView nextG
End Sub

Private Sub AggiornaTuttoDopoModifica()

    ' Ricarica tutto
    CaricaEAnalizzaRisultati
    OrdinaClassifica
    AggiornaListView
    AggiornaPronostici

    ' Trova la prossima giornata dal file
    Dim nextG As Long
    nextG = ProssimaGiornataDaFile()

    ' Seleziona e porta in cima
    SelezionaGiornataInListView nextG
End Sub

Private Sub SelezionaGiornataInListView(ByVal Giornata As Long)
    Dim i As Long

    On Error Resume Next

    For i = 1 To lvPartite.ListItems.Count
        If CLng(lvPartite.ListItems(i).Text) = Giornata Then
            lvPartite.ListItems(i).Selected = True
            lvPartite.ListItems(i).EnsureVisible
            Exit Sub
        End If
    Next

    ' fallback
    If lvPartite.ListItems.Count > 0 Then
        lvPartite.ListItems(1).Selected = True
        lvPartite.ListItems(1).EnsureVisible
    End If

    On Error GoTo 0
End Sub

Private Function ProssimaGiornataDaFile() As Long
    Dim Righe As Variant
    Dim r As Variant
    Dim campi() As String

    Righe = LeggiTestoDaFile("Risultati.txt")
    If Not IsArray(Righe) Then
        ProssimaGiornataDaFile = 1
        Exit Function
    End If

    For Each r In Righe
        If Len(Trim$(r)) > 0 And Left$(r, 1) <> ";" And Left$(r, 3) <> "***" Then
            campi = Split(r, ";")
            If UBound(campi) >= 4 Then
                If Trim$(campi(2)) = "-" Or Trim$(campi(4)) = "-" Then
                    ProssimaGiornataDaFile = CLng(Trim$(campi(0)))
                    Exit Function
                End If
            End If
        End If
    Next

    ProssimaGiornataDaFile = 1
End Function

Public Function InIDE() As Boolean
    Debug.Assert MakeTrue(InIDE)
End Function

Private Function MakeTrue(b As Boolean) As Boolean
    MakeTrue = True
End Function

Private Sub tmrRefresh_Timer()
    tmrRefresh.Enabled = False
    AggiornaListView
End Sub


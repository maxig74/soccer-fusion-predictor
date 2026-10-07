VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PronosticoPartita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ************************************************************
' * VARIABILI MEMBRO (PRIVATE)
' ************************************************************

Private m_Casa As String
Private m_Ospite As String
Private m_E_GoalCasa As Single
Private m_E_GoalOspite As Single
Private m_Prob_1 As Single
Private m_Prob_X As Single
Private m_Prob_2 As Single
Private m_Prob_OVER25 As Single
Private m_Prob_UNDER25 As Single
Private m_Prob_GG As Single
Private m_Prob_NG As Single
Private m_PronosticoGG_NG As String
Private m_ConfidenzaGG_NG As Single
Private m_PunteggioEsattoCasa As Long
Private m_PunteggioEsattoOspite As Long
Private m_ProbabilitaRE As Single
Private m_PronosticoComplesso As String
Private m_Confidenza As Single
Private m_Motivazione As String
Private m_RisultatoSorpresa As Boolean

' Nuovi indicatori avanzati (FUSION)
Private m_IndiceRischio As Single
Private m_MatchAltaVarianza As Boolean
Private m_CoeffSorpresa As Single
Private m_ForzaFusionCasa As Single
Private m_ForzaFusionOspite As Single
Private m_DeltaELO As Single
Private m_DeltaForma As Single
Private m_DeltaExpPoints As Single

' Matrice delle probabilità dei risultati esatti
' Allineata a MAX_GOALS_SUM = 6 nel ModuloPronostico
Private m_ProbabilitaPunteggio(0 To 6, 0 To 6) As Double

' ************************************************************
' * PROPRIETÀ GET/LET
' ************************************************************

Public Property Get Casa() As String
    Casa = m_Casa
End Property
Public Property Let Casa(ByVal vNewValue As String)
    m_Casa = vNewValue
End Property

Public Property Get Ospite() As String
    Ospite = m_Ospite
End Property
Public Property Let Ospite(ByVal vNewValue As String)
    m_Ospite = vNewValue
End Property

Public Property Get E_GoalCasa() As Single
    E_GoalCasa = m_E_GoalCasa
End Property
Public Property Let E_GoalCasa(ByVal vNewValue As Single)
    m_E_GoalCasa = vNewValue
End Property

Public Property Get E_GoalOspite() As Single
    E_GoalOspite = m_E_GoalOspite
End Property
Public Property Let E_GoalOspite(ByVal vNewValue As Single)
    m_E_GoalOspite = vNewValue
End Property

Public Property Get Prob_1() As Single
    Prob_1 = m_Prob_1
End Property
Public Property Let Prob_1(ByVal vNewValue As Single)
    m_Prob_1 = vNewValue
End Property

Public Property Get Prob_X() As Single
    Prob_X = m_Prob_X
End Property
Public Property Let Prob_X(ByVal vNewValue As Single)
    m_Prob_X = vNewValue
End Property

Public Property Get Prob_2() As Single
    Prob_2 = m_Prob_2
End Property
Public Property Let Prob_2(ByVal vNewValue As Single)
    m_Prob_2 = vNewValue
End Property

Public Property Get Prob_OVER25() As Single
    Prob_OVER25 = m_Prob_OVER25
End Property
Public Property Let Prob_OVER25(ByVal vNewValue As Single)
    m_Prob_OVER25 = vNewValue
End Property

Public Property Get Prob_UNDER25() As Single
    Prob_UNDER25 = m_Prob_UNDER25
End Property
Public Property Let Prob_UNDER25(ByVal vNewValue As Single)
    m_Prob_UNDER25 = vNewValue
End Property

Public Property Get Prob_GG() As Single
    Prob_GG = m_Prob_GG
End Property
Public Property Let Prob_GG(ByVal vNewValue As Single)
    m_Prob_GG = vNewValue
End Property

Public Property Get Prob_NG() As Single
    Prob_NG = m_Prob_NG
End Property
Public Property Let Prob_NG(ByVal vNewValue As Single)
    m_Prob_NG = vNewValue
End Property

Public Property Get PronosticoGG_NG() As String
    PronosticoGG_NG = m_PronosticoGG_NG
End Property
Public Property Let PronosticoGG_NG(ByVal vNewValue As String)
    m_PronosticoGG_NG = vNewValue
End Property

Public Property Get ConfidenzaGG_NG() As Single
    ConfidenzaGG_NG = m_ConfidenzaGG_NG
End Property
Public Property Let ConfidenzaGG_NG(ByVal vNewValue As Single)
    m_ConfidenzaGG_NG = vNewValue
End Property

Public Property Get PunteggioEsattoCasa() As Long
    PunteggioEsattoCasa = m_PunteggioEsattoCasa
End Property
Public Property Let PunteggioEsattoCasa(ByVal vNewValue As Long)
    m_PunteggioEsattoCasa = vNewValue
End Property

Public Property Get PunteggioEsattoOspite() As Long
    PunteggioEsattoOspite = m_PunteggioEsattoOspite
End Property
Public Property Let PunteggioEsattoOspite(ByVal vNewValue As Long)
    m_PunteggioEsattoOspite = vNewValue
End Property

Public Property Get ProbabilitaRE() As Single
    ProbabilitaRE = m_ProbabilitaRE
End Property
Public Property Let ProbabilitaRE(ByVal vNewValue As Single)
    m_ProbabilitaRE = vNewValue
End Property

Public Property Get PronosticoComplesso() As String
    PronosticoComplesso = m_PronosticoComplesso
End Property
Public Property Let PronosticoComplesso(ByVal vNewValue As String)
    m_PronosticoComplesso = vNewValue
End Property

Public Property Get Confidenza() As Single
    Confidenza = m_Confidenza
End Property
Public Property Let Confidenza(ByVal vNewValue As Single)
    m_Confidenza = vNewValue
End Property

Public Property Get Motivazione() As String
    Motivazione = m_Motivazione
End Property
Public Property Let Motivazione(ByVal vNewValue As String)
    m_Motivazione = vNewValue
End Property

Public Property Get RisultatoSorpresa() As Boolean
    RisultatoSorpresa = m_RisultatoSorpresa
End Property
Public Property Let RisultatoSorpresa(ByVal vNewValue As Boolean)
    m_RisultatoSorpresa = vNewValue
End Property

' --- Proprietà avanzate FUSION ---

Public Property Get IndiceRischio() As Single
    IndiceRischio = m_IndiceRischio
End Property
Public Property Let IndiceRischio(ByVal vNewValue As Single)
    m_IndiceRischio = vNewValue
End Property

Public Property Get MatchAltaVarianza() As Boolean
    MatchAltaVarianza = m_MatchAltaVarianza
End Property
Public Property Let MatchAltaVarianza(ByVal vNewValue As Boolean)
    m_MatchAltaVarianza = vNewValue
End Property

Public Property Get CoeffSorpresa() As Single
    CoeffSorpresa = m_CoeffSorpresa
End Property
Public Property Let CoeffSorpresa(ByVal vNewValue As Single)
    m_CoeffSorpresa = vNewValue
End Property

Public Property Get ForzaFusionCasa() As Single
    ForzaFusionCasa = m_ForzaFusionCasa
End Property
Public Property Let ForzaFusionCasa(ByVal vNewValue As Single)
    m_ForzaFusionCasa = vNewValue
End Property

Public Property Get ForzaFusionOspite() As Single
    ForzaFusionOspite = m_ForzaFusionOspite
End Property
Public Property Let ForzaFusionOspite(ByVal vNewValue As Single)
    m_ForzaFusionOspite = vNewValue
End Property

Public Property Get DeltaELO() As Single
    DeltaELO = m_DeltaELO
End Property
Public Property Let DeltaELO(ByVal vNewValue As Single)
    m_DeltaELO = vNewValue
End Property

Public Property Get DeltaForma() As Single
    DeltaForma = m_DeltaForma
End Property
Public Property Let DeltaForma(ByVal vNewValue As Single)
    m_DeltaForma = vNewValue
End Property

Public Property Get DeltaExpPoints() As Single
    DeltaExpPoints = m_DeltaExpPoints
End Property
Public Property Let DeltaExpPoints(ByVal vNewValue As Single)
    m_DeltaExpPoints = vNewValue
End Property

' --- Accessori per la Matrice ProbabilitaPunteggio ---

Public Property Get ProbabilitaPunteggio(ByVal Row As Long, ByVal Col As Long) As Double
    Dim rLow As Long, rHigh As Long, cLow As Long, cHigh As Long
    rLow = LBound(m_ProbabilitaPunteggio, 1): rHigh = UBound(m_ProbabilitaPunteggio, 1)
    cLow = LBound(m_ProbabilitaPunteggio, 2): cHigh = UBound(m_ProbabilitaPunteggio, 2)
    If Row < rLow Or Row > rHigh Or Col < cLow Or Col > cHigh Then
        ProbabilitaPunteggio = 0#
    Else
        ProbabilitaPunteggio = m_ProbabilitaPunteggio(Row, Col)
    End If
End Property

Public Property Let ProbabilitaPunteggio(ByVal Row As Long, ByVal Col As Long, ByVal RHS As Double)
    Dim rLow As Long, rHigh As Long, cLow As Long, cHigh As Long
    rLow = LBound(m_ProbabilitaPunteggio, 1): rHigh = UBound(m_ProbabilitaPunteggio, 1)
    cLow = LBound(m_ProbabilitaPunteggio, 2): cHigh = UBound(m_ProbabilitaPunteggio, 2)
    If Row < rLow Or Row > rHigh Or Col < cLow Or Col > cHigh Then Exit Property
    m_ProbabilitaPunteggio(Row, Col) = RHS
End Property

' ************************************************************
' * METODI
' ************************************************************

Private Sub Class_Initialize()
    Dim r As Long, c As Long
    
    m_E_GoalCasa = 0
    m_E_GoalOspite = 0
    m_PunteggioEsattoCasa = -1
    m_PunteggioEsattoOspite = -1
    m_PronosticoComplesso = "N/D"
    m_Motivazione = "Calcolo in attesa."
    m_ProbabilitaRE = 0
    m_Confidenza = 0
    m_Prob_1 = 0: m_Prob_X = 0: m_Prob_2 = 0
    m_Prob_OVER25 = 0: m_Prob_UNDER25 = 0
    m_Prob_GG = 0: m_Prob_NG = 0
    m_RisultatoSorpresa = False
    
    m_IndiceRischio = 0
    m_MatchAltaVarianza = False
    m_CoeffSorpresa = 0
    m_ForzaFusionCasa = 0
    m_ForzaFusionOspite = 0
    m_DeltaELO = 0
    m_DeltaForma = 0
    m_DeltaExpPoints = 0
    
    For r = LBound(m_ProbabilitaPunteggio, 1) To UBound(m_ProbabilitaPunteggio, 1)
        For c = LBound(m_ProbabilitaPunteggio, 2) To UBound(m_ProbabilitaPunteggio, 2)
            m_ProbabilitaPunteggio(r, c) = 0#
        Next c
    Next r
End Sub

Public Sub Reset()
    Class_Initialize
End Sub

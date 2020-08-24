VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtMRZ 
      Height          =   1695
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   10695
   End
   Begin VB.CommandButton Scan1 
      Caption         =   "처리1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function OpenPort Lib "GTF_PASSPORT_COMM.dll" () As Integer
Private Declare Function OpenPortByNumber Lib "GTF_PASSPORT_COMM.dll" (ByVal nPorNum As Integer) As Integer
Private Declare Function ClosePort Lib "GTF_PASSPORT_COMM.dll" () As Integer
Private Declare Function Scan Lib "GTF_PASSPORT_COMM.dll" () As Integer
Private Declare Function ReceiveData Lib "GTF_PASSPORT_COMM.dll" (ByVal nTimeOUt As Integer) As Integer
Private Declare Function GetMRZ1 Lib "GTF_PASSPORT_COMM.dll" (szMRZ1 As Byte) As Integer
Private Declare Function GetMRZ2 Lib "GTF_PASSPORT_COMM.dll" (szMRZ2 As Byte) As Integer
Private Declare Function GetPassportInfo Lib "GTF_PASSPORT_COMM.dll" (szPassportInfo As Byte) As Integer


Private Sub Form_Load()

    txtMRZ.Text = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim ret As Integer

    'ClosePort

End Sub


Private Sub Scan1_Click()
    Dim ret As Integer
    Dim i As Integer
    Dim retry As Integer
    
    Dim szPassportInfo(1024) As Byte
    
    txtMRZ.Text = ""
    
    'ret = OpenPort()
    ret = OpenPortByNumber(3)
    If ret = 1 Then
    
        ret = Scan()
        ret = ReceiveData(10)
        
        If ret = 1 Then
        
            ret = GetPassportInfo(szPassportInfo(0))
            
            For i = LBound(szPassportInfo) To UBound(szPassportInfo) - 1
                txtMRZ.Text = txtMRZ.Text & Chr(szPassportInfo(i))
            Next i
            
        ElseIf ret = 0 Then
            MsgBox ("Time-Out!")
        ElseIf ret < 0 Then
            MsgBox ("여권정보 오류!")
        End If
        
        ret = ClosePort()
    
    End If

End Sub



VERSION 5.00
Begin VB.Form frmService 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCheck 
      Interval        =   1000
      Left            =   4275
      Top             =   225
   End
   Begin VB.Timer tmrConfigre 
      Interval        =   1000
      Left            =   4140
      Top             =   2385
   End
   Begin VB.Label lblwait 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2250
      TabIndex        =   2
      Top             =   300
      Width           =   45
   End
   Begin VB.Image Image2 
      Height          =   9000
      Left            =   245
      Picture         =   "service.frx":0000
      Top             =   0
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   450
      Picture         =   "service.frx":042E
      Top             =   200
      Width           =   465
   End
   Begin VB.Label lblService 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IIS Services"
      Height          =   195
      Left            =   1005
      TabIndex        =   1
      Top             =   300
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EDA84D&
      Height          =   2040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Locator As SWbemLocator
Public services As SWbemServices
Public TimerCount
Public SvcList As Variant

Dim WithEvents eventSink As SWbemSink
Attribute eventSink.VB_VarHelpID = -1

Private Sub tmrCheck_Timer()

    TimerCount = TimerCount + 1

    If TimerCount >= 40 Then
        
        Select Case ServiceCommand

            Case 0
                MsgBox "Unable to stop IIS Services. Please contact your system administrator."
                ServiceCommand = 99

            Case 1
                MsgBox "Your setup has completed. However, one or more IIS Services can not be started. You may restart your server to automatýcally start IIS Services."
            
        End Select

        tmrCheck.Enabled = False
        Unload Me

    End If

End Sub

Private Sub eventSink_OnObjectReady(ByVal Object As WbemScripting.ISWbemObject, ByVal AsyncContext As WbemScripting.ISWbemNamedValueSet)

    Dim ServiceName
    Dim ServiceStatus

    ServiceName = Object.TargetInstance.Name
    ServiceStatus = Object.TargetInstance.State

    TimerCount = 0

End Sub

Public Sub LoadView()

    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    
    ' On Error Resume Next
        
    SavePointer = frmService.MousePointer
    frmService.MousePointer = vbHourglass
    frmService.Enabled = False
    
    eventSink.Cancel
    
    Set services = Locator.ConnectServer("127.0.0.1")
    services.ExecNotificationQueryAsync eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    
    frmService.Enabled = True
    frmService.MousePointer = SavePointer

End Sub

Public Sub Check()

    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    Dim item As String
    
    On Error Resume Next
        
    SavePointer = frmService.MousePointer
    frmService.MousePointer = vbHourglass
    frmService.Enabled = False
    
    eventSink.Cancel
    
    Set services = Locator.ConnectServer("127.0.0.1")
    services.ExecNotificationQueryAsync eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = services.ExecQuery("Select * From Win32_Service Where Name='" & SvcList(0) & "' or Name='" & SvcList(1) & "'  or Name='" & SvcList(2) & "'  or Name='" & SvcList(3) & "'  or Name='" & SvcList(4) & "'")
    
    For Each Object In Enumerator

        item = Object.State

        Select Case ServiceCommand

            Case 0

                If item <> "Stopped" Then
                
                    lblwait.Caption = lblService.Caption
                    lblService.Visible = False
                    lblwait.Visible = True
                    Call service

                End If

            Case 1

                If item <> "Running" Then

                    lblwait.Caption = lblService.Caption
                    lblService.Visible = False
                    lblwait.Visible = True
                        
                    Call service

                End If
        
        End Select

    Next

    frmService.Enabled = True
    frmService.MousePointer = SavePointer

End Sub

Private Sub Form_Load()

    Center Me
    
    Set Locator = New SWbemLocator
    Set eventSink = New SWbemSink
    ReDim SvcList(4)
    SvcList(0) = "W3SVC"
    SvcList(1) = "SMTPSVC"
    SvcList(2) = "MSFTPSVC"
    SvcList(3) = "NNTPSVC"
    SvcList(4) = "IISADMIN"
    lblwait.Visible = False
    lblwait.Top = lblService.Top
    lblwait.Left = lblService.Left
    
End Sub

Private Sub service()
    
    Dim ServiceObject As SWbemObject
    Dim ServiceName
    Dim I%
    On Error Resume Next

    If Err.Number = 0 Then

        Select Case ServiceCommand

            Case 0

                For I% = LBound(SvcList) To UBound(SvcList)

                    Set ServiceObject = services.Get("Win32_Service='" & SvcList(I%) & "'")
                    lblService.Caption = "Stoping " & ServiceObject.Description & " . . ."
                    ServiceObject.StopService
   
                    Set ServiceObject = Nothing
                    Timeout 1

                Next I%

            Case 1

                For I% = LBound(SvcList) To UBound(SvcList)

                    Set ServiceObject = services.Get("Win32_Service='" & SvcList(I%) & "'")
                    lblService.Caption = "Starting " & ServiceObject.Description & " . . ."
                    ServiceObject.StartService

                    Set ServiceObject = Nothing
                    Timeout 1

                Next I%

        End Select

        Timeout 3

    End If

    Check

End Sub

Private Sub tmrConfigre_Timer()

    LoadView
    service
    tmrConfigre.Enabled = False
    Unload Me

End Sub


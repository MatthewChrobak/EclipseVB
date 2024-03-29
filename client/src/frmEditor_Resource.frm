VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Resource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   815
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraQuest 
      Caption         =   "Quest Triggers"
      Height          =   1815
      Left            =   8520
      TabIndex        =   32
      Top             =   120
      Width           =   3375
      Begin VB.HScrollBar scrlQuestTask 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Value           =   1
         Width           =   3135
      End
      Begin VB.HScrollBar scrlQuestIndex 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   3135
      End
      Begin VB.HScrollBar scrlQuestType 
         Height          =   255
         Left            =   120
         Max             =   2
         TabIndex        =   33
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblQuestTask 
         Caption         =   "Quest Task: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label lblQuestIndex 
         Caption         =   "Quest Index: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblQuestType 
         Caption         =   "Quest Type: None"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   27
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   26
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   25
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resource Properties"
      Height          =   7575
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   7080
         Width           =   3975
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   28
         Top             =   6720
         Width           =   4815
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   1920
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   2640
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   22
         Top             =   2280
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3332
         Left            =   960
         List            =   "frmEditor_Resource.frx":3342
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Width           =   3975
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   2295
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   4815
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   8
         Top             =   4920
         Width           =   4815
      End
      Begin VB.HScrollBar scrlHealth 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   7
         Top             =   5520
         Width           =   4815
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   6
         Top             =   2280
         Width           =   2280
      End
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   5
         Top             =   6120
         Width           =   4815
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   6480
         Width           =   1260
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   24
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1470
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Item Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   1440
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Width           =   1530
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         Caption         =   "Health: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   5280
         Width           =   705
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Seconds): 0"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   5880
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource List"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearResource EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ResourceEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlReward.max = MAX_ITEMS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ResourceEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim AnimName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.value = 0 Then
        AnimName = "None"
    Else
        AnimName = Trim$(Animation(scrlAnimation.value).name)
    End If
    
    lblAnim.Caption = "Animation: " & AnimName
    Resource(EditorIndex).Animation = scrlAnimation.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlExhaustedPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblExhaustedPic.Caption = "Exhausted Image: " & scrlExhaustedPic.value
    EditorResource_BltSprite
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlExhaustedPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHealth_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblHealth.Caption = "Health: " & scrlHealth.value
    Resource(EditorIndex).health = scrlHealth.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlHealth_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNormalPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.value
    EditorResource_BltSprite
    Resource(EditorIndex).ResourceImage = scrlNormalPic.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNormalPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlQuestIndex_Change()
    lblQuestIndex.Caption = "Quest Index: " & scrlQuestIndex.value
    Resource(EditorIndex).QuestIndex = scrlQuestIndex.value
End Sub

Private Sub scrlQuestTask_Change()
    lblQuestTask.Caption = "Quest Task: " & scrlQuestTask.value
    Resource(EditorIndex).QuestTask = scrlQuestTask.value
End Sub

Private Sub scrlQuestType_Change()
    scrlQuestIndex.Visible = False
    lblQuestIndex.Visible = False
    scrlQuestTask.Visible = False
    lblQuestTask.Visible = False
    
    Select Case scrlQuestType.value
        Case 0
            lblQuestType.Caption = "Quest Type: None"
        Case 1
            lblQuestType.Caption = "Quest Type: Start Quest"
            scrlQuestIndex.Visible = True
            lblQuestIndex.Visible = True
        Case 2
            lblQuestType.Caption = "Quest Type: Advance Quest"
            scrlQuestIndex.Visible = True
            lblQuestIndex.Visible = True
            scrlQuestTask.Visible = True
            lblQuestTask.Visible = True
    End Select
    
    Resource(EditorIndex).QuestType = scrlQuestType.value
End Sub

Private Sub scrlRespawn_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRespawn.Caption = "Respawn Time (Seconds): " & scrlRespawn.value
    Resource(EditorIndex).RespawnTime = scrlRespawn.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRespawn_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlReward_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlReward.value > 0 Then
        lblReward.Caption = "Item Reward: " & Trim$(Item(scrlReward.value).name)
    Else
        lblReward.Caption = "Item Reward: None"
    End If
    
    Resource(EditorIndex).ItemReward = scrlReward.value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlReward_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTool_Change()
    Dim name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlTool.value
        Case 0: name = "None"
        Case 1: name = "Hatchet"
        Case 2: name = "Rod"
        Case 3: name = "Pickaxe"
    End Select

    lblTool.Caption = "Tool Required: " & name
    Resource(EditorIndex).ToolRequired = scrlTool.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage2_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage2_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Resource(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Resource(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

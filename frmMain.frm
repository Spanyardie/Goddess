VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "G.O.D.D.E.S.S"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   622
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8745
      Left            =   105
      TabIndex        =   1
      Top             =   45
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   15425
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Message Filter"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lstWatch"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lstScoreboard"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstMessages"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFile"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraMessages"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraWatch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraScoreboard"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Game Servers"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstData"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdPktClear"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdDump"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "FTP"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lstFTP"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame6"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Server Filter"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label11"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lstValidGameServers"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame7"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame7 
         Caption         =   "Game Servers"
         Height          =   7905
         Left            =   -68085
         TabIndex        =   90
         Top             =   705
         Width           =   4800
         Begin VB.CommandButton cmdRemoveValidGameServer 
            Caption         =   "Re&move Server"
            Height          =   330
            Left            =   3150
            TabIndex        =   99
            Top             =   4950
            Width           =   1320
         End
         Begin VB.CommandButton cmdAddValidGameServer 
            Caption         =   "Add Ser&ver"
            Height          =   330
            Left            =   3150
            TabIndex        =   97
            Top             =   3345
            Width           =   1320
         End
         Begin VB.TextBox txtValidServerPort 
            Height          =   315
            Left            =   165
            TabIndex        =   96
            Top             =   2775
            Width           =   1710
         End
         Begin VB.TextBox txtValidServerIP 
            Height          =   315
            Left            =   165
            TabIndex        =   95
            Top             =   1770
            Width           =   2835
         End
         Begin VB.TextBox txtValidServerName 
            Height          =   315
            Left            =   165
            TabIndex        =   94
            Top             =   795
            Width           =   4305
         End
         Begin VB.Label Label26 
            Caption         =   "Remove selected Game Server:"
            Height          =   225
            Left            =   240
            TabIndex        =   98
            Top             =   4500
            Width           =   2640
         End
         Begin VB.Line Line12 
            BorderColor     =   &H80000005&
            X1              =   210
            X2              =   4515
            Y1              =   4125
            Y2              =   4125
         End
         Begin VB.Line Line11 
            BorderColor     =   &H80000000&
            X1              =   210
            X2              =   4515
            Y1              =   4110
            Y2              =   4110
         End
         Begin VB.Label Label25 
            Caption         =   "Server Port:"
            Height          =   225
            Left            =   165
            TabIndex        =   93
            Top             =   2460
            Width           =   1290
         End
         Begin VB.Label Label24 
            Caption         =   "Server IP:"
            Height          =   285
            Left            =   165
            TabIndex        =   92
            Top             =   1470
            Width           =   990
         End
         Begin VB.Label Label23 
            Caption         =   "Server Name:"
            Height          =   270
            Left            =   165
            TabIndex        =   91
            Top             =   495
            Width           =   1230
         End
      End
      Begin MSComctlLib.ListView lstValidGameServers 
         Height          =   7770
         Left            =   -74865
         TabIndex        =   88
         Top             =   810
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   13705
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Server Name"
            Object.Width           =   6421
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Server IP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Server Port"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame6 
         Caption         =   "FTP Details:"
         Height          =   2520
         Left            =   7320
         TabIndex        =   78
         Top             =   5925
         Width           =   4335
         Begin VB.CommandButton cmdFTPPassword 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3525
            TabIndex        =   87
            Top             =   2085
            Width           =   705
         End
         Begin VB.CommandButton cmdFTPUsername 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3525
            TabIndex        =   86
            Top             =   1305
            Width           =   705
         End
         Begin VB.CommandButton cmdFTPHostName 
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3525
            TabIndex        =   85
            Top             =   525
            Width           =   705
         End
         Begin VB.TextBox txtFTPPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   180
            PasswordChar    =   "*"
            TabIndex        =   81
            Top             =   2070
            Width           =   3285
         End
         Begin VB.TextBox txtFTPUsername 
            Height          =   315
            Left            =   180
            TabIndex        =   80
            Top             =   1282
            Width           =   3285
         End
         Begin VB.TextBox txtFTPHostName 
            Height          =   315
            Left            =   180
            TabIndex        =   79
            Top             =   495
            Width           =   3285
         End
         Begin VB.Label Label16 
            Caption         =   "Password:"
            Height          =   240
            Left            =   195
            TabIndex        =   84
            Top             =   1830
            Width           =   1155
         End
         Begin VB.Label Label15 
            Caption         =   "Username:"
            Height          =   240
            Left            =   195
            TabIndex        =   83
            Top             =   1035
            Width           =   1155
         End
         Begin VB.Label Label14 
            Caption         =   "Host name:"
            Height          =   240
            Left            =   195
            TabIndex        =   82
            Top             =   270
            Width           =   1155
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Log File Path:"
         Height          =   1215
         Left            =   210
         TabIndex        =   72
         Top             =   7215
         Width           =   7035
         Begin VB.TextBox txtLoggingPath 
            Height          =   315
            Left            =   330
            TabIndex        =   75
            Top             =   405
            Width           =   6120
         End
         Begin VB.CommandButton cmdLoggingPath 
            Caption         =   "..."
            Height          =   255
            Left            =   6555
            TabIndex        =   74
            Top             =   450
            Width           =   330
         End
         Begin VB.CheckBox chkLoggingEnabled 
            Alignment       =   1  'Right Justify
            Caption         =   "Enable Logging:"
            Height          =   240
            Left            =   5085
            TabIndex        =   73
            Top             =   870
            Width           =   1740
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "FTP File Path:"
         Height          =   1215
         Left            =   210
         TabIndex        =   71
         Top             =   5925
         Width           =   7035
         Begin VB.CommandButton cmdFTPFilePath 
            Caption         =   "..."
            Height          =   255
            Left            =   6555
            TabIndex        =   77
            Top             =   540
            Width           =   330
         End
         Begin VB.TextBox txtFTPFilePath 
            Height          =   315
            Left            =   330
            TabIndex        =   76
            Top             =   510
            Width           =   6120
         End
      End
      Begin MSComctlLib.ListView lstWatch 
         Height          =   2595
         Left            =   -74730
         TabIndex        =   63
         Top             =   3570
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   4577
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Message"
            Object.Width           =   4163
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Msg Event"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Msg Value"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstScoreboard 
         Height          =   2595
         Left            =   -74730
         TabIndex        =   62
         Top             =   810
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   4577
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Message"
            Object.Width           =   4163
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Msg Event"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Msg Value"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstMessages 
         Height          =   5340
         Left            =   -67485
         TabIndex        =   61
         Top             =   825
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   9419
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Message"
            Object.Width           =   4164
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Msg Event"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Msg Value"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdDump 
         Caption         =   "&Dump"
         Height          =   300
         Left            =   -68625
         TabIndex        =   59
         Top             =   5820
         Width           =   915
      End
      Begin VB.CommandButton cmdPktClear 
         Caption         =   "Clear"
         Height          =   300
         Left            =   -67650
         TabIndex        =   58
         Top             =   5820
         Width           =   915
      End
      Begin VB.Frame Frame4 
         Caption         =   "Local UDP Server"
         Height          =   7695
         Left            =   -66615
         TabIndex        =   45
         Top             =   750
         Width           =   3315
         Begin VB.CommandButton cmdListenPortChange 
            Caption         =   "Change"
            Height          =   300
            Left            =   2235
            TabIndex        =   68
            Top             =   2715
            Width           =   990
         End
         Begin VB.TextBox txtListenPortChange 
            Height          =   315
            Left            =   165
            TabIndex        =   67
            Top             =   2715
            Width           =   2010
         End
         Begin VB.CommandButton cmdChangeIP 
            Caption         =   "Change"
            Height          =   300
            Left            =   2235
            TabIndex        =   66
            Top             =   1905
            Width           =   990
         End
         Begin VB.TextBox txtLocalIPChange 
            Height          =   315
            Left            =   165
            TabIndex        =   65
            Top             =   1905
            Width           =   2010
         End
         Begin MSComctlLib.ListView lstGameServers 
            Height          =   3435
            Left            =   150
            TabIndex        =   57
            Top             =   4065
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   6059
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Host"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Packets"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label22 
            Caption         =   "Game Server Packets:"
            Height          =   285
            Left            =   180
            TabIndex        =   56
            Top             =   3765
            Width           =   1875
         End
         Begin VB.Line Line8 
            BorderColor     =   &H80000004&
            X1              =   135
            X2              =   3200
            Y1              =   3435
            Y2              =   3435
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000000&
            X1              =   135
            X2              =   3200
            Y1              =   3420
            Y2              =   3420
         End
         Begin VB.Label lblDetsListenPort 
            Caption         =   "io"
            Height          =   225
            Left            =   1065
            TabIndex        =   55
            Top             =   2460
            Width           =   2010
         End
         Begin VB.Label lblDetsLocalIP 
            Caption         =   "io"
            Height          =   270
            Left            =   1245
            TabIndex        =   54
            Top             =   1650
            Width           =   1800
         End
         Begin VB.Label lblDetsPacketLength 
            Caption         =   "io"
            Height          =   255
            Left            =   1980
            TabIndex        =   53
            Top             =   1215
            Width           =   630
         End
         Begin VB.Label lblDetsProtocol 
            Caption         =   "io"
            Height          =   210
            Left            =   900
            TabIndex        =   52
            Top             =   795
            Width           =   705
         End
         Begin VB.Label lblDetsVersion 
            Caption         =   "io"
            Height          =   210
            Left            =   900
            TabIndex        =   51
            Top             =   390
            Width           =   780
         End
         Begin VB.Label Label21 
            Caption         =   "Listen port:"
            Height          =   225
            Left            =   180
            TabIndex        =   50
            Top             =   2460
            Width           =   900
         End
         Begin VB.Label Label20 
            Caption         =   "Local Host IP:"
            Height          =   270
            Left            =   180
            TabIndex        =   49
            Top             =   1656
            Width           =   1065
         End
         Begin VB.Label Label19 
            Caption         =   "Maximum packet length:"
            Height          =   240
            Left            =   180
            TabIndex        =   48
            Top             =   1215
            Width           =   1800
         End
         Begin VB.Label Label18 
            Caption         =   "Protocol:"
            Height          =   210
            Left            =   180
            TabIndex        =   47
            Top             =   802
            Width           =   705
         End
         Begin VB.Label Label17 
            Caption         =   "Version:"
            Height          =   210
            Left            =   180
            TabIndex        =   46
            Top             =   390
            Width           =   780
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "UDP Server"
         Height          =   2070
         Left            =   -69540
         TabIndex        =   41
         Top             =   6375
         Width           =   2805
         Begin VB.CommandButton cmdListen 
            Caption         =   "&Listen"
            Height          =   315
            Left            =   630
            TabIndex        =   43
            Top             =   1275
            Width           =   1575
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Status:"
            Height          =   210
            Left            =   585
            TabIndex        =   42
            Top             =   600
            Width           =   825
         End
         Begin VB.Shape shStatus 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   360
            Left            =   1530
            Top             =   510
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Connected GameServers"
         Height          =   2070
         Left            =   -74850
         TabIndex        =   37
         Top             =   6375
         Width           =   5175
         Begin VB.CommandButton cmdRecapture 
            Caption         =   "Re&capture"
            Height          =   330
            Left            =   3450
            TabIndex        =   44
            Top             =   1620
            Width           =   1320
         End
         Begin VB.CommandButton cmdPauseIP 
            Caption         =   "&Pause Capture"
            Height          =   330
            Left            =   3450
            TabIndex        =   40
            Top             =   1170
            Width           =   1320
         End
         Begin VB.CommandButton cmdRemoveSelected 
            Caption         =   "&Remove IP"
            Height          =   330
            Left            =   3450
            TabIndex        =   39
            Top             =   480
            Width           =   1320
         End
         Begin VB.ListBox lstConnected 
            Height          =   1740
            ItemData        =   "frmMain.frx":0070
            Left            =   135
            List            =   "frmMain.frx":0072
            TabIndex        =   38
            Top             =   240
            Width           =   2895
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000005&
            X1              =   3405
            X2              =   4800
            Y1              =   990
            Y2              =   990
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000000&
            X1              =   3405
            X2              =   4800
            Y1              =   975
            Y2              =   975
         End
      End
      Begin MSComctlLib.ListView lstData 
         Height          =   4965
         Left            =   -74865
         TabIndex        =   36
         Top             =   810
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   8758
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "UDPDate"
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "UDPTime"
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "UDPRemoteIP"
            Text            =   "Remote IP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "UDPRemotePort"
            Text            =   "Remote Port"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "UDPData"
            Text            =   "Data"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Packet Length"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame fraFile 
         Caption         =   "Message Files"
         Height          =   2265
         Left            =   -74745
         TabIndex        =   19
         Top             =   6195
         Width           =   7095
         Begin VB.CommandButton cmdMessageFileMessage 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6630
            TabIndex        =   28
            Top             =   1860
            Width           =   330
         End
         Begin VB.TextBox txtMessageFileMessage 
            Height          =   315
            Left            =   1020
            TabIndex        =   26
            Top             =   1815
            Width           =   5490
         End
         Begin VB.CommandButton cmdMessageFileWatch 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6630
            TabIndex        =   25
            Top             =   1132
            Width           =   330
         End
         Begin VB.TextBox txtMessageFileWatch 
            Height          =   315
            Left            =   1020
            TabIndex        =   23
            Top             =   1095
            Width           =   5490
         End
         Begin VB.CommandButton cmdMessageFileScoreboard 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6630
            TabIndex        =   22
            Top             =   405
            Width           =   330
         End
         Begin VB.TextBox txtMessageFileScoreboard 
            Height          =   315
            Left            =   1020
            TabIndex        =   20
            ToolTipText     =   "fark off"
            Top             =   375
            Width           =   5490
         End
         Begin VB.Label Label8 
            Caption         =   "Message:"
            Height          =   300
            Left            =   120
            TabIndex        =   27
            Top             =   1815
            Width           =   840
         End
         Begin VB.Label Label7 
            Caption         =   "Watch:"
            Height          =   300
            Left            =   120
            TabIndex        =   24
            Top             =   1095
            Width           =   840
         End
         Begin VB.Label Label6 
            Caption         =   "Scoreboard:"
            Height          =   300
            Left            =   120
            TabIndex        =   21
            Top             =   375
            Width           =   915
         End
      End
      Begin VB.Frame fraMessages 
         Caption         =   "Messages"
         Height          =   2265
         Left            =   -67470
         TabIndex        =   9
         Top             =   6195
         Width           =   4095
         Begin VB.CheckBox chkCommandValue 
            Alignment       =   1  'Right Justify
            Caption         =   "Command Value"
            Enabled         =   0   'False
            Height          =   255
            Left            =   135
            TabIndex        =   64
            Top             =   1185
            Width           =   1530
         End
         Begin VB.CheckBox chkCommandEvent 
            Alignment       =   1  'Right Justify
            Caption         =   "Command Event"
            Height          =   255
            Left            =   135
            TabIndex        =   60
            Top             =   780
            Width           =   1530
         End
         Begin VB.CommandButton cmdClearSel 
            Caption         =   "&Clear Selections"
            Height          =   345
            Left            =   150
            TabIndex        =   35
            Top             =   1785
            Width           =   1650
         End
         Begin VB.CommandButton cmdRemoveMessage 
            Caption         =   "Remove Messa&ge"
            Height          =   345
            Left            =   2310
            TabIndex        =   18
            Top             =   1785
            Width           =   1650
         End
         Begin VB.CommandButton cmdAddMsg 
            Caption         =   "Add Me&ssage"
            Height          =   345
            Left            =   2310
            TabIndex        =   12
            Top             =   720
            Width           =   1650
         End
         Begin VB.TextBox txtNewMsg 
            Height          =   345
            Left            =   1110
            TabIndex        =   10
            Top             =   255
            Width           =   2850
         End
         Begin VB.Line Line10 
            BorderColor     =   &H80000004&
            X1              =   150
            X2              =   3930
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Line Line9 
            BorderColor     =   &H80000000&
            X1              =   150
            X2              =   3930
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Add:"
            Height          =   345
            Left            =   195
            TabIndex        =   11
            Top             =   315
            Width           =   810
         End
      End
      Begin VB.Frame fraWatch 
         Caption         =   "Watch"
         Height          =   2685
         Left            =   -70425
         TabIndex        =   3
         Top             =   3480
         Width           =   2775
         Begin VB.CommandButton cmdAddValueWatch 
            Caption         =   "A&dd"
            Height          =   240
            Left            =   1575
            TabIndex        =   34
            Top             =   2250
            Width           =   1065
         End
         Begin VB.TextBox txtValueWatch 
            Height          =   315
            Left            =   135
            TabIndex        =   33
            Top             =   2205
            Width           =   1380
         End
         Begin VB.CommandButton cmdRemoveAllWatch 
            Caption         =   "Re&move All"
            Height          =   240
            Left            =   1590
            TabIndex        =   17
            Top             =   1230
            Width           =   1065
         End
         Begin VB.CommandButton cmdRemoveWatch 
            Caption         =   ">>"
            Height          =   240
            Left            =   1590
            TabIndex        =   15
            Top             =   795
            Width           =   1065
         End
         Begin VB.CommandButton cmdAddWatch 
            Caption         =   "<<"
            Height          =   240
            Left            =   1590
            TabIndex        =   13
            Top             =   345
            Width           =   1065
         End
         Begin VB.Label Label10 
            Caption         =   "Add value for selected message"
            Height          =   240
            Left            =   225
            TabIndex        =   32
            Top             =   1830
            Width           =   2400
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000005&
            X1              =   135
            X2              =   2700
            Y1              =   1635
            Y2              =   1635
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000000&
            X1              =   135
            X2              =   2700
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Remove"
            Height          =   315
            Left            =   165
            TabIndex        =   16
            Top             =   750
            Width           =   1245
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Add"
            Height          =   300
            Left            =   165
            TabIndex        =   14
            Top             =   300
            Width           =   1230
         End
      End
      Begin VB.Frame fraScoreboard 
         Caption         =   "Scoreboard"
         Height          =   2685
         Left            =   -70425
         TabIndex        =   2
         Top             =   720
         Width           =   2775
         Begin VB.CommandButton cmdAddValueScoreboard 
            Caption         =   "&Add"
            Height          =   240
            Left            =   1590
            TabIndex        =   31
            Top             =   2250
            Width           =   1065
         End
         Begin VB.TextBox txtValueScoreboard 
            Height          =   315
            Left            =   150
            TabIndex        =   30
            Top             =   2205
            Width           =   1380
         End
         Begin VB.CommandButton cmdRemoveAllScoreboard 
            Caption         =   "&Remove All"
            Height          =   240
            Left            =   1590
            TabIndex        =   8
            Top             =   1185
            Width           =   1065
         End
         Begin VB.CommandButton cmdRemoveScoreboardMsg 
            Caption         =   ">>"
            Height          =   240
            Left            =   1590
            TabIndex        =   6
            Top             =   750
            Width           =   1065
         End
         Begin VB.CommandButton cmdAddScoreboardMsg 
            Caption         =   "<<"
            Height          =   240
            Left            =   1605
            TabIndex        =   4
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label Label9 
            Caption         =   "Add value for selected message"
            Height          =   240
            Left            =   240
            TabIndex        =   29
            Top             =   1830
            Width           =   2400
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   150
            X2              =   2715
            Y1              =   1635
            Y2              =   1635
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            X1              =   150
            X2              =   2715
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Remove"
            Height          =   315
            Left            =   165
            TabIndex        =   7
            Top             =   750
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Add"
            Height          =   300
            Left            =   180
            TabIndex        =   5
            Top             =   300
            Width           =   1230
         End
      End
      Begin MSComctlLib.ListView lstFTP 
         Height          =   4965
         Left            =   135
         TabIndex        =   69
         Top             =   810
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8758
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "UDPDate"
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "UDPTime"
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "FTP Status"
            Text            =   "FTP Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "Current Ranked Game Servers:"
         Height          =   300
         Left            =   -74850
         TabIndex        =   89
         Top             =   495
         Width           =   3345
      End
      Begin VB.Label Label13 
         Caption         =   "Current FTP Status:"
         Height          =   285
         Left            =   150
         TabIndex        =   70
         Top             =   495
         Width           =   1515
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1665
      Top             =   7665
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
      RequestTimeout  =   15
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   360
      Left            =   10125
      TabIndex        =   0
      Top             =   8880
      Width           =   1755
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   7725
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   41388
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'-          GODDESS Version 1.0.0
'-
'-  Author: Sebastian Quelcutti
'-
'-          Main application form
'------------------------------------------------------------------------------------
Option Explicit

'the udp listener
Private WithEvents moUDPListener As CUDPListener
Attribute moUDPListener.VB_VarHelpID = -1

Private Const MODULE As String = "frmMain:"

'list of paused servers
Private mcolPaused As Collection
Private mbHasClicked As Boolean

Private Sub chkCommandEvent_Click()

    If chkCommandEvent.Value = vbChecked Then
        chkCommandValue.Enabled = True
    Else
        chkCommandValue.Enabled = False
    End If
    
End Sub

Private Sub chkLoggingEnabled_Click()

    Dim lRet As Long
    
    'are we currently logging and has this changed to vbchecked
    If g_bIsLogging And chkLoggingEnabled.Value = vbChecked Then
        Exit Sub
    End If
    
    'have we set TO logging
    If chkLoggingEnabled.Value = vbChecked Then
        'set logging then
        g_oGoddess.SetLoggingEnabled True
        StartLogging
    Else
        lRet = MsgBox("Are you sure you wish to disable logging?", vbQuestion + vbYesNo, "Disable Logging")
        If lRet = vbNo Then
            Exit Sub
        End If
        If g_bIsLogging Then
            'just to be on the safe side
            Log "User disabling log."
        End If
        g_oGoddess.SetLoggingEnabled False
        StopLogging
    End If
    
End Sub

Private Sub cmdAddMsg_Click()

    Dim lRet As Long
    Dim bCommand As Boolean
    Dim bValue As Boolean
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdAddMsg_Click_Error
    
    bCommand = False
    
    'make sure there is some text in the box first
    If Trim$(txtNewMsg.Text) = "" Then
        MsgBox "Please type a new message into the provided box!", vbExclamation + vbOKOnly, "Missing message name"
        txtNewMsg.SetFocus
        GoTo Exit_Properly
    End If
    
    'got something, so add it and refresh
    If chkCommandEvent.Value = vbChecked Then
        bCommand = True
    End If
    If chkCommandValue.Value = vbChecked Then
        bValue = True
    End If
    
    lRet = g_oGoddess.AddMessage(LCase$(Trim$(txtNewMsg.Text)), bCommand, bValue, True)
    If Not lRet Then
        MsgBox "Failed to add new message to list!", vbExclamation + vbOKOnly, "Failed AddMessage"
        txtNewMsg.SetFocus
    End If
    
    Log "New message - " & LCase$(Trim$(txtNewMsg.Text)) & " - has been added."
    
    'now repopulate the list
    lstMessages.ListItems.Clear
    
    PopulateMessages
    
    'finito
    txtNewMsg.Text = ""
    
    'reset checkbox
    chkCommandEvent.Value = vbUnchecked
    
Exit_Properly:
    Exit Sub
    
cmdAddMsg_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdAddMsg_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdAddMsg_Click - " & sSource
    Log "**************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdAddScoreboardMsg_Click()

    Dim sName As String
    Dim lIndex As Long
    Dim bInList As Boolean
    Dim sOutstr As String
    Dim bCommandEvent As Boolean
    Dim bCommandValue As Boolean
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdAddScoreboardMsg_Click_Error
    
    bInList = False
    sOutstr = "The following items were not added to the scoreboard list" & vbCr & _
    "because they are already present there:" & vbCr & vbCr
    
    'first make sure we have a selected message
    If Not MessagesSelCount > 0 Then
        MsgBox "Please select a message to add to the Scoreboard filter!", vbExclamation + vbOKOnly, "No Message selected"
        Exit Sub
    End If
    
    bCommandEvent = False
    bCommandValue = False
    
    'loop through selected messages
    lIndex = MessagesSelCount
    If lIndex > 0 Then
        For lIndex = 1 To lstMessages.ListItems.Count
            If lstMessages.ListItems(lIndex).Selected Then
                'now we must add this to the scoreboard list
                sName = lstMessages.ListItems(lIndex).Text
                'is this item from messages list already in the scoreboard list
                If IsItemInList(lstScoreboard, sName) Then
                    bInList = True
                    sOutstr = sOutstr & sName & vbCr
                Else
                    If LCase$(Trim$(lstMessages.ListItems(lIndex).SubItems(1))) = "true" Then
                        bCommandEvent = True
                    End If
                    If LCase$(Trim$(lstMessages.ListItems(lIndex).SubItems(2))) = "true" Then
                        bCommandValue = True
                    End If
                    g_oGoddess.AddFilter "Scoreboard", sName, "", bCommandEvent, bCommandValue, True
                    Log "Scoreboard filter '" & sName & "' has been added."
                End If
            End If
        Next lIndex
        PopulateFilters
    Else
        MsgBox "Please select a message to add to Scoreboard list!", vbExclamation + vbOKOnly, "Message not selected"
    End If
    
    If bInList Then
        MsgBox sOutstr, vbInformation + vbOKOnly, "Duplicate messages"
    End If
    
Exit_Properly:
    Exit Sub
    
cmdAddScoreboardMsg_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdAddScoreboardMsg_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "****************************************************"
    Log "Error occured in GODDESS"
    Log "****************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdAddScoreboardMsg_Click - " & sSource
    Log "****************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdAddValidGameServer_Click()

    Dim lRet As Long
    Dim bRet As Boolean
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdAddValidGameServer_Click_Error
    
    'do we have required values in boxes
    'valid servername?
    If Trim$(txtValidServerName.Text) = "" Then
        lRet = MsgBox("You must provide a valid Server Name before adding a Ranked Server!" & vbCr & _
        "Please enter a valid Server Name into the box provided.", vbExclamation + vbOKOnly, _
        "Invalid Server Name")
        txtValidServerName.SetFocus
        Exit Sub
    End If
    
    'valid ip address?
    If Trim$(txtValidServerIP.Text) = "" Or Not ValidateIPAddressFormat(Trim$(txtValidServerIP.Text)) Then
        lRet = MsgBox("You must provide a valid Server IP before adding a Ranked Server!" & vbCr & _
        "Please enter a valid Server IP into the box provided.", vbExclamation + vbOKOnly, _
        "Invalid Server IP")
        txtValidServerIP.SetFocus
        Exit Sub
    End If
    
    'valid port?
    If Trim$(txtValidServerPort.Text) = "" Or Not IsNumeric(Trim$(txtValidServerPort.Text)) Then
        lRet = MsgBox("You must provide a valid Server Port before adding a Ranked Server!" & vbCr & _
        "Please enter a valid Server Port into the box provided.", vbExclamation + vbOKOnly, _
        "Invalid Server Port")
        txtValidServerPort.SetFocus
        Exit Sub
    End If

    'ok, valid entries - does this server already exist
    bRet = IsValidServerInList(LCase$(Trim$(txtValidServerIP.Text)), Trim$(txtValidServerPort.Text))
    If bRet Then
        lRet = MsgBox("The Ranked Server you have specified already exists!" & vbCr & _
        "Please enter a valid new Ranked Server Port into the boxes provided.", vbExclamation + vbOKOnly, _
        "Existing Ranked Server")
        txtValidServerName.SetFocus
        Exit Sub
    End If
    
    'server is valid and doesn't exist, so add it
    lRet = g_oGoddess.AddValidGameServer(Trim$(txtValidServerName.Text), Trim$(txtValidServerIP.Text), Trim$(txtValidServerPort.Text), True)
    If Not lRet Then
        MsgBox "Failed to add new Game Server to list!", vbExclamation + vbOKOnly, "Failed Adding GameServer"
        txtValidServerName.SetFocus
        Log "Failed to add new Ranked Server - Host '" & Trim$(txtValidServerName.Text) & _
        "', IP '" & Trim$(txtValidServerIP.Text) & "', Port '" & Trim$(txtValidServerPort.Text) & "'."
        Exit Sub
    End If
    
    Log "New Ranked Server - " & Trim$(txtValidServerName.Text) & " - has been added."
    
    txtValidServerName.Text = ""
    txtValidServerIP.Text = ""
    txtValidServerPort.Text = ""
    
    PopulateValidGameServers
    
Exit_Properly:
    Exit Sub
    
cmdAddValidGameServer_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdAddValidGameServer_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "****************************************************"
    Log "Error occured in GODDESS"
    Log "****************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdAddValidGameServer_Click - " & sSource
    Log "****************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdAddValueScoreboard_Click()

    Dim oItem As ListItem
    Dim bEvent As Boolean
    Dim bValue As Boolean
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdAddValueScoreboard_Click_Error
    
    'if there is more than one selection, bomb out
    If ScoreboardSelCount > 1 Then
        MsgBox "Unable to add values to multiple selections!", vbExclamation + vbOKOnly, "Multiple files"
        Exit Sub
    End If
    
    If ScoreboardSelCount = 0 Then
        MsgBox "Please select a Scoreboard message to add a value to!", vbExclamation + vbOKOnly, "No Message selected"
        Exit Sub
    End If
    
    Set oItem = lstScoreboard.SelectedItem
    
    If LCase$(Trim$(oItem.SubItems(1))) = "true" Then bEvent = True
    If LCase$(Trim$(oItem.SubItems(2))) = "true" Then bValue = True
    
    If bEvent And Not bValue Then
        MsgBox "'" & Trim$(oItem.Text) & "' is a command Event that does not allow a value!", vbExclamation + vbOKOnly, "Command Event"
        Exit Sub
    End If
    
    'finally we can add the value
    g_oGoddess.Filters("Scoreboard_" & StripValueFromCommand(Trim$(oItem.Text))).FilterValue = Trim$(txtValueScoreboard.Text)
    'update the value in the xml
    g_oGoddess.UpdateScoreboardValue StripValueFromCommand(Trim$(oItem.Text)), Trim$(txtValueScoreboard.Text)
    g_oGoddess.SaveXML xtScoreboard
    
    'and now the listbox
    Dim sName As String
    sName = StripValueFromCommand(Trim$(oItem.Text))
    
    If Trim$(txtValueScoreboard.Text) = "" Then
        oItem.Text = sName
    Else
        oItem.Text = sName & " '" & Trim$(txtValueScoreboard.Text) & "'"
    End If
    
    'done
    txtValueScoreboard.Text = ""
    
Exit_Properly:
    Set oItem = Nothing
    Exit Sub
    
cmdAddValueScoreboard_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdAddValueScoreboard_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdAddValueScoreboard_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdAddValueWatch_Click()

    Dim oItem As ListItem
    Dim bEvent As Boolean
    Dim bValue As Boolean
    Dim lErrno As Long
    Dim sSource As String, sDesc As String

    On Error GoTo cmdAddValueWatch_Click_Error
    
    'if there is more than one selection, bomb out
    If WatchSelCount > 1 Then
        MsgBox "Unable to add values to multiple selections!", vbExclamation + vbOKOnly, "Multiple files"
        Exit Sub
    End If
    
    If WatchSelCount = 0 Then
        MsgBox "Please select a Watch message to add a value to!", vbExclamation + vbOKOnly, "No Message selected"
        Exit Sub
    End If
    
    Set oItem = lstWatch.SelectedItem
    
    If LCase$(Trim$(oItem.SubItems(1))) = "true" Then bEvent = True
    If LCase$(Trim$(oItem.SubItems(2))) = "true" Then bValue = True
    
    If bEvent And Not bValue Then
        MsgBox "'" & Trim$(oItem.Text) & "' is a command Event that does not allow a value!", vbExclamation + vbOKOnly, "Command Event"
        Exit Sub
    End If
    
    'finally we can add the value
    g_oGoddess.Filters("Watch_" & StripValueFromCommand(Trim$(oItem.Text))).FilterValue = Trim$(txtValueWatch.Text)
    'update the value in the xml
    g_oGoddess.UpdateWatchValue StripValueFromCommand(Trim$(oItem.Text)), Trim$(txtValueWatch.Text)
    g_oGoddess.SaveXML xtWatch
    
    'and now the listbox
    Dim sName As String
    sName = StripValueFromCommand(Trim$(oItem.Text))
    
    If Trim$(txtValueWatch.Text) = "" Then
        oItem.Text = sName
    Else
        oItem.Text = sName & " '" & Trim$(txtValueWatch.Text) & "'"
    End If
    
    'done
    txtValueWatch.Text = ""
    
Exit_Properly:
    Exit Sub
    
cmdAddValueWatch_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdAddValueWatch_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdAddValueWatch_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdAddWatch_Click()

    Dim sName As String
    Dim lIndex As Long
    Dim bInList As Boolean
    Dim sOutstr As String
    Dim bCommandEvent As Boolean
    Dim bCommandValue As Boolean
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdAddWatch_Click_Error
    
    bInList = False
    sOutstr = "The following items were not added to the Watch list" & vbCr & _
    "because they are already present there:" & vbCr & vbCr
    
    'first make sure we have a selected message
    If Not MessagesSelCount > 0 Then
        MsgBox "Please select a message to add to the Watch filter!", vbExclamation + vbOKOnly, "No Message selected"
        Exit Sub
    End If
    
    bCommandEvent = False
    bCommandValue = False
    
    'loop through selected messages
    lIndex = MessagesSelCount
    If lIndex > 0 Then
        For lIndex = 1 To lstMessages.ListItems.Count
            If lstMessages.ListItems(lIndex).Selected Then
                'now we must add this to the watch list
                sName = lstMessages.ListItems(lIndex).Text
                If IsItemInList(lstWatch, sName) Then
                    bInList = True
                    sOutstr = sOutstr & sName & vbCr
                Else
                    If LCase$(Trim$(lstMessages.ListItems(lIndex).SubItems(1))) = "true" Then
                        bCommandEvent = True
                    End If
                    If LCase$(Trim$(lstMessages.ListItems(lIndex).SubItems(2))) = "true" Then
                        bCommandValue = True
                    End If
                    g_oGoddess.AddFilter "Watch", sName, "", bCommandEvent, bCommandValue, True
                    Log "Watch filter '" & sName & "' has been added."
                End If
            End If
        Next lIndex
        PopulateFilters
    Else
        MsgBox "Please select a message to add to Watch list!", vbExclamation + vbOKOnly, "Message not selected"
    End If
    
    If bInList Then
        MsgBox sOutstr, vbInformation + vbOKOnly, "Duplicate messages"
    End If

Exit_Properly:
    Exit Sub
    
cmdAddWatch_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdAddWatch_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdAddWatch_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdChangeIP_Click()

    Dim bRet As Boolean
    Dim lRet As Long
            
    If Trim$(txtLocalIPChange.Text) = "" Then
        MsgBox "Please enter a new IP address ( xxx.xxx.xxx.xxx ) into the box!", vbExclamation + vbOKOnly, "Missing IP Address"
        txtLocalIPChange.SetFocus
        Exit Sub
    End If
    
    If Not ValidateIPAddressFormat(Trim$(txtLocalIPChange.Text)) Then
        MsgBox "Please enter a valid IP address ( xxx.xxx.xxx.xxx ) into the box!", vbExclamation + vbOKOnly, "Missing IP Address"
        txtLocalIPChange.SetFocus
        Exit Sub
    End If
    
    'are we currently listening?
    If Trim$(cmdListen.Caption) <> "&Listen" Then
        lRet = MsgBox("To change the Local Listen IP address you must stop listening on the port." & vbCr & vbCr & _
        "Would you like to stop the server listening now?", vbQuestion + vbYesNo, "Stop Listen server")
        If lRet = vbNo Then
            txtLocalIPChange.Text = ""
            Exit Sub
        End If
        'stop the listen server
        cmdListen_Click
    End If
    
    'now we have a valid ip address, change it in the goddess
    g_oGoddess.SetHostIP Trim$(txtLocalIPChange.Text)
    moUDPListener.LocalIP = Trim$(txtLocalIPChange.Text)
    
    Log "Host IP address changed to '" & Trim$(txtLocalIPChange.Text) & "'"
    
    'update the udp details
    UpdateUDPDetails
    
    txtLocalIPChange.Text = ""
        
End Sub

Private Sub cmdClearSel_Click()

    ClearSelections lstMessages, xtMessages
    
End Sub

Private Sub cmdDump_Click()

    'this routine dumps the display of the list view for each entry
    Dim fs As FileSystemObject
    Dim sDump As String
    Dim lIndex As Long
    Dim sPath As String
    Dim sSep As String
    Dim oStream As TextStream
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdDump_Click_Error
    
    'only if there is data
    If lstData.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    Set fs = New FileSystemObject
    
    'cycle through the list
    For lIndex = 1 To lstData.ListItems.Count
        sDump = sDump & lstData.ListItems(lIndex).SubItems(4) & vbCrLf
    Next lIndex
    
    'now save to dump file
    sSep = ""
    If Not LCase$(Trim$(App.Path)) = "c:" And Not LCase$(Trim$(App.Path)) = "c:\" Then
        sSep = "\"
    End If
    
    sPath = App.Path & sSep & "GoddessDump_" & Replace(Time, ":", "_") & ".txt"
    
    Set oStream = fs.CreateTextFile(sPath, True)
    
    oStream.Write sDump
    
    oStream.Close

Exit_Properly:
    Set oStream = Nothing
    Set fs = Nothing
    Exit Sub
    
cmdDump_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdDump_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCr & vbCr & _
    "Source: " & Err.Source, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdDump_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdFTPFilePath_Click()

    Dim frmMF As New frmMessageFiles
    Dim lRet As Long
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdFTPFilePath_Click_error
        
    frmMF.PathIncoming = Trim(txtFTPFilePath.Text)
    frmMF.MessageFileType = "FTP"
        
    frmMF.Show vbModal
    
    If Len(frmMF.MessageFilePath) > 0 Then
        txtFTPFilePath.Text = frmMF.MessageFilePath
        txtFTPFilePath.ToolTipText = frmMF.MessageFilePath
        'found a new path, so add to goddess
        g_oGoddess.FTPFilePath = Trim$(txtFTPFilePath.Text)
        g_oGoddess.SetFTPFilePath
        Log "User changed FTP File Path to '" & Trim$(frmMF.MessageFilePath) & "'"
    End If

    
Exit_Properly:
    If Not frmMF Is Nothing Then
        Unload frmMF
        Set frmMF = Nothing
    End If

    Exit Sub

cmdFTPFilePath_Click_error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdFTPFilePath_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdFTPFilePath_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly
    
End Sub

Private Sub cmdFTPHostName_Click()

    Dim lRet As Long
    
    'is there any data
    If Trim$(txtFTPHostName.Text) = "" Then
        MsgBox "Please enter a valid Host name into the box provided.", vbExclamation + vbOKOnly, "Missing Host name"
        txtFTPHostName.Text = g_oGoddess.FTPRemoteHost
        txtFTPHostName.SetFocus
        Exit Sub
    End If
    
    'is the data the same as the current
    If Trim$(LCase$(txtFTPHostName.Text)) = Trim$(LCase$(g_oGoddess.FTPRemoteHost)) Then
        'nothing to do, exit
        Exit Sub
    End If
    
    lRet = MsgBox("Are you sure you want to change the FTP Host name?", vbQuestion + vbYesNo, "Change FTP Host")
    If lRet = vbYes Then
        g_oGoddess.FTPRemoteHost = Trim$(txtFTPHostName.Text)
        g_oGoddess.SetFTPHostName
        Log "FTP Host name changed to '" & g_oGoddess.FTPRemoteHost & "'"
    Else
        txtFTPHostName.Text = g_oGoddess.FTPRemoteHost
        txtFTPHostName.SetFocus
    End If
        
End Sub

Private Sub cmdFTPPassword_Click()

    Dim lRet As Long
    
    'is there any data
    If Trim$(txtFTPPassword.Text) = "" Then
        MsgBox "Please enter a valid password into the box provided.", vbExclamation + vbOKOnly, "Missing Password"
        txtFTPPassword.Text = g_oGoddess.FTPPassword
        txtFTPPassword.SetFocus
        Exit Sub
    End If
    
    'is the data the same as the current
    If Trim$(LCase$(txtFTPPassword.Text)) = Trim$(LCase$(g_oGoddess.FTPPassword)) Then
        'nothing to do, exit
        Exit Sub
    End If
    
    lRet = MsgBox("Are you sure you want to change the FTP Password?", vbQuestion + vbYesNo, "Change FTP Password")
    If lRet = vbYes Then
        g_oGoddess.FTPPassword = Trim$(txtFTPPassword.Text)
        g_oGoddess.SetFTPPassword
        Log "FTP Password changed."
    Else
        txtFTPPassword.Text = g_oGoddess.FTPPassword
        txtFTPPassword.SetFocus
    End If
        
End Sub

Private Sub cmdFTPUsername_Click()

    Dim lRet As Long
    
    'is there any data
    If Trim$(txtFTPUsername.Text) = "" Then
        MsgBox "Please enter a valid User name into the box provided.", vbExclamation + vbOKOnly, "Missing Host name"
        txtFTPUsername.Text = g_oGoddess.FTPUserName
        txtFTPUsername.SetFocus
        Exit Sub
    End If
    
    'is the data the same as the current
    If Trim$(LCase$(txtFTPUsername.Text)) = Trim$(LCase$(g_oGoddess.FTPUserName)) Then
        'nothing to do, exit
        Exit Sub
    End If
    
    lRet = MsgBox("Are you sure you want to change the FTP User name?", vbQuestion + vbYesNo, "Change FTP Username")
    If lRet = vbYes Then
        g_oGoddess.FTPUserName = Trim$(txtFTPUsername.Text)
        g_oGoddess.SetFTPUserName
        Log "FTP User name changed to '" & g_oGoddess.FTPUserName & "'"
    Else
        txtFTPUsername.Text = g_oGoddess.FTPUserName
        txtFTPUsername.SetFocus
    End If
        
End Sub

Private Sub cmdListen_Click()

    Dim lIndex As Long
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdListen_Click_Error
    
    'set the listener to listen and change the command button text
    'if the caption is "&Listen" then we are off
    If cmdListen.Caption = "&Listen" Then
        'change caption
        cmdListen.Caption = "&Stop Listening"
        moUDPListener.UDPListen
        'status is now green
        shStatus.FillColor = &HFF00&
        Log "GODDESS listen mode activated."
    Else
        'caption is already '&Stop Listening" so we are listening and want to stop
        cmdListen.Caption = "&Listen"
        moUDPListener.UDPClose
        'status is now red
        shStatus.FillColor = &HFF&
        'now clear up any gameserver data that hasn't been ftp'd yet
        moUDPListener.GameServers.Clear
        'and remove any pauses in its collection
        For lIndex = 1 To mcolPaused.Count
            mcolPaused.Remove (0)
        Next lIndex
        UpdateUDPDetails
        lstConnected.Clear
        Log "GODDESS listen mode deactivated."
    End If

Exit_Properly:
    Exit Sub
    
cmdListen_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdListen_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCr & vbCr & _
    "Source: " & Err.Source, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdListen_Click - " & sSource
    Log "******************************************************"
    cmdListen.Caption = "&Listen"
    shStatus.FillColor = &HFF&
    GoTo Exit_Properly

End Sub

Private Sub cmdListenPortChange_Click()

    Dim bRet As Boolean
    Dim lRet As Long
    Dim lVal As Long
    
    If Trim$(txtListenPortChange.Text) = "" Then
        MsgBox "Please enter a new port (1500 - 65535) into the box!", vbExclamation + vbOKOnly, "Missing IP Address"
        txtListenPortChange.SetFocus
        Exit Sub
    End If
    
    'is it in range 1500 65535
    If Not IsNumeric(Trim$(txtListenPortChange.Text)) Then
        MsgBox "Please enter a valid port (1500 - 65535) into the box!", vbExclamation + vbOKOnly, "Missing IP Address"
        txtListenPortChange.SetFocus
        Exit Sub
    End If
    
    
    'are we currently listening?
    If Trim$(cmdListen.Caption) <> "&Listen" Then
        lRet = MsgBox("To change the Local Listen port you must stop listening on the port." & vbCr & vbCr & _
        "Would you like to stop the server listening now?", vbQuestion + vbYesNo, "Stop Listen server")
        If lRet = vbNo Then
            txtListenPortChange.Text = ""
            Exit Sub
        End If
        'stop the listen server
        cmdListen_Click
    End If
    
    'now we have a valid ip address, change it in the goddess
    g_oGoddess.SetHostPort Trim$(txtListenPortChange.Text)
    moUDPListener.LocalPort = Trim$(txtListenPortChange.Text)
    
    Log "GODDESS Listen Port has been changed to '" & Trim$(txtListenPortChange.Text) & "'"
    
    'update the udp details
    UpdateUDPDetails
    
    txtListenPortChange.Text = ""
        
End Sub

Private Sub cmdLoggingPath_Click()

    Dim frmMF As New frmMessageFiles
    Dim lRet As Long
    
    On Error GoTo cmdLoggingPath_Click_error
    
    'are we currently logging
    If g_bIsLogging Then
        'user should know
        lRet = MsgBox("Logging is currently active!" & vbCr & _
        "Are you sure you want to change the Log File path (current logging will stop)?", _
        vbQuestion + vbYesNo, "Change Log File Path")
        If lRet = vbNo Then
            Exit Sub
        End If
        'user wants to, stop logging
        Log "User is changing Log Path - logging is terminating."
        StopLogging
        g_bIsLogging = False
    End If
    
    frmMF.PathIncoming = Trim(txtLoggingPath.Text)
    frmMF.MessageFileType = "Logging"
        
    frmMF.Show vbModal
    
    If Len(frmMF.MessageFilePath) > 0 Then
        txtLoggingPath.Text = frmMF.MessageFilePath
        txtLoggingPath.ToolTipText = frmMF.MessageFilePath
        'found a new path, so add to goddess
        g_oGoddess.SetLoggingFilePath Trim$(frmMF.MessageFilePath)
        'ask if start logging now
        lRet = MsgBox("You have successfully changed the Log File Path!" & vbCr & _
        "Would you like to start GODDESS Logging now?", vbQuestion + vbYesNo, "Start Logging")
        If lRet = vbYes Then
            StartLogging
            Log "GODDESS Log File Path changed to '" & Trim$(frmMF.MessageFilePath) & "'"
        End If
    End If

    
Exit_Properly:
    If Not frmMF Is Nothing Then
        Unload frmMF
        Set frmMF = Nothing
    End If

    Exit Sub

cmdLoggingPath_Click_error:
    MsgBox "The following error has occured in frmMain:cmdLoggingPath_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    GoTo Exit_Properly
    
End Sub

Private Sub cmdMessageFileMessage_Click()

    Dim frmMF As New frmMessageFiles
    Dim lErrno As Long
    Dim sSource As String, sDesc As String

    On Error GoTo cmdMessageFileMessage_Click_error
        
    frmMF.PathIncoming = Trim(txtMessageFileMessage.Text)
    frmMF.MessageFileType = "Message"
    
    frmMF.Show vbModal
    
    If Len(frmMF.MessageFilePath) > 0 Then
        txtMessageFileMessage.Text = frmMF.MessageFilePath
        txtMessageFileMessage.ToolTipText = frmMF.MessageFilePath
        'found a new path, so add to goddess
        g_oGoddess.SetMessageFilePath "Message", Trim$(frmMF.MessageFilePath)
        Log "Message file path changed to '" & Trim$(frmMF.MessageFilePath) & "'"
    End If


Exit_Properly:
    If Not frmMF Is Nothing Then
        Unload frmMF
        Set frmMF = Nothing
    End If

    Exit Sub

cmdMessageFileMessage_Click_error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdMessageFileMessage_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdMessageFileMessage_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly
    
End Sub

Private Sub cmdMessageFileScoreboard_Click()

    Dim frmMF As New frmMessageFiles
    Dim lErrno As Long
    Dim sSource As String, sDesc As String

    On Error GoTo cmdMessageFileScoreboard_Click_error
        
    frmMF.PathIncoming = Trim(txtMessageFileScoreboard.Text)
    frmMF.MessageFileType = "Scoreboard"
        
    frmMF.Show vbModal
    
    If Len(frmMF.MessageFilePath) > 0 Then
        txtMessageFileScoreboard.Text = frmMF.MessageFilePath
        txtMessageFileScoreboard.ToolTipText = frmMF.MessageFilePath
        'found a new path, so add to goddess
        g_oGoddess.SetMessageFilePath "Scoreboard", Trim$(frmMF.MessageFilePath)
        Log "Scoreboard filter file path changed to '" & Trim$(frmMF.MessageFilePath) & "'"
    End If


Exit_Properly:
    If Not frmMF Is Nothing Then
        Unload frmMF
        Set frmMF = Nothing
    End If

    Exit Sub

cmdMessageFileScoreboard_Click_error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdMessageFileScoreboard_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdMessageFileScoreboard_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly
    
End Sub

Private Sub cmdMessageFileWatch_Click()

    Dim frmMF As New frmMessageFiles
    Dim lErrno As Long
    Dim sSource As String, sDesc As String

    On Error GoTo cmdMessageFileWatch_Click_error
        
    frmMF.PathIncoming = Trim(txtMessageFileWatch.Text)
    frmMF.MessageFileType = "Watch"
    
'    Load frmMF
    
    frmMF.Show vbModal
    
    If Len(frmMF.MessageFilePath) > 0 Then
        txtMessageFileWatch.Text = frmMF.MessageFilePath
        txtMessageFileWatch.ToolTipText = frmMF.MessageFilePath
        'found a new path, so add to goddess
        g_oGoddess.SetMessageFilePath "Watch", Trim$(frmMF.MessageFilePath)
        Log "Watch filter file path changed to '" & Trim$(frmMF.MessageFilePath) & "'"
    End If


Exit_Properly:
    If Not frmMF Is Nothing Then
        Unload frmMF
        Set frmMF = Nothing
    End If

    Exit Sub

cmdMessageFileWatch_Click_error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdMessageFileWatch_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdMessageFileWatch_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly
    
End Sub

Private Sub cmdPauseIP_Click()

    Dim sName As String
    Dim bRet As Boolean
    Dim sRemoteIP As String
    Dim lRemotePort As Long
    Dim oItem As ListItem
    
    'is there a selection?
    If lstConnected.SelCount = 0 Then
        MsgBox "Please select a server to pause from the connected list!", vbExclamation + vbOKOnly, "No server selected"
        Exit Sub
    End If
    
    'is this server already paused?
    sName = lstConnected.List(lstConnected.ListIndex)
    
    If Trim$(sName) = "" Then
        Exit Sub
    End If
    
    bRet = GetPausedIPAndPort(sName, sRemoteIP, lRemotePort)
    
    bRet = IsServerPaused(sRemoteIP, lRemotePort)
    
    If Not bRet Then
        'not paused, so pause it
        PauseServer sRemoteIP, lRemotePort
        Log "Paused incoming UDP packet data from remote IP '" & sRemoteIP & "' port '" & lRemotePort & "'"
    End If
    
End Sub

Private Sub cmdPktClear_Click()

    lstData.ListItems.Clear
    
End Sub

Private Sub cmdQuit_Click()

    'stop listening
    moUDPListener.UDPClose
    
    Unload Me

End Sub

Private Sub cmdRecapture_Click()

    Dim sName As String
    Dim sRemoteIP As String
    Dim lRemotePort As Long
    Dim bRet As Boolean
    
    'remove this from the paused list
    If lstConnected.SelCount = 0 Then
        MsgBox "Please select a paused server to recapture data from!", vbExclamation + vbOKOnly, "No server selected"
        Exit Sub
    End If
 
    'is this server already paused?
    sName = lstConnected.List(lstConnected.ListIndex)
    
    If Trim$(sName) = "" Then
        Exit Sub
    End If
    
    bRet = GetPausedIPAndPort(sName, sRemoteIP, lRemotePort)
    
    bRet = IsServerPaused(sRemoteIP, lRemotePort)
    
    If bRet Then
        'not paused, so pause it
        UnpauseServer sRemoteIP, lRemotePort
        Log "Recapturing UDP packet data for remote IP '" & sRemoteIP & "' port '" & lRemotePort & "'"
    End If
 
End Sub

Private Sub cmdRemoveAllScoreboard_Click()

    'remove all entries from the scoreboard
    Dim lIndex As Long
    Dim lRet As Long
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdRemoveAllScoreboard_Click_Error
    
    lRet = MsgBox("Are you sure you want to remove all items from the Scoreboard?", vbQuestion + vbYesNo, "Confirm multiple removal")
    
    If lRet = vbNo Then
        Exit Sub
    End If
    
    For lIndex = 1 To lstScoreboard.ListItems.Count
        If LCase$(Trim$(lstScoreboard.ListItems(1).Text)) = "true" Then
            g_oGoddess.RemoveScoreboardFilter lstScoreboard.ListItems(1).Text, True
        Else
            g_oGoddess.RemoveScoreboardFilter StripValueFromCommand(lstScoreboard.ListItems(1).Text), True
        End If
        lstScoreboard.ListItems.Remove (1)
    Next lIndex
    
    Log "All Scoreboard filter entries have been removed."
    
Exit_Properly:
    Exit Sub
    
cmdRemoveAllScoreboard_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdRemoveAllScoreboard_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdRemoveAllScoreboard_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdRemoveAllWatch_Click()

    'remove all entries from the scoreboard
    Dim lIndex As Long
    Dim lRet As Long
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdRemoveAllWatch_Click_Error
    
    lRet = MsgBox("Are you sure you want to remove all items from the Watch list?", vbQuestion + vbYesNo, "Confirm multiple removal")
    
    If lRet = vbNo Then
        Exit Sub
    End If
    
    For lIndex = 1 To lstWatch.ListItems.Count
        If LCase$(Trim$(lstWatch.ListItems(1).SubItems(1))) = "true" Then
            g_oGoddess.RemoveWatchFilter lstWatch.ListItems(1).Text, True
        Else
            g_oGoddess.RemoveWatchFilter StripValueFromCommand(lstWatch.ListItems(1).Text), True
        End If
        lstWatch.ListItems.Remove (1)
    Next lIndex
    
    Log "All Watch filter entries have been removed."
    
Exit_Properly:
    Exit Sub

cmdRemoveAllWatch_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdRemoveAllWatch_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdRemoveAllWatch_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdRemoveMessage_Click()

    'remove the selected message
    Dim lRet As Long
    Dim lIndex As Long
    Dim arDeletions() As String
    Dim lDelIndex As Long
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdRemoveMessage_Click_Error
    
    'have we a currently selected item?
    If MessagesSelCount > 0 Then
        'yes we do, find it and remove it
        'the key in the collection IS the message
        lRet = MsgBox("Are you sure you want to remove the selected Messages?", vbQuestion + vbYesNo, "Confirm delete")
        
        If lRet = vbNo Then
            GoTo Exit_Properly
        End If
        
        lDelIndex = 0
        
        'retrieve all the selected messages
        For lIndex = 1 To lstMessages.ListItems.Count
            If lstMessages.ListItems(lIndex).Selected Then
                ReDim Preserve arDeletions(lDelIndex)
                arDeletions(lDelIndex) = lstMessages.ListItems(lIndex).Text
                lDelIndex = lDelIndex + 1
            End If
        Next lIndex
        For lDelIndex = 0 To UBound(arDeletions)
            For lIndex = 1 To lstMessages.ListItems.Count
                If lstMessages.ListItems(lIndex).Text = arDeletions(lDelIndex) Then
                    If LCase$(lstMessages.ListItems(lIndex).SubItems(1)) = "true" Then
                        g_oGoddess.RemoveMessage lstMessages.ListItems(lIndex).Text, True
                    Else
                        g_oGoddess.RemoveMessage StripValueFromCommand(lstMessages.ListItems(lIndex).Text), True
                    End If
                    Log "The following message was removed from the Message List - " & lstMessages.ListItems(lIndex).Text & "."
                    'remove in the list
                    lstMessages.ListItems.Remove lIndex
                    Exit For
                End If
            Next lIndex
        Next lDelIndex
        'finito
    End If
    
Exit_Properly:
    Exit Sub
    
cmdRemoveMessage_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdRemoveMessage_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdRemoveMessage_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdRemoveScoreboardMsg_Click()

    'remove the selected message
    Dim lRet As Long
    Dim lIndex As Long
    Dim arDeletions() As String
    Dim lDelIndex As Long
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdRemoveScoreboardMsg_Click_Error
    
    'have we a currently selected item?
    If ScoreboardSelCount > 0 Then
        'yes we do, find it and remove it
        'the key in the collection IS the message
        lRet = MsgBox("Are you sure you want to remove the selected Messages?", vbQuestion + vbYesNo, "Confirm delete")
        
        If lRet = vbNo Then
            Exit Sub
        End If
        
        lDelIndex = 0
        
        'retrieve all the selected messages
        For lIndex = 1 To lstScoreboard.ListItems.Count
            If lstScoreboard.ListItems(lIndex).Selected Then
                ReDim Preserve arDeletions(lDelIndex)
                arDeletions(lDelIndex) = Trim$(lstScoreboard.ListItems(lIndex).Text)
                lDelIndex = lDelIndex + 1
            End If
        Next lIndex
        For lDelIndex = 0 To UBound(arDeletions)
            For lIndex = 1 To lstScoreboard.ListItems.Count
                If Trim$(lstScoreboard.ListItems(lIndex).Text) = arDeletions(lDelIndex) Then
                    If LCase$(Trim$(lstScoreboard.ListItems(lIndex).SubItems(1))) = "true" Then
                        If LCase$(Trim$(lstScoreboard.ListItems(lIndex).SubItems(2))) = "true" Then
                            g_oGoddess.RemoveScoreboardFilter StripValueFromCommand(lstScoreboard.ListItems(lIndex).Text), True
                        Else
                            g_oGoddess.RemoveScoreboardFilter lstScoreboard.ListItems(lIndex).Text, True
                        End If
                    Else
                        g_oGoddess.RemoveScoreboardFilter StripValueFromCommand(lstScoreboard.ListItems(lIndex).Text), True
                    End If
                    Log "The following message was removed from the Scoreboard filter - " & lstScoreboard.ListItems(lIndex).Text & "."
                    'remove in the list
                    lstScoreboard.ListItems.Remove lIndex
                    Exit For
                End If
            Next lIndex
        Next lDelIndex
        'finito
    End If
    
    ClearSelections lstScoreboard, xtScoreboard
    
Exit_Properly:
    Exit Sub
    
cmdRemoveScoreboardMsg_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdRemoveScoreboardMsg_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdRemoveScoreboardMsg_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdRemoveSelected_Click()

    Dim lRet As Long
    Dim sRemoteIP As String
    Dim lRemotePort As Long
    Dim sName As String
    
    'is there a selection?
    If lstConnected.SelCount = 0 Then
        MsgBox "Please select a server to remove from the list box!", vbExclamation + vbOKOnly, "No server selected"
        Exit Sub
    End If
    
    
    lRet = MsgBox("Are you sure you want to remove this game server?", vbQuestion + vbYesNo, "Confirm server remove")
    If lRet = vbNo Then
        Exit Sub
    End If
    
    'we need to remove this item from the connected list...
    '... and the internal structure
    'first get the ip and port
    lRet = GetPausedIPAndPort(lstConnected.List(lstConnected.ListIndex), sRemoteIP, lRemotePort)

    'now remove internally
    sName = sRemoteIP & lRemotePort
    moUDPListener.GameServers.Remove (sName)
    
    'is this item paused?
    If IsServerPaused(sRemoteIP, lRemotePort) Then
        'unpause it (ie remove frompause list ready for any further incoming packets
        'from this server
        UnpauseServer sRemoteIP, lRemotePort
    End If
    
    'and remove from connected list
    lstConnected.RemoveItem (lstConnected.ListIndex)
        
    Log "Removed GameServer at IP '" & sRemoteIP & "' on port '" & lRemotePort & "'"
    
End Sub

Private Sub cmdRemoveValidGameServer_Click()

    Dim oItem As ListItem
    Dim bFound As Boolean
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    Dim lRet As Long
    
    On Error GoTo cmdRemoveValidGameServer_Click_Error
    
    For Each oItem In lstValidGameServers.ListItems
        'has the user selected a ranked server to delete
        If oItem.Selected Then
            lRet = MsgBox("Are you sure you want to remove Ranked Server '" & _
            Trim$(oItem.Text) & "'?", vbQuestion + vbYesNo, "Remove Ranked Server")
            If lRet = vbNo Then
                bFound = False
                Exit For
            End If
            g_oGoddess.RemoveValidServer Trim$(oItem.SubItems(1)), Val(Trim$(oItem.SubItems(2))), True
            lstValidGameServers.ListItems.Remove oItem.Index
            bFound = True
            Exit For
        End If
    Next oItem
    
    If bFound Then
        MsgBox "Ranked Server has been successfully removed!", vbExclamation + vbOKOnly, "Removed Ranked Server"
    End If
    
Exit_Properly:
    Set oItem = Nothing
    Exit Sub
    
cmdRemoveValidGameServer_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdRemoveValidGameServer_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdRemoveValidGameServer_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub cmdRemoveWatch_Click()

    'remove the selected message
    Dim lRet As Long
    Dim lIndex As Long
    Dim arDeletions() As String
    Dim lDelIndex As Long
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo cmdRemoveWatch_Click_Error
    
    'have we a currently selected item?
    If WatchSelCount > 0 Then
        'yes we do, find it and remove it
        'the key in the collection IS the message
        lRet = MsgBox("Are you sure you want to remove the selected Messages?", vbQuestion + vbYesNo, "Confirm delete")
        
        If lRet = vbNo Then
            Exit Sub
        End If
        
        lDelIndex = 0
        
        'retrieve all the selected messages
        For lIndex = 1 To lstWatch.ListItems.Count
            If lstWatch.ListItems(lIndex).Selected Then
                ReDim Preserve arDeletions(lDelIndex)
                arDeletions(lDelIndex) = Trim$(lstWatch.ListItems(lIndex).Text)
                lDelIndex = lDelIndex + 1
            End If
        Next lIndex
        For lDelIndex = 0 To UBound(arDeletions)
            For lIndex = 1 To lstWatch.ListItems.Count
                If Trim$(lstWatch.ListItems(lIndex).Text) = arDeletions(lDelIndex) Then
                    If LCase$(Trim$(lstWatch.ListItems(lIndex).SubItems(1))) = "true" Then
                        If LCase$(Trim$(lstWatch.ListItems(lIndex).SubItems(2))) = "true" Then
                            g_oGoddess.RemoveWatchFilter StripValueFromCommand(lstWatch.ListItems(lIndex).Text), True
                        Else
                            g_oGoddess.RemoveWatchFilter lstWatch.ListItems(lIndex).Text, True
                        End If
                    Else
                        g_oGoddess.RemoveWatchFilter StripValueFromCommand(lstWatch.ListItems(lIndex).Text), True
                    End If
                    Log "The following message was removed from the Watch filters - " & lstWatch.ListItems(lIndex).Text & "."
                    'remove in the list
                    lstWatch.ListItems.Remove lIndex
                    Exit For
                End If
            Next lIndex
        Next lDelIndex
        'finito
    End If
    
    ClearSelections lstWatch, xtWatch
    
Exit_Properly:
    Exit Sub
    
cmdRemoveWatch_Click_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:cmdRemoveWatch_Click:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:cmdRemoveWatch_Click - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Sub

Private Sub Form_Load()

    Dim oSup As cSupport
    
    'set up GODDESS
    Set g_oGoddess = New CXMLFuncs
        
    'initialise from file
    If Not LoadGoddessXML() Then
        MsgBox "Unable to initialise Goddess!", vbCritical + vbOKOnly, "Critical load error"
        End
    End If
    
    'create the UDP listener
    Set moUDPListener = New CUDPListener
    
    With moUDPListener
        Set .Socket = Winsock1
        .LocalIP = g_oGoddess.SupportedProtocols("udp").LocalHost
        .LocalPort = g_oGoddess.SupportedProtocols("udp").ListenPort
    End With
    
    Set oSup = New cSupport
    oSup.IsListen
    
    Set mcolPaused = New Collection
    
    UpdateUDPDetails
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set g_oGoddess = Nothing
    
    Set mcolPaused = Nothing
    
    If Not moUDPListener Is Nothing Then
        moUDPListener.UDPClose
        Set moUDPListener = Nothing
    End If
    
    'are we logging
    If g_bIsLogging Then
        Log "GODDESS is shutting down."
        StopLogging
    End If
        
End Sub

Private Sub lstMessages_KeyDown(KeyCode As Integer, Shift As Integer)

    'if we have a delete key, click it
    If KeyCode = 46 Then
        cmdRemoveMessage_Click
    End If

End Sub

Private Sub lstScoreboard_KeyDown(KeyCode As Integer, Shift As Integer)

    'if we have a delete key, click it
    If KeyCode = 46 Then
        cmdRemoveScoreboardMsg_Click
    End If

End Sub

Private Sub lstWatch_KeyDown(KeyCode As Integer, Shift As Integer)

    'if we have a delete key, click it
    If KeyCode = 46 Then
        cmdRemoveWatch_Click
    End If

End Sub

Private Sub moUDPListener_AddedServerPacket()

    UpdateUDPDetails
    
End Sub

Private Sub moUDPListener_DataPacketReceived(ByVal RemoteIP As String, ByVal RemotePort As String, ByVal DateRcv As Date, ByVal TimeRcv As Date, ByVal BytesTotal As Long, ByVal PacketData As String)

    Dim oItem As ListItem
    Dim bRet As Boolean
    
    'ignore this packet if this server is currently paused
    bRet = IsServerPaused(RemoteIP, RemotePort)
    If bRet Then
        Exit Sub
    End If
    
    'is this a valid packet that we are looking for?
    
    Set oItem = lstData.ListItems.Add
   
    With oItem
        .Selected = True
        .EnsureVisible
        .Text = DateRcv
        .SubItems(1) = TimeRcv
        .SubItems(2) = RemoteIP
        .SubItems(3) = RemotePort
        .SubItems(4) = Trim$(StripChrZerosFromPacket(PacketData))
        .SubItems(5) = BytesTotal
    End With
    
    'has the packet come from a server who has not sent a packet in this session
    If Not moUDPListener.GameServers.IsGameServerConnected(RemoteIP, RemotePort) Then
        'we need to add it to the list
        lstConnected.AddItem Trim$(RemoteIP & ":" & RemotePort)
        lstConnected.ItemData(lstConnected.NewIndex) = RemotePort
    End If
       
End Sub

Private Sub moUDPListener_FTPStateChanged(ByVal StateString As String)

    'add a new line to the ftp list
    Dim oItem As ListItem
    Dim lErrno As Long
    Dim sDesc As String, sSource As String
    
    On Error GoTo moUDPListener_FTPStateChanged_Error
    
    Set oItem = lstFTP.ListItems.Add
    
    If Not oItem Is Nothing Then
        'add the item
        With oItem
            .Text = Date
            .SubItems(1) = Time
            .SubItems(2) = StateString
        End With
    End If
    
Exit_Properly:
    Set oItem = Nothing
    Exit Sub
    
moUDPListener_FTPStateChanged_Error:
    lErrno = Err.Number
    sDesc = Err.Description
    sSource = Err.Source
    MsgBox "The following error has occured in GODDESS:" & vbCr & vbCr & _
    "Error number: " & lErrno & vbCr & _
    "Error description: " & sDesc & vbCr & _
    "Error source: " & sSource, vbExclamation + vbOKOnly, "GODDESS Error"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:moUDPListener_FTPStateChanged - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly
    
End Sub

Private Sub moUDPListener_GameStatsFileSent(ByVal RemoteHostIP As String, ByVal RemoteHostPort As Long)

    'add a new line to the ftp list
    Dim oItem As ListItem
    Dim lErrno As Long
    Dim sDesc As String, sSource As String
    
    On Error GoTo moUDPListener_GameStatsFileSent_Error
    
    Set oItem = lstFTP.ListItems.Add
    
    If Not oItem Is Nothing Then
        'add the item
        With oItem
            .Text = Date
            .SubItems(1) = Time
            .SubItems(2) = "FTP file successfully sent for remote IP '" & RemoteHostIP & _
            "' on port '" & RemoteHostPort & "'"
        End With
    End If
    
    'refresh displays
    RefreshConnectedServers
    
Exit_Properly:
    Set oItem = Nothing
    Exit Sub
    
moUDPListener_GameStatsFileSent_Error:
    lErrno = Err.Number
    sDesc = Err.Description
    sSource = Err.Source
    MsgBox "The following error has occured in GODDESS:" & vbCr & vbCr & _
    "Error number: " & lErrno & vbCr & _
    "Error description: " & sDesc & vbCr & _
    "Error source: " & sSource, vbExclamation + vbOKOnly, "GODDESS Error"
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:moUDPListener_GameStatsFileSent - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly
    
End Sub

Private Sub moUDPListener_UDPError(ByVal ErrNumber As Long, ByVal ErrDescription As String, ByVal ErrSource As String, ByVal UDPData As String)

    If Trim$(UDPData) <> "" Then
        'log a dropped packet here
        Log "**************************************"
        Log "Dropped packet due to error in GODDESS"
        Log "**************************************"
        Log "   Error number:   " & ErrNumber
        Log "   Description:    " & ErrDescription
        Log "   Source:         " & ErrSource
        Log "   UDP data:       " & Trim$(UDPData)
        Log "**************************************"
    Else
        MsgBox "The following error has occured in the UDP Listener:" & vbCr & vbCr & _
        "Number: " & ErrNumber & vbCr & vbCr & _
        "Description: " & ErrDescription & vbCr & vbCr & _
        "Source: " & ErrSource, vbCritical + vbOKOnly, "UDP Error"
        Log "**************************************"
        Log "Error occured in the UDP Listener"
        Log "**************************************"
        Log "   Error number:   " & ErrNumber
        Log "   Description:    " & ErrDescription
        Log "   Source:         " & ErrSource
        Log "**************************************"
    End If
    
End Sub

Private Sub txtFTPHostName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdFTPHostName_Click
    End If
    
End Sub

Private Sub txtFTPHostName_LostFocus()

    Dim oCtrl As Control
    'only do this if the activecontrol is not the associated command button
    Set oCtrl = Me.ActiveControl
    If oCtrl.Name = "cmdFTPHostName" Then
        Exit Sub
    End If
    
    'has data changed in the textbox
    If Trim$(LCase$(txtFTPHostName.Text)) <> Trim$(LCase$(g_oGoddess.FTPRemoteHost)) Then
        'yep they want to change
        cmdFTPHostName_Click
    End If
    
End Sub

Private Sub txtFTPPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdFTPPassword_Click
    End If
    
End Sub

Private Sub txtFTPPassword_LostFocus()

    Dim oCtrl As Control
    'only do this if the activecontrol is not the associated command button
    Set oCtrl = Me.ActiveControl
    If oCtrl.Name = "cmdFTPPassword" Then
        Exit Sub
    End If
    
    'has data changed in the textbox
    If Trim$(LCase$(txtFTPPassword.Text)) <> Trim$(LCase$(g_oGoddess.FTPPassword)) Then
        'yep they want to change
        cmdFTPPassword_Click
    End If
    
End Sub

Private Sub txtFTPUsername_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdFTPUsername_Click
    End If
    
End Sub

Private Sub txtFTPUsername_LostFocus()

    Dim oCtrl As Control
    'only do this if the activecontrol is not the associated command button
    Set oCtrl = Me.ActiveControl
    If oCtrl.Name = "cmdFTPUsername" Then
        Exit Sub
    End If
    
    'has data changed in the textbox
    If Trim$(LCase$(txtFTPUsername.Text)) <> Trim$(LCase$(g_oGoddess.FTPUserName)) Then
        'yep they want to change
        cmdFTPUsername_Click
    End If
    
End Sub

Private Sub txtListenPortChange_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdListenPortChange_Click
    End If
    
End Sub

Private Sub txtLocalIPChange_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdChangeIP_Click
    End If
    
End Sub

Private Sub txtNewMsg_KeyPress(KeyAscii As Integer)

    'if the key is a 'enter' key, then do the click
    If KeyAscii = 13 Then
        cmdAddMsg_Click
    End If
    
End Sub

Private Sub txtValueScoreboard_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdAddValueScoreboard_Click
    End If
    
End Sub

Private Sub txtValueWatch_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdAddValueWatch_Click
    End If
    
End Sub

Private Function LoadGoddessXML() As Boolean

    Dim lErrno As Long
    Dim sSource As String, sDesc As String

    On Error GoTo LoadGoddessXML_Error
    
    'set global logging flag in case of error with logging file
    g_bIsLogging = False
    
    With g_oGoddess
        If Not .LoadXML() Then
            MsgBox "Unable to load Goddess XML file!", vbCritical + vbOKOnly, "Failed to load Goddess XML"
            LoadGoddessXML = False
            GoTo Exit_Properly
        End If
        
        'first get the logging so that if enabled we can start straight away
        .GetLogging
        If .LoggingEnabled Then
            StartLogging
            Log "GODDESS logging initiated."
        End If
        
        'good, loaded the xml, now get info
        .GetVersion
        
        'now the protocols
        .GetSupportedProtocols
        
        'message file paths for Scoreboard, Watch, Message
        .GetMessageFilePaths
        
        'FTP information
        .GetFTPInfo
        
        'load the xml for each other category
        If Not .LoadScoreboardXML() Then
            MsgBox "Unable to load Scoreboard XML file!", vbCritical + vbOKOnly, "Failed to load Scoreboard XML"
            LoadGoddessXML = False
            GoTo Exit_Properly
        End If
        If Not .LoadWatchXML() Then
            MsgBox "Unable to load Watch XML file!", vbCritical + vbOKOnly, "Failed to load Watch XML"
            LoadGoddessXML = False
            GoTo Exit_Properly
        End If
         If Not .LoadMessageXML() Then
            MsgBox "Unable to load Message XML file!", vbCritical + vbOKOnly, "Failed to load Message XML"
            LoadGoddessXML = False
            GoTo Exit_Properly
        End If
       
        'Message filters
        .GetFilters
        
        'and populate the listbox
        PopulateFilters
        
        'total messages
        .GetMessages
        
        'populate the message listbox
        PopulateMessages
        
        'valid game servers
        .GetValidGameServers
        
        PopulateValidGameServers
        
        'now enter in the pathing info
        txtMessageFileScoreboard.Text = .MessageFiles("Scoreboard").FilePath
        txtMessageFileScoreboard.ToolTipText = .MessageFiles("Scoreboard").FilePath
      
        txtMessageFileWatch.Text = .MessageFiles("Watch").FilePath
        txtMessageFileWatch.ToolTipText = .MessageFiles("Watch").FilePath
        
        txtMessageFileMessage.Text = .MessageFiles("Message").FilePath
        txtMessageFileMessage.ToolTipText = .MessageFiles("Message").FilePath
        
        'and the logging details
        txtLoggingPath.Text = .LogFilePath
        If .LoggingEnabled Then
            chkLoggingEnabled.Value = vbChecked
        Else
            chkLoggingEnabled.Value = vbUnchecked
        End If
            
        txtFTPFilePath.Text = .FTPFilePath
        txtFTPHostName.Text = .FTPRemoteHost
        txtFTPUsername.Text = .FTPUserName
        txtFTPPassword.Text = .FTPPassword
        
    End With
    
    LoadGoddessXML = True
    
Exit_Properly:
    Exit Function
    
LoadGoddessXML_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:LoadGoddessXML:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    LoadGoddessXML = False
    Log "******************************************************"
    Log "Error occured in GODDESS"
    Log "******************************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:LoadGoddessXML - " & sSource
    Log "******************************************************"
    GoTo Exit_Properly

End Function

Private Function PopulateFilters()

    Dim oFilter As CFilter
    Dim oItem As ListItem
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo PopulateFilters_Error
    
    lstScoreboard.ListItems.Clear
    lstScoreboard.MultiSelect = False
    lstWatch.ListItems.Clear
    lstWatch.MultiSelect = False
    
    For Each oFilter In g_oGoddess.Filters
        With oFilter
            'what type of filter is this?
            If LCase$(Trim$(oFilter.FilterType)) = "scoreboard" Then
                Set oItem = lstScoreboard.ListItems.Add
                oItem.Selected = True
                oItem.EnsureVisible
            Else
                Set oItem = lstWatch.ListItems.Add
                oItem.Selected = True
                oItem.EnsureVisible
            End If
            If oFilter.CommandEvent Then
                If oFilter.CommandValue Then
                    If Trim$(oFilter.FilterValue) <> "" Then
                        oItem.Text = oFilter.FilterEvent & " '" & oFilter.FilterValue & "'"
                    Else
                        oItem.Text = oFilter.FilterEvent
                    End If
                Else
                    oItem.Text = oFilter.FilterEvent
                End If
            Else
                If Trim$(oFilter.FilterValue) <> "" Then
                    oItem.Text = oFilter.FilterEvent & " '" & oFilter.FilterValue & "'"
                Else
                    oItem.Text = oFilter.FilterEvent
                End If
            End If
            'and the commandevent
            If oFilter.CommandEvent Then
                oItem.SubItems(1) = "True"
            End If
            If oFilter.CommandValue Then
                oItem.SubItems(2) = "True"
            End If
        End With
    Next oFilter
    
    lstScoreboard.MultiSelect = True
    lstWatch.MultiSelect = True
    
    PopulateFilters = True
    
Exit_Properly:
    Set oFilter = Nothing
    Set oItem = Nothing
    Exit Function
    
PopulateFilters_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:PopulateFilters:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    PopulateFilters = False
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:PopulateFilters - " & sSource
    Log "**************************************"
    GoTo Exit_Properly

End Function

Private Function PopulateMessages() As Boolean

    Dim oMsg As CMessage
    Dim oItem As ListItem
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo PopulateMessages_Error
    
    'turn off multiselect temporarily
    lstMessages.MultiSelect = False
    
    For Each oMsg In g_oGoddess.Messages
        With oMsg
            Set oItem = lstMessages.ListItems.Add
            oItem.Selected = True
            oItem.EnsureVisible
            oItem.Text = oMsg.MessageName
            If oMsg.CommandEvent Then
                oItem.SubItems(1) = "True"
            End If
            If oMsg.CommandValue Then
                oItem.SubItems(2) = "True"
            End If
        End With
    Next oMsg
   
    'now turn multiselect back on
    lstMessages.MultiSelect = True
   
    PopulateMessages = True
    
Exit_Properly:
    Set oMsg = Nothing
    Exit Function
    
PopulateMessages_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:PopulateMessages:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    PopulateMessages = False
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:PopulateMessages - " & sSource
    Log "**************************************"
    GoTo Exit_Properly

End Function

Private Function StripValueFromCommand(ByVal sCommand As String) As String

    Dim lPos As Long
    Dim sEvent As String
    
    lPos = InStr(sCommand, "'")
    
    If lPos > 1 Then
        StripValueFromCommand = Trim$(Mid$(sCommand, 1, lPos - 2))
    Else
        StripValueFromCommand = Trim$(sCommand)
    End If
    
End Function

Private Function IsItemInList(ByRef oList As ListView, ByVal sEvent As String) As Boolean

    Dim lIndex As Long
    Dim sName As String
    
    On Error GoTo IsItemInList_Error
    
    IsItemInList = False
    
    For lIndex = 1 To oList.ListItems.Count
        If InStr(oList.ListItems(lIndex).Text, sEvent) Then
            IsItemInList = True
            Exit For
        End If
    Next lIndex
    
    
Exit_Properly:
    Exit Function
    
IsItemInList_Error:
    'simply default to IS in the list
    IsItemInList = True
    GoTo Exit_Properly
    
End Function

Private Sub ClearSelections(ByRef oList As Object, ByVal ListType As XML_TYPE)

    Dim oBox As ListBox
    Dim oView As ListView
    
    'clear all selections in the message list
    Dim lIndex As Long
    Dim lCount As Long
    
    If LCase$(TypeName(oList)) = "listbox" Then
        Set oBox = oList
        For lIndex = 0 To oBox.ListCount - 1
            oBox.Selected(lIndex) = False
        Next lIndex
    Else
        Set oView = oList
        For lIndex = 1 To oView.ListItems.Count
            oView.ListItems(lIndex).Selected = False
        Next lIndex
    End If
    
End Sub

Private Function StripChrZerosFromPacket(ByVal PacketData As String) As String

    Dim lIndex As Long
    Dim sStrip As String
    
    For lIndex = 1 To Len(PacketData)
        If Not Asc(Mid$(PacketData, lIndex, 1)) = 0 Then
            sStrip = sStrip & Mid$(PacketData, lIndex, 1)
        End If
    Next lIndex
    
    StripChrZerosFromPacket = sStrip
    
End Function

Private Function IsServerPaused(ByVal RemoteIP As String, ByVal RemotePort As Long) As Boolean

    Dim sName As String
    Dim sRemoteIP As String
    Dim lRemotePort As Long
    Dim bRet As Boolean
    Dim lIndex As Long
    
    IsServerPaused = False
    
    For lIndex = 1 To mcolPaused.Count
        sName = mcolPaused(lIndex)
        bRet = GetPausedIPAndPort(sName, sRemoteIP, lRemotePort)
        If RemoteIP = sRemoteIP And RemotePort = lRemotePort Then
            'we have a match, therefore it is true, it is paused
            IsServerPaused = True
            Exit Function
        End If
    Next lIndex
    
End Function

Private Sub PauseServer(ByVal RemoteIP As String, ByVal RemotePort As Long)

    Dim sName As String
    
    sName = RemoteIP & ":" & RemotePort
    
    mcolPaused.Add sName, sName
    
    'the selected item can now be marked as paused
    sName = sName & " (Paused)"
    
    lstConnected.List(lstConnected.ListIndex) = sName
    
End Sub

Private Sub UnpauseServer(ByVal RemoteIP As String, ByVal RemotePort As Long)

    Dim sName As String
    
    sName = RemoteIP & ":" & RemotePort
    
    mcolPaused.Remove sName
    

    'we can now remove the pause indicator
    lstConnected.List(lstConnected.ListIndex) = sName
    
End Sub

Private Function GetPausedIPAndPort(ByVal ServerIPPort As String, ByRef RemoteIP As String, ByRef RemotePort As Long) As Boolean

    Dim lPos As Long
    Dim lPausePos As Long
    
    lPos = InStr(ServerIPPort, ":")
    lPausePos = InStr(ServerIPPort, " ")
    
    RemoteIP = Mid$(ServerIPPort, 1, lPos - 1)
    
    If lPausePos = 0 Then
        RemotePort = Val(Mid$(ServerIPPort, lPos + 1))
    Else
        RemotePort = Val(Mid$(ServerIPPort, lPos + 1, ((lPausePos) - (lPos + 1))))
    End If
    
End Function

Private Sub UpdateUDPDetails()

    'version
    'protocol
    'packet length
    'localip
    'listenport
    Dim oItem As ListItem
    Dim oSvr As CGameServer
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo UpdateUDPDetails_Error
    
    With g_oGoddess
        lblDetsVersion.Caption = .Version
        'only 1 supported protocol at this time
        lblDetsProtocol.Caption = .SupportedProtocols("udp").ProtocolName
        lblDetsPacketLength.Caption = .SupportedProtocols("udp").PacketLength
        lblDetsLocalIP.Caption = .SupportedProtocols("udp").LocalHost
        lblDetsListenPort.Caption = .SupportedProtocols("udp").ListenPort
    End With
    
    'now fill in the game servers list view
    lstGameServers.ListItems.Clear
    
    For Each oSvr In moUDPListener.GameServers
        Set oItem = lstGameServers.ListItems.Add
        If Not oItem Is Nothing Then
            oItem.Text = oSvr.RemoteHostIP & ":" & oSvr.RemotePort
            oItem.SubItems(1) = oSvr.Packets.Count
        End If
    Next oSvr
        
Exit_Properly:
    Set oItem = Nothing
    Set oSvr = Nothing
    Exit Sub
    
UpdateUDPDetails_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    MsgBox "The following error has occured in frmMain:UpdateUDPDetails:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Description: " & Err.Description, vbExclamation + vbOKOnly, "GODDESS Error"
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:UpdateUDPDetails - " & sSource
    Log "**************************************"
    GoTo Exit_Properly
    
End Sub

Private Function MessagesSelCount() As Long

    Dim lIndex As Long
    Dim lCount As Long
    
    lIndex = 0
    
    For lIndex = 1 To lstMessages.ListItems.Count
        If lstMessages.ListItems(lIndex).Selected Then
            lCount = lCount + 1
        End If
    Next lIndex
    
    MessagesSelCount = lCount
    
End Function

Private Function ScoreboardSelCount() As Long

    Dim lIndex As Long
    Dim lCount As Long
    
    lIndex = 0
    
    For lIndex = 1 To lstScoreboard.ListItems.Count
        If lstScoreboard.ListItems(lIndex).Selected Then
            lCount = lCount + 1
        End If
    Next lIndex
    
    ScoreboardSelCount = lCount
    
End Function

Private Function WatchSelCount() As Long

    Dim lIndex As Long
    Dim lCount As Long
    
    lIndex = 0
    
    For lIndex = 1 To lstWatch.ListItems.Count
        If lstWatch.ListItems(lIndex).Selected Then
            lCount = lCount + 1
        End If
    Next lIndex
    
    WatchSelCount = lCount
    
End Function

Private Sub RefreshConnectedServers()

    Dim oSvr As CGameServer
    
    lstConnected.Clear
    lstGameServers.ListItems.Clear
    
    For Each oSvr In moUDPListener.GameServers
        lstConnected.AddItem oSvr.RemoteHostIP & ":" & oSvr.RemotePort
    Next oSvr
    
    UpdateUDPDetails
    
End Sub

Private Function PopulateValidGameServers() As Boolean

    Dim oSvr As CGameServer
    Dim oItem As ListItem
    Dim lErrno As Long
    Dim sSource As String, sDesc As String
    
    On Error GoTo PopulateValidGameServers_Error
    
    lstValidGameServers.ListItems.Clear
    
    For Each oSvr In g_oGoddess.ValidGameServers
        With oSvr
            Set oItem = lstValidGameServers.ListItems.Add
            oItem.Selected = True
            oItem.EnsureVisible
            oItem.Text = oSvr.RemoteHostName
            oItem.SubItems(1) = oSvr.RemoteHostIP
            oItem.SubItems(2) = oSvr.RemotePort
        End With
    Next oSvr
   
    PopulateValidGameServers = True
    
Exit_Properly:
    Set oSvr = Nothing
    Exit Function
    
PopulateValidGameServers_Error:
    lErrno = Err.Number
    sSource = Err.Source
    sDesc = Err.Description
    PopulateValidGameServers = False
    MsgBox "The following error has occured in frmMain:PopulateValidGameServers:" & vbCrLf & vbCrLf & _
    "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly, "Goddess - Error occured"
    Log "**************************************"
    Log "Error occured in GODDESS"
    Log "**************************************"
    Log "   Error number:   " & lErrno
    Log "   Description:    " & sDesc
    Log "   Source:         frmMain:PopulateValidGameServers - " & sSource
    Log "**************************************"
    GoTo Exit_Properly

End Function

Private Function IsValidServerInList(ByVal ServerIP As String, ByVal ServerPort As Long) As Boolean

    Dim lIndex As Long
    Dim sName As String
    
    On Error GoTo IsValidServerInList_Error
    
    IsValidServerInList = False
    
    For lIndex = 1 To lstValidGameServers.ListItems.Count
        If InStr(lstValidGameServers.ListItems(lIndex).SubItems(1), ServerIP) Then
            'found the ip, what about the port
            If lstValidGameServers.ListItems(lIndex).SubItems(2) = ServerPort Then
                'yep, we found the bugger
                IsValidServerInList = True
                Exit For
            End If
        End If
    Next lIndex
    
    
Exit_Properly:
    Exit Function
    
IsValidServerInList_Error:
    'simply default to IS in the list
    IsValidServerInList = True
    GoTo Exit_Properly
    
End Function

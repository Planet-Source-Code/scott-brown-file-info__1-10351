VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFileInfo 
   Caption         =   "Basic File Information"
   ClientHeight    =   5265
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12480
   Icon            =   "frmFileinfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   2400
      TabIndex        =   8
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin ComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5010
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   19024
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdRtrnFileInfo 
      Caption         =   "&Get File Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "List View State"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   2295
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin ComctlLib.ImageList img2 
      Left            =   3240
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFileinfo.frx":0442
            Key             =   "fldr"
            Object.Tag             =   "fldr"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFileinfo.frx":075C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList img1 
      Left            =   2520
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFileinfo.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFileinfo.frx":0D90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuGet 
         Caption         =   "&Get File Info"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "&Options"
      Begin VB.Menu mnuLog 
         Caption         =   "Lo&gging"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSize 
         Caption         =   "&Sizing"
         Begin VB.Menu mnuSizeCols 
            Caption         =   "Size Cols to Data"
            Index           =   0
         End
         Begin VB.Menu mnuSizeCols 
            Caption         =   "Size Cols to Headers"
            Index           =   1
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View"
         Begin VB.Menu mnuLvwView 
            Caption         =   "&Icon"
            Index           =   0
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuLvwView 
            Caption         =   "&Small Icon"
            Index           =   1
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuLvwView 
            Caption         =   "&List"
            Index           =   2
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuLvwView 
            Caption         =   "&Report"
            Index           =   3
            Shortcut        =   ^R
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objFileProp As CFileInfo
Private sCurPath As String
Dim iFldr As Integer
Dim iFile As Integer
Dim lheight As Long
Dim lwidth As Long

Function RecurseFolderList(foldername)
    Dim fso, fldr, sfldrs, sfldr, fils, fil
        
    'get a handle to the FileSystemObject from scrrun.dll
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Locate required folder from selected folderlist
    Set fldr = fso.GetFolder(foldername)
    
    'Return handle to subfolders collection
    Set sfldrs = fldr.Subfolders
    
    'Return handle to files collection
    Set fils = fldr.Files
    
    'For each subfolder in the Folder
    For Each sfldr In sfldrs
        'Do something with the Folder Name
        shwAttr sfldr, "FOLDER"
        'Then recurse this function with the sub-folder to get any sub-folders ad infinitum
        RecurseFolderList (sfldr)
    Next
    
    'For each folder check for any files
    For Each fil In fils
        shwAttr fil, "FILE"
    Next
    
    'Reset Statusbar once finished
    sBar.Panels(1).Text = ""
    
End Function

Private Sub shwAttr(ByVal fsoObj As String, ByVal fsoObjTyp As String)
    Dim lngDBQ As Long
    Dim objFileProp As CFileInfo
    Dim itmX As ListItem
            
    'Return reference to new File Info class
    Set objFileProp = New CFileInfo
    
    With objFileProp
        'Instantiate class and find file
        lngDBQ = .FindFileInfo(fsoObj)
        
        'If file found
        If lngDBQ = 1 Then
            
            'Set variable equal to listview item and add using filename (Unique in PC!) as key
            Set itmX = ListView1.ListItems.Add(, CStr(.FileName), CStr(.FileName))          'Filename
                
                If (mnuLog.Checked = True) Then
                    'Notify user whats happening
                    sBar.Panels(1).Text = "Parsing " & .FileName
                    sBar.Panels(2).Text = "Folders found: " & iFldr & "  Files found: " & iFile
                End If
                
                'If a folder add a folder icon
                If fsoObjTyp = "FOLDER" Then
                    itmX.Icon = 1
                    itmX.SmallIcon = 1  ' Set an icon from ImageList2.
                    iFldr = iFldr + 1
                Else
                    'Else add a file icon
                    itmX.Icon = 2
                    itmX.SmallIcon = 2  ' Set an icon from ImageList2.
                    iFile = iFile + 1
                    
                    'Add a subitem for each of the file attributes retrieved from the FileInfo class
                    If Not IsNull(.FileName) Then
                        itmX.SubItems(1) = IIf(.mByte = " bytes", "0 Bytes", Format(.mByte, "###,###,###"))
                    End If
                            
                    If Not IsNull(.FileName) Then
                        itmX.SubItems(2) = Format(.CreationTime, "DD/MM/YYYY HH:MM:SS")
                    End If
                    
                    If Not IsNull(.FileName) Then
                        itmX.SubItems(3) = Format(.LastAccessTime, "DD/MM/YYYY HH:MM:SS")
                    End If
                    
                    If Not IsNull(.FileName) Then
                        itmX.SubItems(4) = Format(.LastWriteTime, "DD/MM/YYYY HH:MM:SS")
                    End If
                    
                    If Not IsNull(.FileName) Then
                        itmX.SubItems(5) = .ReadOnly
                    End If
                    
                End If
        End If
    End With

End Sub

Private Sub cmdRtrnFileInfo_Click()
    
    'Reset count variables and StatusBar content
    iFile = 0
    iFldr = 0
    sBar.Panels(2).Text = ""
    
    'If we've parsed this sub-dir already don't do it again, it might be enormous ?!
    If sCurPath = Dir1 Then
        Exit Sub
    Else
        'Set pointer to hourglass
        Screen.MousePointer = vbHourglass
        'Clear the listview
        Me.ListView1.ListItems.Clear
        'Call main function 'used to recurse through directory structure
        RecurseFolderList Dir1
        'Reset pointer once finished
        Screen.MousePointer = vbDefault
    End If
    
    sCurPath = Dir1
    
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    
    'Set initial path to wherever you unzipped me to
    Dir1.Path = App.Path
    lwidth = 12000
    lheight = 5955
    
    'Setup Form states at startup
    Me.ListView1.Left = 2350
    Me.ListView1.Top = 110
    Me.ListView1.Width = lwidth
    Me.ListView1.Height = lheight
    
    EnhListView_Add_AllowRepositioning ListView1, True
        
    'Add Column Headers for each of the attributes we want to show
    ListView1.ColumnHeaders.Add , "Filename", "Filename", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , "FileSize", "File Size", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , "Creation", "Creation Time", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , "Accessed", "Last Accessed", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , "Written", "Last Written", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , "ReadOnly", "Read Only", ListView1.Width / 6
    
    Option1(0).Caption = "Icon"
    Option1(1).Caption = "SmallIcon"
    Option1(2).Caption = "List"
    Option1(3).Caption = "Report"
    
    'Associate imageView controls with ListView Control
    ListView1.Icons = img2
    ListView1.SmallIcons = img1
    ListView1.View = lvwList
    
End Sub

Private Sub Form_Resize()

    If Me.Width < lwidth Then Me.Width = lwidth
    If Me.Height < lheight Then Me.Height = lheight
    
    'Resize the list box if you have a huge sub-directory structure
    Me.ListView1.Width = Me.Width - 2505
    Me.ListView1.Height = Me.Height - 1230
    
    'Dynamically resize controls on form Resize event
    Me.Dir1.Height = Me.Height - 3400
    Me.cmdRtrnFileInfo.Top = Dir1.Top + Dir1.Height + 100
    Me.Frame1.Top = cmdRtrnFileInfo.Top + cmdRtrnFileInfo.Height + 100
    
    EnhListView_ResizeColumnHeaders ListView1, True
        
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    
    'Call enhanced ListView function to enable sorting by columnheader
    EnhListView_SortColumns ListView1, ColumnHeader.Index, False

End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show vbModal, Me

End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuGet_Click()
    'just call btnClick event from menu
    cmdRtrnFileInfo_Click
End Sub

Private Sub mnuLog_Click()

mnuLog.Checked = Not (mnuLog.Checked)

End Sub

Private Sub mnuLvwView_Click(Index As Integer)
        
    'Set ListView state from menu
    ListView1.View = Index
    
    'Call setMenu function
    setLvwMenu Index
    
    'Set option buttons to match menu click
   Option1(Index).Value = True
End Sub

Private Sub mnuSizeCols_Click(Index As Integer)

    'Resize columns to either data in records or column headers by user selection
    Select Case Index
        Case 0
            EnhListView_ResizeColumns ListView1, True
        Case 1
            EnhListView_ResizeColumnCaptions ListView1, True
    End Select

End Sub

Private Sub Option1_Click(Index As Integer)
    
    'Set ListView state from menu
    ListView1.View = Index
    
    'Call setMenu function
    setLvwMenu Index
    
    'Set Menu States to match menu click
    mnuLvwView(Index).Checked = True
    
End Sub

Private Function setLvwMenu(ByVal Index As Integer)
    Dim i As Integer
        
    'Do some clever state checking to uncheck menus not selected/check menu selected
    For i = 0 To mnuLvwView.Count - 1
        mnuLvwView(i).Checked = (i = Index)
    Next i
        
End Function

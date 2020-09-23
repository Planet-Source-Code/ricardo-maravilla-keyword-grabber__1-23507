VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Caption         =   "Keyword Grabber"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEditGrid2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5085
      TabIndex        =   26
      Top             =   6240
      Width           =   3615
   End
   Begin VB.TextBox txtEditGrid 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4005
      TabIndex        =   12
      Top             =   2640
      Width           =   3615
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8520
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search"
      Height          =   2895
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtNumberToGet 
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Enter words to search for."
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search!"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Click here when you're ready to search."
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Num. of responses to get (1-100):"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblSearch 
         Caption         =   "Enter search terms:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraKeywords 
      Caption         =   "Keywords"
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   8895
      Begin VB.CommandButton cmdReset2 
         Caption         =   "Reset"
         Height          =   375
         Left            =   7320
         TabIndex        =   32
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chkEdit2 
         Caption         =   "Edit Cell"
         Height          =   255
         Left            =   7800
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdClearList 
         Caption         =   "Clear List"
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "Remove Item"
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdGenHTML 
         Caption         =   "Generate HTML Code"
         Height          =   375
         Left            =   5640
         TabIndex        =   22
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton cmdDeleteCell 
         Caption         =   "Delete Cell"
         Height          =   375
         Left            =   6120
         TabIndex        =   20
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddAllToMK 
         Caption         =   "Add all to My Keywords"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddToMK 
         Caption         =   "Add to MK"
         Height          =   375
         Left            =   3240
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid FlxKey 
         Height          =   2250
         Left            =   4800
         TabIndex        =   14
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3969
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedCols       =   0
         ScrollBars      =   2
         FormatString    =   "^Keywords"
      End
      Begin VB.ListBox lstKeywords 
         BackColor       =   &H00C0FFFF&
         Height          =   3180
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton cmdAddKeyword 
         Caption         =   "Add Keyword"
         Height          =   375
         Left            =   4800
         TabIndex        =   21
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Keywords"
         Height          =   255
         Left            =   6720
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblKeyNumber 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   6000
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "My Keywords:"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraWebpages 
      Caption         =   "Web Pages"
      Height          =   2895
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   3960
         TabIndex        =   31
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteCellURLs 
         Caption         =   "Delete Cell"
         Height          =   375
         Left            =   3960
         TabIndex        =   17
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CheckBox chkEdit 
         Caption         =   "Edit Cell"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdViewPage 
         Caption         =   "View Page"
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdGetKWs 
         Caption         =   "Get Keywords"
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid FlxURLs 
         Height          =   2250
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3969
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         ScrollBars      =   2
         SelectionMode   =   1
         FormatString    =   "^URL"
      End
   End
   Begin VB.Label DragLabel 
      Caption         =   "DragLabel"
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblStatus 
      Caption         =   "Ready"
      Height          =   255
      Left            =   4380
      TabIndex        =   11
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3660
      TabIndex        =   10
      Top             =   3120
      Width           =   615
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aString As String 'will hold downloaded text
Dim bLoading As Boolean 'to tell if form is loading
Dim GettingKeys As Boolean 'to tell if were getting keywords or URLs
Dim NumberToGet As Integer 'how many responses the user wants to get
Dim Keywords() As String 'array that will hold keywords before they're placed in the listbox
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit

'<%%%%%%%%%%%%%%%%%Code for the Form%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private Sub Form_Load()
    Dim K As Integer
    Dim sLine As String
    Dim Handle1 As Integer: Handle1 = FreeFile
    bLoading = True
    
    FlxURLs.ColWidth(0) = FlxURLs.Width - 98
    FlxKey.ColWidth(0) = FlxKey.Width - 98
    Show

    lblKeyNumber.Caption = "0"
    
    'open file that saves words from "My Keywords"
    Open App.Path & "\MYKWs.txt" For Input As #Handle1
    While Not EOF(Handle1)
        Line Input #Handle1, sLine
        AddCell FlxKey, sLine
    Wend
    Close #Handle1
    bLoading = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim K As Integer
    Dim Handle1 As Integer: Handle1 = FreeFile
    bLoading = True
    Open App.Path & "\MYKWs.txt" For Output As #Handle1
    For K = 1 To FlxKey.Rows - 1
        FlxKey.Row = K
        Print #Handle1, FlxKey.Text
    Next
    Close #Handle1
    bLoading = False
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "FlxKey" Then
        cmdDeleteCell_Click
    End If
End Sub
'%%%%%%%%%%%%%%%%%End of Code for the Form%%%%%%%%%%%%%%%%%%%%%%%%%%%%>

'<%%%%%%%%%%%%%%%%%%%%Code for all the Frames%%%%%%%%%%%%%%%%%
Private Sub fraKeywords_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "FlxKey" Then 'deletes item from flxurls
        cmdDeleteCell_Click                  'when it's dragged off
    ElseIf Me.ActiveControl.Name = "FlxURLs" Then
        cmdDeleteCellUrls_Click
    End If
End Sub

Private Sub fraSearch_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "FlxKey" Then
        cmdDeleteCell_Click
    ElseIf Me.ActiveControl.Name = "FlxURLs" Then
        cmdDeleteCellUrls_Click
    End If
End Sub

Private Sub fraWebpages_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "FlxKey" Then
        cmdDeleteCell_Click
    ElseIf Me.ActiveControl.Name = "FlxURLs" Then
        cmdDeleteCellUrls_Click
    End If
End Sub
'%%%%%%%%%%%%%%%%%%%%End of Code for all Frames%%%%%%%%%%%%%>

'<%%%%%%%%%%%%%%%%%Code for Search Box%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private Sub cmdSearch_Click()
    Dim SearchTerms As String
    Dim SearchStringP1 As String 'part 1
    Dim SearchStringP2 As String 'part 2
    Dim SearchString As String
    
    NumberToGet = Val(txtNumberToGet.Text)
    SearchStringP1 = "http://hotbot.lycos.com/?MT="
    SearchStringP2 = "&SM=MC&DV=0&LG=any&DC=100&DE=0&AM1=MC&x=66&y=16"
    
    If txtSearch.Text = vbNullString Then
        MsgBox "Please enter words to search for!", vbOKOnly + vbExclamation, "Keyword Grabber"
        Exit Sub
    ElseIf NumberToGet <= 0 Or NumberToGet > 100 Then
        MsgBox "Please enter a valid number of responses to get."
        Exit Sub
    End If
    
    SearchTerms = Trim(txtSearch.Text)
    'put a plus sign between all words
    SearchTerms = Replace(SearchTerms, " ", "+")
    'put the search terms in the URL
    SearchString = SearchStringP1 & SearchTerms & SearchStringP2
    GettingKeys = False
    Inet1.Execute SearchString
End Sub
'%%%%%%%%%%%%%%%%%%% End of code for Search Box%%%%%%%%%%%%%%%%%%%%%>

'<%%%%%%%%%%%%%%%%%Code for FlxURls%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private Sub FlxURLs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ih 'item height
    ih = FlxURLs.CellHeight
    DragLabel.Move FlxURLs.Left + fraWebpages.Left, FlxURLs.Top + _
                   fraWebpages.Top + FlxURLs.CellTop, FlxURLs.CellWidth, ih
    DragLabel.Drag 'drag the invisible label instead of the grid
End Sub

Private Sub FlxURLs_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "FlxKey" Then
        cmdDeleteCell_Click
    End If
End Sub

Private Sub FlxURLs_EnterCell()
    If FlxURLs.MouseRow = 0 Then Exit Sub
    With FlxURLs
        If bLoading = False Then
            'put cell contents in text box
            txtEditGrid.Text = .Text
            'move focus to text box if it's enabled
            txtEditGrid.Visible = True
            If txtEditGrid.Enabled Then txtEditGrid.SetFocus
        End If
    End With
End Sub

Private Sub FlxURLs_LeaveCell()
    If bLoading = False Then
        FlxURLs.Text = txtEditGrid.Text
    End If
End Sub

Private Sub chkEdit_Click()
    txtEditGrid.Enabled = Not txtEditGrid.Enabled
End Sub

Private Sub cmdGetKWs_Click()
    If txtEditGrid.Text = "" Or txtEditGrid.Text = "Empty" _
    Or InStr(1, txtEditGrid.Text, "ttp://", vbTextCompare) = 0 _
    Or InStr(1, txtEditGrid.Text, ".com", vbTextCompare) = 0 Then
        MsgBox "This is not a valid URL!", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    Inet1.Cancel
    GettingKeys = True 'must set this to true so that inet doesn't try to look for urls.
    Inet1.Execute FlxURLs.Text
End Sub

Private Sub cmdViewPage_Click()
    'This launches the user's default browser to view the URL.
    If txtEditGrid.Text <> "" Or txtEditGrid.Text <> "Empty" _
    And InStr(1, txtEditGrid.Text, "ttp://", vbTextCompare) <> 0 _
    And InStr(1, txtEditGrid.Text, ".com", vbTextCompare) <> 0 Then

    Call ShellExecute(0&, vbNullString, FlxURLs.Text, vbNullString, _
                      vbNullString, vbNormalFocus)
    End If
End Sub

Private Sub cmdDeleteCellUrls_Click()
    bLoading = True
    If FlxURLs.Rows >= 3 Then
        FlxURLs.RemoveItem FlxURLs.Row
        ColorCells FlxURLs
        FlxURLs.Row = 1
        txtEditGrid.Text = FlxURLs.Text
    ElseIf FlxURLs.Rows = 2 Then
        FlxURLs.Rows = 1
        txtEditGrid.Text = vbNullString
    End If
    bLoading = False
End Sub

Private Sub cmdReset_Click()
    FlxURLs.Rows = 1 'gets rid of all rows exept the fixed ones
    txtEditGrid.Text = vbNullString 'empty the text box
End Sub
'%%%%%%%%%%%%%%%%%End of Code for FlxURls%%%%%%%%%%%%%%%%%%%%%%%%%%%%>

'<%%%%%%%%%%%%%%%%%Code for lstKeywords%%%%%%%%%%%%%%%%%%%%%%%
Private Sub lstKeywords_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "FlxURLs" Then
        Call cmdGetKWs_Click
    ElseIf Me.ActiveControl.Name = "FlxKey" Then
        cmdDeleteCell_Click
    End If
End Sub

Private Sub lstKeywords_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ih As Integer 'height of an item in the list box
    ih = TextHeight("A")
    DragLabel.Move lstKeywords.Left + fraKeywords.Left, lstKeywords.Top + _
                   fraKeywords.Top + Y - ih / 2, lstKeywords.Width, ih
    DragLabel.Drag 'drag label instead of the whole list box
End Sub

Private Sub cmdRemoveItem_Click()
    If lstKeywords.Text = vbNullString Then Exit Sub
    lstKeywords.RemoveItem lstKeywords.ListIndex
End Sub

Private Sub cmdRemoveItem_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "FlxKey" Then 'deletes item from flxurls
        cmdDeleteCell_Click                  'when it's dragged off
    ElseIf Me.ActiveControl.Name = "FlxURLs" Then
        cmdDeleteCellUrls_Click
    End If
End Sub

Private Sub cmdClearList_Click()
    If lstKeywords.ListCount = 0 Then Exit Sub
    Dim K As Integer
    For K = 0 To lstKeywords.ListCount - 1
        lstKeywords.RemoveItem 0
    Next
End Sub

Private Sub cmdClearList_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "FlxKey" Then 'deletes item from flxurls
        cmdDeleteCell_Click                  'when it's dragged off
    ElseIf Me.ActiveControl.Name = "FlxURLs" Then
        cmdDeleteCellUrls_Click
    End If
End Sub

Private Sub cmdAddToMK_Click()
    bLoading = True
    If lstKeywords.Text = vbNullString Then Exit Sub
    AddCell FlxKey, lstKeywords.Text
    bLoading = False
    FlxKey.Row = FlxKey.Rows - 1
End Sub

Private Sub cmdAddToMK_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "FlxKey" Then
        cmdDeleteCell_Click
    ElseIf Me.ActiveControl.Name = "FlxURLs" Then
        cmdDeleteCellUrls_Click
    End If
End Sub

Private Sub cmdAddAllToMK_Click()
    bLoading = True
    If lstKeywords.ListCount = 0 Then Exit Sub
    Dim K As Integer
    For K = 0 To lstKeywords.ListCount - 1
        lstKeywords.ListIndex = K
        AddCell FlxKey, lstKeywords.Text
    Next
    bLoading = False
    FlxKey.Row = 1
End Sub
'%%%%%%%%%%%%%%%%%End of Code for lstKeywords%%%%%%%%%%%%%%%%%%%%%%%>

'<%%%%%%%%%%%%%%%%Code for FlxKey%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private Sub FlxKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ih As Integer 'item height
    ih = FlxKey.CellHeight
    DragLabel.Move FlxKey.Left + fraKeywords.Left, FlxKey.Top + _
                   fraKeywords.Top + FlxKey.CellTop, FlxKey.CellWidth, ih
    DragLabel.Drag
End Sub
Private Sub FlxKey_EnterCell()
    If FlxKey.MouseRow = 0 Then Exit Sub
    With FlxKey
        If bLoading = False Then
            'put cell contents in text box
            txtEditGrid2.Text = .Text
            'move focus to text box fi it's enabled
            txtEditGrid2.Visible = True
            If txtEditGrid2.Enabled Then txtEditGrid2.SetFocus
        End If
    End With
End Sub

Private Sub FlxKey_LeaveCell()
    If bLoading = False And txtEditGrid2.Text <> "" Then
        FlxKey.Text = txtEditGrid2.Text
    End If
End Sub

Private Sub FlxKey_DragDrop(Source As Control, X As Single, Y As Single)
    If Me.ActiveControl.Name = "lstKeywords" And lstKeywords.Text <> "" Then
        AddCell FlxKey, lstKeywords.Text
    ElseIf Me.ActiveControl.Name = "FlxURLs" Then
        cmdDeleteCellUrls_Click
    End If
End Sub

Private Sub FlxKey_RowColChange()
    lblKeyNumber.Caption = CStr(FlxKey.Rows - 1)
End Sub

Private Sub chkEdit2_Click()
    txtEditGrid2.Enabled = Not txtEditGrid2.Enabled
End Sub

Private Sub cmdAddKeyword_Click()
    bLoading = True
    Dim KW As String
    KW = Trim(InputBox("Enter your Keyword:", Me.Caption))
    AddCell FlxKey, KW
    bLoading = False
End Sub

Private Sub cmdDeleteCell_Click()
    bLoading = True
    If FlxKey.Rows >= 3 Then
        FlxKey.RemoveItem FlxKey.Row
        ColorCells FlxKey
        FlxKey.Row = 1
        txtEditGrid2.Text = FlxKey.Text
        lblKeyNumber.Caption = CStr(FlxKey.Rows - 1)
    ElseIf FlxKey.Rows = 2 Then
        FlxKey.Rows = 1
        txtEditGrid2.Text = vbNullString
    End If
    bLoading = False
End Sub

Private Sub cmdReset2_Click()
    FlxKey.Rows = 1
    txtEditGrid2.Text = vbNullString
    lblKeyNumber.Caption = "0"
End Sub

Private Sub cmdGenHTML_Click()
    bLoading = True
    If FlxKey.Rows < 2 Then MsgBox "Please add some keywords first!", vbOKOnly + vbInformation, Me.Caption
    Dim kString As String: kString = "<META NAME=""Keywords"" CONTENT="""
    Dim K As Integer
    For K = 1 To FlxKey.Rows - 1
        FlxKey.Row = K
        If Len(kString) = 31 Then
            kString = kString & FlxKey.Text
        Else
            kString = kString & ", " & FlxKey.Text
        End If
    Next
    kString = kString & """>"
    InputBox "Here is the HTML code for your keywords. You can copy and paste this to your HTML file", Me.Caption, kString
    bLoading = False
End Sub
'%%%%%%%%%%%%%%%%%End of Code for FlxKey%%%%%%%%%%%%%%%%%%%%%%%%%%%>

'<%%%%%%%%%%%%%%%%%%% FUNCTIONS %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private Sub Inet1_StateChanged(ByVal State As Integer)
   ' Retrieve server response using the GetChunk
   ' method when State = 12. This assumes the
   ' data is text.

   Select Case State
    Case icError
        MsgBox "ERROR! Can't complete request.", vbCritical + vbOKOnly, "Keyword Grabber"
   Case Is = 1
        lblStatus.Caption = "Looking for host..."
    Case Is = 2
        lblStatus.Caption = "Host found!"
    Case Is = 3
        lblStatus.Caption = "Connecting..."
    Case Is = 4
        lblStatus.Caption = "Connected!"
    Case Is = 5
        lblStatus.Caption = "Sending request..."
   Case icResponseCompleted ' 12
        lblStatus.Caption = "Receiving response..."
    If GettingKeys = False Then
      Dim vtData As Variant ' Data variable.
      Dim bDone As Boolean: bDone = False 'tells me when I'm done
      Dim Count As Integer: Count = 0 'counts how many urls we've downloaded

      ' Get first chunk.
      vtData = Inet1.GetChunk(2000, icString) 'icstring tells it that it's going to
                                              'download text.
      Do While Not bDone 'if we haven't downloaded enough, keep going.
         aString = aString & vtData

         ' Get next chunk.
         vtData = Inet1.GetChunk(2000, icString)

         'The search engine puts "clsResultURL" before every urls of a result in the
         'html so we use it to count how many of the URL's we've retrieved. This will
         'stop downloading once we've got enough url's because i thought it would be
         'more efficient than downloading the entire file.
         If InStr(1, vtData, "Inktomi") Then
            Count = Count + 1 'got another URL
         End If

         If Len(vtData) = 0 Or Count = NumberToGet + 1 Then 'if we've got one more than enough, stop.
            bDone = True
         End If
      Loop
            Call GetURLs
    Else
        aString = Inet1.GetChunk(2000, icString)
        Call GetKeys
    End If
    
   End Select
End Sub

Private Sub GetURLs()
    On Error GoTo URLError
    bLoading = True
    'looks through search results, finds URL's and
    'puts them in the grid

    lblStatus.Caption = "Getting URL's..."
    Dim Count As Integer: Count = 0
    Dim PlaceHolder As Long: PlaceHolder = 5727 'skip the first few thou. char's
    Dim PlaceHolder2 As Long                     'because URL's are later.
    Dim URL As String
    
    Do
        PlaceHolder = InStr(PlaceHolder, aString, "Inktomi")
        PlaceHolder = InStr(PlaceHolder, aString, "http://")
        PlaceHolder2 = InStr(PlaceHolder, aString, "</font>")
        PlaceHolder2 = PlaceHolder2 - PlaceHolder
        URL = Mid(aString, PlaceHolder, PlaceHolder2)
        Count = Count + 1
        AddCell FlxURLs, URL
    Loop While Count <> NumberToGet
    
    aString = vbNullString
    FlxURLs.Row = 1
    txtEditGrid.Text = FlxURLs.Text
    lblStatus.Caption = "Ready"
    bLoading = False
    Exit Sub
    
URLError:
    bLoading = False
    MsgBox "There was an error processing your request. Please try again.", vbOKOnly + vbCritical, Me.Caption
End Sub

Private Sub GetKeys()
    bLoading = True
    lblStatus.Caption = "Getting Keywords"
    Dim PlaceHolder As Integer: PlaceHolder = 12
    Dim PlaceHolder2 As Integer
    Dim K As Integer
    
    If InStr(PlaceHolder, aString, "<META NAME=""Keywords""", vbTextCompare) Then
        PlaceHolder = InStr(PlaceHolder, aString, "<META NAME=""Keywords""", vbTextCompare)
        PlaceHolder = PlaceHolder + 21
        PlaceHolder = InStr(PlaceHolder, aString, "CONTENT=""", vbTextCompare)
        PlaceHolder = PlaceHolder + 9
        PlaceHolder2 = InStr(PlaceHolder, aString, """>")
        aString = Mid(aString, PlaceHolder, PlaceHolder2 - PlaceHolder)
        
        Keywords() = Split(aString, ",")
        For K = 0 To UBound(Keywords)
            Keywords(K) = Trim(Keywords(K))
            lstKeywords.AddItem Keywords(K)
        Next
    End If
    lblStatus.Caption = "Ready"
    bLoading = False
End Sub

Private Sub AddCell(ByRef grid As MSFlexGrid, Text As String)
    grid.Rows = grid.Rows + 1
    grid.Row = grid.Rows - 1
    If grid.Row Mod 2 = 0 Then grid.CellBackColor = &HC0E0FF
    grid.Text = Text
End Sub

Private Sub ColorCells(grid As MSFlexGrid)
    Dim I As Integer
    For I = 1 To grid.Rows - 1
        grid.Row = I
        If grid.Row Mod 2 = 0 Then
            grid.CellBackColor = &HC0E0FF
        Else
            grid.CellBackColor = vbWhite
        End If
    Next
End Sub

'%%%%%%%%%%%%%%%%%%%%% End of FUNCTIONS %%%%%%%%%%%%%%%%%%%%%%%%%%>
Private Sub mnuHelpAbout_Click()
    MsgBox "This program was created by Misael Pateyro" & vbCrLf & _
    "Please send any comments or questions to:" & vbCrLf & "papitas66@hotmail.com"
End Sub

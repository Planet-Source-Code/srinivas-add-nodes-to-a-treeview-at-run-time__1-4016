VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   1785
   ClientTop       =   1740
   ClientWidth     =   5415
   LinkTopic       =   "Form10"
   ScaleHeight     =   4875
   ScaleWidth      =   5415
   Begin ComctlLib.TreeView OrgTree 
      DragIcon        =   "Form1.frx":0000
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7646
      _Version        =   327682
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   3
      Left            =   4800
      Picture         =   "Form1.frx":0152
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   2
      Left            =   4800
      Picture         =   "Form1.frx":02A4
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   1
      Left            =   4800
      Picture         =   "Form1.frx":03F6
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   6
      Left            =   5160
      Picture         =   "Form1.frx":0548
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   5
      Left            =   5160
      Picture         =   "Form1.frx":064A
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   4
      Left            =   5160
      Picture         =   "Form1.frx":098C
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   3
      Left            =   4800
      Picture         =   "Form1.frx":0CCE
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   2
      Left            =   4800
      Picture         =   "Form1.frx":0DD0
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   1
      Left            =   4800
      Picture         =   "Form1.frx":1112
      Top             =   1440
      Width           =   240
   End
   Begin ComctlLib.ImageList TreeImages 
      Left            =   4800
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuNodes 
      Caption         =   "&Nodes"
      Begin VB.Menu mnuAddFactory 
         Caption         =   "Add &Factory"
      End
      Begin VB.Menu mnuAddGroup 
         Caption         =   "Add &Group"
      End
      Begin VB.Menu mnuAddPerson 
         Caption         =   "Add &Person"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ObjectType
    otNone = 0
    otFactory = 1
    otGroup = 2
    otPerson = 3
    otFactory2 = 4
    otGroup2 = 5
    otPerson2 = 6
End Enum

Private SourceNode As Object
Private SourceType As ObjectType
Private TargetNode As Object
' ***********************************************
' Return the node's object type.
' ***********************************************
Private Function NodeType(test_node As Node) As ObjectType
    If test_node Is Nothing Then
        NodeType = otNone
    Else
        Select Case Left$(test_node.Key, 1)
            Case "f"
                NodeType = otFactory
            Case "g"
                NodeType = otGroup
            Case "p"
                NodeType = otPerson
        End Select
    End If
End Function
' ***********************************************
' Prepare the ImageList and TreeView controls.
' ***********************************************
Private Sub Form_Load()
Dim i As Integer
Dim factory As Node
Dim group As Node
Dim person As Node

    ' Load pictures into the ImageList.
    For i = 1 To 6
        TreeImages.ListImages.Add , , TreeImage(i).Picture
    Next i
    
    ' Attach the TreeView to the ImageList.
    OrgTree.ImageList = TreeImages
    
    ' Create some nodes.
    Set factory = OrgTree.Nodes.Add(, , "f R & D", "R & D", otFactory, otFactory2)
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Engineering", "Engineering", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Cameron, Charlie", "Cameron, Charlie", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Davos, Debbie", "Davos, Debbie", otPerson, otPerson2)
    person.EnsureVisible
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Test", "Test", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Able, Andy", "Andy, Able", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Baker, Betty", "Baker, Betty", otPerson, otPerson2)
    person.EnsureVisible
    
    Set factory = OrgTree.Nodes.Add(, , "f Sales & Support", "Sales & Support", otFactory, otFactory2)
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Showroom Sales", "Showroom Sales", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Gaines, Gina", "Gaines, Gina", otPerson, otPerson2)
    person.EnsureVisible
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Field Service", "Field Service", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Helms, Harry", "Helms, Harry", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Ives, Irma", "Ives, Irma", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Jackson, Josh", "Jackson, Josh", otPerson, otPerson2)
    person.EnsureVisible
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Customer Support", "Customer Support", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Klug, Karl", "Klug, Karl", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Landau, Linda", "Landau, Linda", otPerson, otPerson2)
    person.EnsureVisible
End Sub
' ***********************************************
' Make the TreeView as large as possible.
' ***********************************************
Private Sub Form_Resize()
    OrgTree.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


' Add a new factory.
Private Sub mnuAddFactory_Click()
Dim name As String
Dim factory As Node

    name = InputBox("Factory Name", "New Factory", "")
    If name = "" Then Exit Sub
    
    Set factory = OrgTree.Nodes.Add(, , "f " & name, name, otFactory, otFactory2)
    factory.EnsureVisible
End Sub

' Add a new group.
Private Sub mnuAddGroup_Click()
Dim name As String
Dim factory As Node
Dim group As Node

    name = InputBox("Group Name", "New Group", "")
    If name = "" Then Exit Sub
    
    ' Find the factory that should hold the new group.
    Set factory = OrgTree.SelectedItem
    If NodeType(factory) = otPerson Then _
        Set factory = factory.Parent
    If NodeType(factory) = otGroup Then _
        Set factory = factory.Parent

    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g " & name, name, otGroup, otGroup2)
    group.EnsureVisible
End Sub

Private Sub mnuNodes_Click()
Dim selected_node As Node
Dim selected_type As ObjectType

    Set selected_node = OrgTree.SelectedItem
    If selected_node Is Nothing Then
        selected_type = otNone
    Else
        selected_type = NodeType(selected_node)
    End If

    ' You can always add a factory.
    
    ' You can add a group if a factory, person, or
    ' group is selected.
    mnuAddGroup.Enabled = (selected_type <> otNone)
    
    ' You can add a person if a group or person
    ' is selected.
    mnuAddPerson.Enabled = (selected_type = otPerson) _
        Or (selected_type = otGroup)
End Sub
' Add a new person.
Private Sub mnuAddPerson_Click()
Dim name As String
Dim group As Node
Dim person As Node

    name = InputBox("Person Name", "New Person", "")
    If name = "" Then Exit Sub
    
    ' Find the group that should hold the new person.
    Set group = OrgTree.SelectedItem
    If NodeType(group) = otPerson Then _
        Set group = group.Parent

    Set person = OrgTree.Nodes.Add(group, tvwChild, "p " & name, name, otPerson, otPerson2)
    person.EnsureVisible
End Sub

' ***********************************************
' Save the node pressed so we can drag it later.
' ***********************************************
Private Sub OrgTree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set SourceNode = OrgTree.HitTest(x, y)
End Sub

' ***********************************************
' Start a drag if one is not in progress.
' ***********************************************
Private Sub OrgTree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ' Start a new drag. Note that we do not get
        ' other MouseMove events while the drag is
        ' in progress.
        
        ' See what node we are dragging.
        SourceType = NodeType(SourceNode)

        ' Select this node. When no node is highlighted,
        ' this node will be displayed as selected. That
        ' shows where it will land if dropped.
        Set OrgTree.SelectedItem = SourceNode

        ' Set the drag icon for this source.
        OrgTree.DragIcon = IconImage(SourceType)
        OrgTree.Drag vbBeginDrag
    End If
End Sub
' ***********************************************
' The user is dropping. See if the drop is valid.
' ***********************************************
Private Sub OrgTree_DragDrop(Source As Control, x As Single, y As Single)
    If Not (OrgTree.DropHighlight Is Nothing) Then
        ' It's a valid drop. Set source node's
        ' parent to be the target node.
        Set SourceNode.Parent = OrgTree.DropHighlight
        Set OrgTree.DropHighlight = Nothing
    End If

    Set SourceNode = Nothing
    SourceType = otNone
End Sub
' ***********************************************
' The mouse is being dragged over the control.
' Highlight the appropriate node.
' ***********************************************
Private Sub OrgTree_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Dim target As Node
Dim highlight As Boolean

    ' See what node we're above.
    Set target = OrgTree.HitTest(x, y)
    
    ' If it's the same as last time, do nothing.
    If target Is TargetNode Then Exit Sub
    Set TargetNode = target
    
    highlight = False
    If Not (TargetNode Is Nothing) Then
        ' See what kind of node were above.
        If NodeType(TargetNode) + 1 = SourceType Then _
            highlight = True
    End If
    
    If highlight Then
        Set OrgTree.DropHighlight = TargetNode
    Else
        Set OrgTree.DropHighlight = Nothing
    End If
End Sub


<div align="center">

## Auto Check/UnCheck TreeView


</div>

### Description

This basic procedure will handle all parent and child Checkboxes in a TreeView control. If you check a child node it will automatically check the parent(s) and vise versa.
 
### More Info
 
This procedure requires you to pass in the Treeview and the current node.

Basic TreeView(s). This code assumes you already have a TreeView populated with "CheckBoxes" turned ON.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[kc7dji](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kc7dji.md)
**Level**          |Intermediate
**User Rating**    |4.4 (35 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kc7dji-auto-check-uncheck-treeview__1-30895/archive/master.zip)





### Source Code

```

'Put this call into the TreeViews NODECHECK procedure.
Private Sub MyTreeView_NodeCheck(ByVal Node As MSComctlLib.Node)
	Call TreeCheckBoxes(MyTreeView, Node)
end Sub
'Add this procedure to a Module or to the form the TreeView is contained.
Public Sub TreeCheckBoxes(TR As TreeView, CurrentNode As Node)
'This code is copyright (c)2002 by Scott Durrett - All Rights Reserved
'No changes are allow without written approval from the Author.
Dim liNodeIndex As Integer
Dim lbDirty As Boolean
Dim liCounter As Integer
Dim lParentNode As Node
Dim lChildNode As Node
lbDirty = False
liNodeIndex = CurrentNode.Index
If CurrentNode.Checked = True Then 'node is checked
  'Children Check/UnCheck
  If Not TR.Nodes.Item(CurrentNode.Index).Child Is Nothing Then
    Set lParentNode = TR.Nodes.Item(liNodeIndex).Child.FirstSibling
      Do While Not lParentNode Is Nothing
        lParentNode.Checked = CurrentNode.Checked
        If Not lParentNode.Child Is Nothing Then
          Set lChildNode = lParentNode.Child
            Do While Not lChildNode Is Nothing
              lChildNode.Checked = CurrentNode.Checked
                If Not lChildNode.Next Is Nothing Then
                  Set lChildNode = lChildNode.Next
                Else
                  Set lChildNode = lChildNode.Child
                End If
            Loop
        End If
        Set lParentNode = lParentNode.Next
      Loop
  End If
  '============================================================
  'Check all parent nodes
  Do While Not TR.Nodes.Item(liNodeIndex).Parent Is Nothing
    TR.Nodes.Item(liNodeIndex).Parent.Checked = CurrentNode.Checked
    liNodeIndex = TR.Nodes.Item(liNodeIndex).Parent.Index
  Loop
  '===========================
ElseIf CurrentNode.Checked = False Then 'node is unchecked
  'Children Check/UnCheck
  If Not TR.Nodes.Item(CurrentNode.Index).Child Is Nothing Then
    Set lParentNode = TR.Nodes.Item(liNodeIndex).Child.FirstSibling
      Do While Not lParentNode Is Nothing
        lParentNode.Checked = CurrentNode.Checked
        If Not lParentNode.Child Is Nothing Then
          Set lChildNode = lParentNode.Child
            Do While Not lChildNode Is Nothing
              lChildNode.Checked = CurrentNode.Checked
                If Not lChildNode.Next Is Nothing Then
                  Set lChildNode = lChildNode.Next
                Else
                  Set lChildNode = lChildNode.Child
                End If
            Loop
        End If
        Set lParentNode = lParentNode.Next
      Loop
  End If
  '============================================================
Set lParentNode = Nothing
Set lChildNode = Nothing
  If Not CurrentNode.Parent Is Nothing Then
    Set lParentNode = CurrentNode.Parent.Child
      Do While Not lParentNode Is Nothing
        Set lChildNode = lParentNode.FirstSibling
          Do While Not lChildNode Is Nothing
            If lChildNode.Checked = True Then
              lbDirty = True
              Exit Do
            End If
            'If Not lChildNode.Next Is Nothing Then
              Set lChildNode = lChildNode.Next
            'End If
          Loop
          If lbDirty = False Then
            If Not lParentNode.Parent Is Nothing Then
              lParentNode.Parent.Checked = False
              lbDirty = False
            End If
          Else
            Exit Do
          End If
      If Not lParentNode.Parent Is Nothing Then
        Set lParentNode = lParentNode.Parent
      Else
        Set lParentNode = lParentNode.Parent
      End If
    Loop
  End If
End If
Set CurrentNode = Nothing
Set lParentNode = Nothing
Set lChildNode = Nothing
End Sub
'The End
```


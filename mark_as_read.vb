Private WithEvents Items As Outlook.Items

Private Sub Application_Startup()
  Dim olNs As Outlook.NameSpace
  Dim Folder As Outlook.MAPIFolder

  Set olNs = Application.GetNamespace("MAPI")
  Set Folder = olNs.GetDefaultFolder(olFolderInbox)
  '// change the folder if need here
  Set Folder = olFolder.Folders("Deleted")
  Set Items = Folder.Items
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
  Item.UnRead = False
  Item.Save
End Sub
Attribute VB_Name = "LIPPackageBuilder"

Public Sub OpenPackageBuilder()
    On Error GoTo ErrorHandler
    Dim oDialog As New Lime.Dialog
    Dim idpersons As String
    Dim oItem As Lime.ExplorerItem
    oDialog.Type = lkDialogHTML
    oDialog.Property("url") = Application.WebFolder & "lbs.html?ap=packagebuilder&type=tab"
    oDialog.Property("height") = 900
    oDialog.Property("width") = 1600
    oDialog.show

    Exit Sub
ErrorHandler:
    Call UI.ShowError("Globals.OpenPackageBuilder")
End Sub

Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    If Me.NewRecord Then Me.Codigo = NovoCodigo(Me.RecordSource, Me.Codigo.ControlSource)
End Sub

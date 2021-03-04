Attribute VB_Name = "AtualizarDaBase"
Sub atualizar_da_base()

    ActiveSheet.PivotTables("Booking Dinamica").PivotCache.Refresh

End Sub


Attribute VB_Name = "LimparFiltros"
Sub limpar_filtros()
    
    Range("B2").Select
    
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Qual_a_restri��o?").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Confirma��o___Amendment").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Categoria").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Etapa").ClearManualFilter
    
End Sub


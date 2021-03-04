Attribute VB_Name = "LimparFiltros"
Sub limpar_filtros()
    
    Range("B2").Select
    
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Qual_a_restrição?").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Confirmação___Amendment").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Categoria").ClearManualFilter
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Etapa").ClearManualFilter
    
End Sub


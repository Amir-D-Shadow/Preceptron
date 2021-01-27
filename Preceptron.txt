VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "¤u§@ªí3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

    Dim Config_sheet As Worksheet
    Set Config_sheet = Sheets("Config")
    
    'Clear Config
    If Config_sheet.ChartObjects.Count <> 0 Then
    
        Config_sheet.ChartObjects.Delete
    
    End If
    

    Dim oChartObj As ChartObject
    Dim oChart As Chart
    
    Dim Data_sheet As Worksheet
    Set Data_sheet = Sheets("Data")
    
    
    Set oChartObj = Config_sheet.ChartObjects.Add(Top:=100, Left:=175, Width:=600, Height:=300)
    Set oChart = oChartObj.Chart
    
    oChart.ChartType = xlXYScatter
    oChart.SetSourceData Source:=Data_sheet.UsedRange.Offset(1, 0).Resize(Data_sheet.UsedRange.Rows.Count - 1, Data_sheet.UsedRange.Columns.Count - 1)
    oChart.ChartStyle = 239
    oChart.HasTitle = True
    oChart.ChartTitle.Text = "Binary Classification"
    
    oChart.Axes(xlCategory, xlPrimary).HasTitle = True
    oChart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = Data_sheet.Cells(1, 1).Value
    
    oChart.Axes(xlValue, xlPrimary).HasTitle = True
    oChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = Data_sheet.Cells(1, 2).Value

End Sub

Public Function num_of_row() As Long

    Dim last_row As Long
    Dim Data_sheet As Worksheet
    Set Data_sheet = Sheets("Data")
    
    'last_row = Data_sheet.UsedRange.Rows.Count
    last_row = Data_sheet.Cells(Rows.Count, 1).End(xlUp).Row
    num_of_row = last_row
    
End Function

Public Function num_of_col() As Long

    Dim last_col As Long
    Dim Data_sheet As Worksheet
    Set Data_sheet = Sheets("Data")
    
    'last_col = Data_sheet.UsedRange.Columns.Count
    last_col = Data_sheet.Cells(1, Columns.Count).End(xlToLeft).Column
    num_of_col = last_col
    
End Function

Private Sub CommandButton2_Click()

    Dim Data_sheet As Worksheet
    Set Data_sheet = Sheets("Data")
    
    Dim Para_sheet As Worksheet
    Set Para_sheet = Sheets("Parameter")
    
    Dim x_tensor() As Variant
    'ReDim x_tensor(Data_sheet.UsedRange.Rows.Count - 1, Data_sheet.UsedRange.Columns.Count - 1)
    ReDim x_tensor(num_of_row - 1, num_of_col - 1)
    Dim w_tensor() As Variant
    'ReDim w_tensor(Para_sheet.UsedRange.Rows.Count, Para_sheet.UsedRange.Columns.Count)
    ReDim w_tensor(Para_sheet.Cells(Rows.Count, 1).End(xlUp).Row, Para_sheet.Cells(1, Columns.Count).End(xlToLeft).Column)
    Dim y_tensor() As Variant
    'ReDim y_tensor(Data_sheet.UsedRange.Rows.Count - 1, 1)
    ReDim y_tensor(num_of_row - 1, 1)
    
    'Prepare data
    
    Dim i As Integer
    Dim j As Integer
    
    'x_tensor
    
    For i = 2 To num_of_row 'Data_sheet.UsedRange.Rows.Count
    
        For j = 1 To num_of_col - 1 'Data_sheet.UsedRange.Columns.Count - 1
        
            x_tensor(i - 2, j - 1) = Data_sheet.Cells(i, j).Value
            'MsgBox VarType(Data_sheet.Cells(i, j).Value)
            'MsgBox x_tensor(i - 2, j - 1)
        
        Next j
        
    Next i
    
    'w_tensor
    
        
    For i = 1 To Para_sheet.Cells(Rows.Count, 1).End(xlUp).Row 'Para_sheet.UsedRange.Rows.Count
    
        For j = 1 To Para_sheet.Cells(1, Columns.Count).End(xlToLeft).Column 'Para_sheet.UsedRange.Columns.Count
        
            w_tensor(i - 1, j - 1) = Para_sheet.Cells(i, j).Value
            'MsgBox VarType(Para_sheet.Cells(i, j).Value)
            'MsgBox w_tensor(i - 1, j - 1)
        
        Next j
        
    Next i
    
    'y_tensor
    
    j = num_of_col 'Data_sheet.UsedRange.Columns.Count
    
    For i = 2 To num_of_row 'Data_sheet.UsedRange.Rows.Count
    
        y_tensor(i - 2, 0) = Data_sheet.Cells(i, j).Value
        'MsgBox y_tensor(i - 2, 0)
        
    Next i
    
    
    'Debug Matrix
    
    'Dim debug_matrix() As Variant
    'debug_matrix = Matrix_Mult(x_tensor, w_tensor)
    
    'MsgBox VarType(debug_matrix)
    
    'Dim Config_sheet As Worksheet
    'Set Config_sheet = Sheets("Config")
    
    'For i = 0 To UBound(debug_matrix, 1) - 1
    
        'For j = 0 To UBound(debug_matrix, 2) - 1
        
            'Config_sheet.Cells(i + 1, j + 1).Value = debug_matrix(i, j)
            
        'Next j
        
    'Next i
    
    
    'Dim Cost_Function As Variant
    'Dim tmp_mat() As Variant
    'tmp_mat = Matrix_Mult(x_tensor, w_tensor)
    'MsgBox Matrix_Size(tmp_mat)
    'tmp_mat = Matrix_Transpose(tmp_mat)
    'Cost_Function = Matrix_Mult(tmp_mat, y_tensor)
    'MsgBox Cost_Function(0, 0)
    
    'Train model
    
    Dim cost As Double
    Dim learning_rate As Double
    
    learning_rate = 1
    
    Dim x_vector() As Variant
    ReDim x_vector(1, UBound(x_tensor, 2))
    
    Dim k As Integer
    Dim l As Integer
    
    Dim complete_flag As Integer
    complete_flag = 1
    
    
    Do While complete_flag = 1
        
        
        For i = 2 To num_of_row 'Data_sheet.UsedRange.Rows.Count
    
            complete_flag = 0
        
            'x_vector
            For j = 1 To num_of_col - 1 'Data_sheet.UsedRange.Columns.Count - 1
            
                x_vector(0, j - 1) = Data_sheet.Cells(i, j).Value
                'MsgBox x_vector(0, j - 1)
            
            Next j
            
            'Print_Vector (x_vector)
            
            cost = Matrix_Mult(x_vector, w_tensor)(0, 0)
            
            cost = cost * y_tensor(i - 2, 0)
            
            'MsgBox cost
            
            Do While cost <= 0
            
                'Update w tensor
                For k = 2 To Para_sheet.Cells(Rows.Count, 1).End(xlUp).Row 'Para_sheet.UsedRange.Rows.Count
                
                    For l = 1 To Para_sheet.Cells(1, Columns.Count).End(xlToLeft).Column 'Para_sheet.UsedRange.Columns.Count
                    
                        learing_rate = 1
                        w_tensor(k - 1, l - 1) = w_tensor(k - 1, l - 1) + learing_rate * y_tensor(i - 2, 0) * x_vector(0, k - 1)
                        Para_sheet.Cells(k, l).Value = w_tensor(k - 1, l - 1)
    
                    Next l
                    
                Next k
                
               'Update constant
               learing_rate = 1
               w_tensor(0, 0) = w_tensor(0, 0) + y_tensor(i - 2, 0) * learing_rate
               Para_sheet.Cells(1, 1).Value = w_tensor(0, 0)
               'MsgBox w_tensor(0, 0)
               
               'MsgBox learing_rate
  
               
               'Find Cost function
                cost = Matrix_Mult(x_vector, w_tensor)(0, 0)
            
                cost = cost * y_tensor(i - 2, 0)
                'MsgBox cost
                
                'MsgBox cost & " , " & i
                
                'check complete
                complete_flag = 1
               
            Loop
        

        Next i

    
    Loop
    
      
    'y_tensor = Application.WorksheetFunction.Transpose(y_tensor)
        
    'Release Memory
     
     Erase x_tensor
     Erase w_tensor
     Erase y_tensor
     Erase x_vector

End Sub

Public Function Sigmoid_Function(data_val As Double) As Double

    If data_val >= 0 Then
    
        Sigmoid_Function = 1
        
    Else
        
        Sigmoid_Function = -1
        
    End If

End Function

Public Function Matrix_Mult(data_set_1 As Variant, data_set_2 As Variant) As Variant

    Dim x_dimension_1 As Integer
    Dim y_dimension_1 As Integer
    
    x_dimension_1 = UBound(data_set_1, 1) - LBound(data_set_1, 1)
    y_dimension_1 = UBound(data_set_1, 2) - LBound(data_set_1, 2)
    
    Dim x_dimension_2 As Integer
    Dim y_dimension_2 As Integer
    
    x_dimension_2 = UBound(data_set_2, 1) - LBound(data_set_2, 1)
    y_dimension_2 = UBound(data_set_2, 2) - LBound(data_set_2, 2)
    
    If y_dimension_1 <> x_dimension_2 Then
    
        MsgBox "ERROR: Dimension Mismatch (" & CStr(x_dimension_1) & "X" & CStr(y_dimension_1) & " , " & CStr(x_dimension_2) & "X" & CStr(y_dimension_2) & ")"
        
    End If
    
    'MsgBox x_dimension_1 & "X" & CStr(y_dimension_1)
    'MsgBox x_dimension_2 & "X" & CStr(y_dimension_2)
    
    Dim new_matrix() As Variant
    ReDim new_matrix(x_dimension_1, y_dimension_2)
    
    Dim i, j, k As Integer
    
    
    For i = 0 To (x_dimension_1)
    
    
        For j = 0 To (y_dimension_2)
        
            For k = 0 To (y_dimension_1)
            
                new_matrix(i, j) = new_matrix(i, j) + data_set_1(i, k) * data_set_2(k, j)
            
            Next k
        
        Next j
    
    Next i
    
    Matrix_Mult = new_matrix
    
    Erase new_matrix
    
End Function

Public Function Matrix_Transpose(data_set As Variant) As Variant

    Dim i, j As Integer
    
    Dim new_matrix As Variant
    ReDim new_matrix(UBound(data_set, 2), UBound(data_set, 1))
    
    For i = 0 To UBound(data_set, 1)
    
        For j = 0 To UBound(data_set, 2)
        
            new_matrix(j, i) = data_set(i, j)
            
        Next j
        
    Next i
    
    Matrix_Transpose = new_matrix
    
    Erase new_matrix
    
End Function

Public Function Matrix_Size(data_set As Variant) As String

    Dim size As String
    
    size = UBound(data_set, 1) & "X" & UBound(data_set, 2)
    
    Matrix_Size = size

    
End Function

Public Function Matrix_Copy(data_set_target As Variant) As Variant

    
    Dim x_dimen_1 As Integer
    Dim y_dimen_1 As Integer
    
    x_dimen_1 = UBound(data_set_target, 1)
    y_dimen_1 = UBound(data_set_target, 2)
    
    Dim new_matrix As Variant
    ReDim new_matrix(x_dimen_1, y_dimen_2)
    
    For i = 0 To x_dimen_1 - 1
    
        For j = 0 To y_dimen_1 - 1
        
            new_matrix(i, j) = data_set_target(i, j)
            
        Next j
        
    Next i
    
    Matrix_Copy = new_matrix
    
    Erase new_matrix
        
End Function

Public Function Print_Vector(data_set As Variant)

    Dim vec_str As String
    
    Dim i As Integer
    
    For i = 0 To UBound(data_set, 1) - 1
    
        vec_str = vec_str & data_set(i, 0) & " , "
        
    Next i
    
    MsgBox vec_str
    
End Function

Private Sub CommandButton3_Click()
    
    Dim Data_sheet As Worksheet
    Set Data_sheet = Sheets("Data")

    MsgBox Data_sheet.UsedRange.Rows.Count
    
    MsgBox Data_sheet.Cells(Rows.Count, 1).End(xlUp).Row
    

    

End Sub

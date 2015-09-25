Attribute VB_Name = "Mod_Matrix"
''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �������z��̎��������擾����B
' �@�@�@�@�F array2[9][9] -> 2
' �@�@�@�@�F array3[10][20][30] -> 3
' �����P�@�F �������z��B
' �Ԃ�l�@�F �������z��̎�������Ԃ��B
' �g�p���@�F getDimension(xyArray)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getDimension(matrix As Variant) As Integer
    ' �����������߂�B
    getDimension = Application.DimSize(matrix)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �������z��̎w�莟���̗v�f�����擾����B
' �@�@�@�@�F array2[9][9], 1 -> 9
' �@�@�@�@�F array3[10][20][30], 1 -> 10
' �����P�@�F �������z��B
' �����Q�@�F �����B
' �Ԃ�l�@�F �w�莟���̗v�f����Ԃ��B
' �g�p���@�F getSize(xyArray, 1)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getSize(matrix As Variant, dimension As Integer) As Integer
    ' �w�莟���̗v�f�������߂�B
    getSize = Application.RBound(matrix, dimension)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �Q�����z��̎�����ꊷ����B
' �����P�@�F �Q�����z��B
' �Ԃ�l�@�F ��������ς݂̔z���Ԃ��B
' �g�p���@�F xyArray = swapXYAxis(xyArray)
' �@�@�@�@�F     A1  B1  C1      A1  A2  A3
' �@�@�@�@�F     A2  B2  C2  ->  B1  B2  B3
' �@�@�@�@�F     A3  B3  C3      C1  C2  C3
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function swapXYAxis(matrix As Variant) As Variant
    ' XY�������ւ���B
    swapXYAxis = Application.Transpose(matrix)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �Q�����z��̍s�͈͂����߂�B
' �����P�@�F �Q�����z��B
' �����Q�@�F �J�n�C���f�b�N�X�B
' �����R�@�F �z�񒷁B
' �Ԃ�l�@�F �s�͈͂����߂��z���Ԃ��B
' �g�p���@�F xyArray = narrowXYRow(xyArray, 2, 1)
' �@�@�@�@�F     A1  B1  C1      A2  B2  C2
' �@�@�@�@�F     A2  B2  C2  ->
' �@�@�@�@�F     A3  B3  C3
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function narrowXYRow(matrix As Variant, start_idx As Long, Optional length As Long = -1) As Variant
    ' ���[�v�̍ŏIIndex���擾����B
    Dim loopEndIndex As Long
    If length < 0 Then
        ' �����w��Ȃ��B
        loopEndIndex = UBound(matrix)
    Else
        ' �����w�肠��B
        loopEndIndex = start_idx + length - 1
        If UBound(matrix) < loopEndIndex Then
            loopEndIndex = UBound(matrix)
        End If
    End If
    
    
    Dim newMatrix As Variant
    ReDim newMatrix(1 To loopEndIndex - start_idx + 1) As Variant
    Dim newIndex As Integer
    newIndex = 1
    For idx = start_idx To loopEndIndex
        newMatrix(newIndex) = Application.Index(matrix, idx)
        newIndex = newIndex + 1
    Next
    
    narrowXYRow = newMatrix
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �Q�����z��̗�͈͂����߂�B
' �����P�@�F �Q�����z��B
' �����Q�@�F �J�n�C���f�b�N�X�B
' �����R�@�F �z�񒷁B
' �Ԃ�l�@�F ��͈͂����߂��z���Ԃ��B
' �g�p���@�F xyArray = narrowXYColumn(xyArray, 1, 2)
' �@�@�@�@�F     A1  B1  C1      A1  B2
' �@�@�@�@�F     A2  B2  C2  ->  A2  B2
' �@�@�@�@�F     A3  B3  C3      A3  B3
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function narrowXYColumn(ByVal matrix As Variant, start_idx As Long, Optional length As Long = -1) As Variant
    matrix = Application.Transpose(matrix)
    matrix = narrowXYRow(matrix, start_idx, length)
    matrix = Application.Transpose(matrix)
    
    narrowXYColumn = matrix
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''
' �����@�@�F �Q�����z����J�e�S���ŕ�������B
' �����P�@�F �Q�����z��B
' �����Q�@�F �J�e�S����C���f�b�N�X������B
' �Ԃ�l�@�F �J�e�S���i��̓���f�[�^�j�ŕ��������}�g���N�X�z���Ԃ��B
' �g�p���@�F xyArray = getMatrixDivideRow(xyArray, "1,2")
' �@�@�@�@�F     A1  B1  C1      A1  B1 C1  xyArray(1)
' �@�@�@�@�F     A1  B1  C2  ->  A1  B1 C2
' �@�@�@�@�F     A1  B2  C1      ---------
' �@�@�@�@�F                     A1  B2 C1  xyArray(2)
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getMatrixDivideRow(matrix As Variant, categoryIndexString As String) As Variant()
    Dim col As Collection
    Set col = getCategoryIndexList(matrix, categoryIndexString)

    Dim return_value() As Variant
    ReDim return_value(1 To col.Count) As Variant
    Dim newIndex, length, column_size As Long
    For idx1 = 1 To col.Count
        Dim start_idx, end_idx As Long
        start_idx = col.Item(idx1)
        If idx1 = col.Count Then
            end_idx = UBound(matrix, 1)
        Else
            end_idx = col.Item(idx1 + 1) - 1
        End If
        
        newIndex = 1
        length = end_idx - start_idx + 1
        column_size = UBound(matrix, 2)
        Dim newMatrix() As Variant
        ReDim newMatrix(1 To length, 1 To column_size) As Variant
        For idx2 = start_idx To end_idx
            For idx3 = 1 To column_size
                newMatrix(newIndex, idx3) = matrix(idx2, idx3)
            Next
            newIndex = newIndex + 1
        Next
        
        return_value(idx1) = newMatrix
    Next
    
    getMatrixDivideRow = return_value
End Function

Private Function getCategoryIndexList(matrix As Variant, categoryIndexString As String) As Collection
    Dim categoryIndexArray  As Variant
    categoryIndexArray = Split(categoryIndexString, ",")
    
    Dim col As Collection
    Set col = New Collection
    
    ' �X�^�[�g�C���f�b�N�X�̓J�e�S���J�n�f�[�^�B
    Call col.Add(LBound(matrix, 1))
    
    ' �J�e�S���̃C���f�b�N�X���w�肵�āA�T�C�Y�����߂�B
    For idx1 = LBound(matrix, 1) + 1 To UBound(matrix, 1)
        Dim currData As Variant
        Dim previousData As Variant
        currData = Application.Index(matrix, idx1)
        previousData = Application.Index(matrix, idx1 - 1)
        
        flg = True
        ' �w��J�e�S�����ƂɃ��[�v�B
        For idx2 = LBound(categoryIndexArray) To UBound(categoryIndexArray)
            categoryIndex = categoryIndexArray(idx2)
            ' �y�J�e�S���ɂ��āz
            ' ���݃f�[�^�ƂP�O�̃f�[�^�ňႢ������ΈقȂ�J�e�S���Ɣ��f����B
            If currData(categoryIndex) <> "" Then
                If currData(categoryIndex) <> previousData(categoryIndex) Then
                    flg = False
                    Exit For
                End If
            End If
        Next
        
        If flg = False Then
            Call col.Add(idx1)
        End If
    Next
    
    Set getCategoryIndexList = col
End Function


Private Function isSameCategory(data1 As Variant, data2 As Variant, categoryIndexString As String) As Boolean
    Dim flg As Boolean
    flg = True
    
    Dim categoryIndexArray  As Variant
    categoryIndexArray = Split(categoryIndexString, ",")
    
    ' �w��J�e�S�����ƂɃ��[�v�B
    For idx = LBound(categoryIndexArray) To UBound(categoryIndexArray)
        categoryIndex = categoryIndexArray(idx)
        ' �Q�̃f�[�^�ňႢ������ΈقȂ�J�e�S���Ɣ��f����B
        If data1(categoryIndex) <> "" And data2(categoryIndex) <> "" Then
            If data1(categoryIndex) <> data2(categoryIndex) Then
                flg = False
                Exit For
            End If
        End If
    Next
    
    isSameCategory = flg
End Function



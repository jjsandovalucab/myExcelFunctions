'declaro la funcion insertSort cuyo unico argumento es el vector de datos 
'(que debe ser de dimension 1xn o nx1)
Public Function insertsort(datos As Variant) As Variant

Dim elem As Variant
Dim v1 As Variant
Dim n As Integer
n = 0
'PASO 1: CUENTO CUANTOS DATOS HAY
For Each elem In datos:
    n = n + 1
Next elem
'PASO 2: TRASPASO AL VECTOR V1
'redimensiono el vector v1
ReDim v1(1 To n)
'lleno el vector v1
i = 1
For Each elem In datos:
    v1(i) = elem
    i = i + 1
Next elem

'PASO 3: ORDENO EL VECTOR V1 CON EL METODO DE INSERCION
For i = 2 To n
    'selecciono el numero a comparar 
    comparado = v1(i) 
    
    For j = i - 1 To 1 Step -1
        If v1(j) > comparado Then 'si el numero es mayor que el comparado, 
            v1(j + 1) = v1(j) 'se corre a la izquierda
            v1(j) = comparado 'el comparado ocupa su lugar
        Else 'si el numero es menor que el comparado
            v1(j + 1) = comparado 'el comparado se pone justo delante de el
            Exit For 'compara al siguiente numero
        End If
    Next j
Next i

'PASO 4: DEVUELVO EL VECTOR V1 (RECUERDA QUE V1 ES UNIDIMENSIONAL, DIMENSION=N)
insertsort = v1
End Function

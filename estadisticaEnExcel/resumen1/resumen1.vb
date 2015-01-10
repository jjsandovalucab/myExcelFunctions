Option Base 1
'declaro la funcion resumen1 cuyo unico argumento es el vector de datos 
'a resumir. Esta funcion devuelve una matriz con los resultados 
'y sus nombres (media, varianza,etc)
Public Function resumen1(datos As Variant) As Variant
'declaro datos2 como un vector de singles, cuya tamaño(el numero de datos) sera
'asignada luego
Dim datos2() As Single

'cada dato
Dim elem As Variant 
'suma, media, y numero de datos
Dim suma As Single, media As Single, n As Single 
'varianza, desvest, 2do momento y CV
Dim varianza As Single, desvest As Single, mom2 As Single, CV As Single
'minimo, cuartil1, mediana, cuartil3 y maximo
Dim minimo As Single, q1 As Single, mediana As Single, q3 As Single, maximo As Single

'PASO 1: CUANTOS DATOS HAY?
n = 0
For Each elem In datos:
    n = n + 1
Next elem
'ahora puedo dimensionar datos2
ReDim datos2(1 To n)

'PASO 2: CALCULO DE SUMA, MOM2, Y REGISTRO DE DATOS2
cont = 0
For Each elem In datos:
    cont = cont + 1
    suma = suma + elem
    mom2 = mom2 + elem ^ 2
    datos2(cont) = elem
Next elem

'PASO 3: CALCULO DE PROMEDIO, VARIANZA, DESVEST y CV
'calculo de promedio
media = suma / n
'calculo de varianza
varianza = mom2 - media ^ 2
'calculo de desvest
desvest = (varianza) ^ 0.5
'calculo de CV
CV = desvest / media

'PASO 4: DETERMINACION DE PUESTOS PERCENTILES
'se redondea hacia arriba porque el percentil x es el que acumula al menos el x% de los datos
p025 = Application.RoundUp(0.25 * n, 0)
p050 = Application.RoundUp(0.5 * n, 0)
p075 = Application.RoundUp(0.75 * n, 0)

'PASO 5: ORDENAR DATOS
Dim datos3 As Variant 'observa que no es un vector sino Variant, para que sea mas facil almacenar
'el resultado de insertSort
datos3 = insertsort(datos2) 

'PASO 6: EXTRACCION DE MIN, Q1, MEDIANA, Q3 Y MAX
'ahora que los datos estan ordenados, las posiciones indican cada percentil
minimo = datos3(1)
q1 = datos3(p025)
mediana = datos2(p050)
q3 = datos3(p075)
maximo = datos3(n)

'PASO 7: CONSTRUIR MATRIZ DE RESULTADOS
la matriz de resultados debe ser un String porque los nombres son strings, 
'y debe ser homogenea
Dim resultados() As String
Dim nombres As Variant
nombres = Array("suma", "media", "varianza", "desvest" _
, "CV", "minimo", "q1", "mediana", "q3", "maximo")
'ahora puedo dimensionar la matriz de resultados
ReDim resultados(1 To 2, 1 To UBound(nombres))

'puedo llenar los nombres
For j = 1 To UBound(nombres)
    resultados(1, j) = nombres(j)
Next j

'llenado de los resultados
resultados(2, 1) = suma
resultados(2, 2) = media
resultados(2, 3) = varianza
resultados(2, 4) = desvest
resultados(2, 5) = CV
resultados(2, 6) = minimo
resultados(2, 7) = q1
resultados(2, 8) = mediana
resultados(2, 9) = q3
resultados(2, 10) = maximo
'devuelvo la matriz resultados
resumen1 = resultados

End Function

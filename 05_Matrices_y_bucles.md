# Vectores Unidimensionales en VB.NET

Un **vector** (o array unidimensional) es una estructura de datos que almacena una colección de elementos del mismo tipo bajo un solo nombre, diferenciados por un índice.

### 1. Sintaxis y Declaración

En Visual Basic, los arreglos son base cero (el primer índice es 0).

| Acción | Sintaxis |
| --- | --- |
| **Declaración vacía** | `Dim nombre(tamaño) As Tipo` |
| **Declaración con valores** | `Dim nombre() As Tipo = {val1, val2, ...}` |
| **Asignación** | `nombre(0) = valor` |
| **Obtener longitud** | `nombre.Length` |

> **Nota:** Si declaras `Dim v(4) As Integer`, el vector tendrá **5 elementos** (índices del 0 al 4).

---

### 2. Recorrido con Ciclos

Existen tres formas principales de recorrer un vector:

* **For...Next:** Ideal cuando necesitas conocer el índice actual.
* **For Each...Next:** Más limpio y rápido si solo necesitas el valor del elemento (no el índice).
* **Do While / While:** Útil cuando la detención depende de una condición externa además del índice.

---

### 3. Ejemplo Práctico en Windows Forms

Para este ejemplo, imagina un formulario con un **Button** (`btnProcesar`) y un **ListBox** (`lstResultados`).

```vb
Public Class Form1

    Private Sub btnProcesar_Click(sender As Object, e As EventArgs) Handles btnProcesar.Click
        ' 1. Definición del vector
        Dim frutas() As String = {"Manzana", "Pera", "Uva", "Sandía", "Mango"}
        
        lstResultados.Items.Clear()

        ' --- CASO 1: Ciclo For...Next (Tradicional) ---
        lstResultados.Items.Add("--- Ciclo For ---")
        For i As Integer = 0 To frutas.Length - 1
            lstResultados.Items.Add("Índice " & i & ": " & frutas(i))
        Next

        lstResultados.Items.Add("") ' Espacio

        ' --- CASO 2: Ciclo For Each (Simplificado) ---
        lstResultados.Items.Add("--- Ciclo For Each ---")
        For Each fruta As String In frutas
            lstResultados.Items.Add("Elemento: " & fruta)
        Next

        lstResultados.Items.Add("") ' Espacio

        ' --- CASO 3: Ciclo While (Condicional) ---
        lstResultados.Items.Add("--- Ciclo While ---")
        Dim contador As Integer = 0
        While contador < frutas.Length
            lstResultados.Items.Add("Posición " & contador & ": " & frutas(contador))
            contador += 1
        End While

    End Sub
End Class

```

## Llenado y Recorrido de Vectores en VB.NET

A continuación, un ejemplo donde se declara un vector vacío de 5 posiciones y se llena dinámicamente utilizando diferentes estructuras de control.

### Ejemplo en Windows Forms

```vb
Public Class Form1

    Private Sub btnEjecutar_Click(sender As Object, e As EventArgs) Handles btnEjecutar.Click
        ' Declaración del vector con espacio para 5 elementos (índices 0 a 4)
        Dim datos(4) As Integer
        lstSalida.Items.Clear()

        ' --- 1. LLENADO CON CICLO FOR ---
        ' Se llena con números pares (2, 4, 6, 8, 10)
        For i As Integer = 0 To datos.Length - 1
            datos(i) = (i + 1) * 2
        Next
        lstSalida.Items.Add("Vector llenado con FOR.")

        ' --- 2. RECORRIDO CON FOR EACH (Para lectura) ---
        lstSalida.Items.Add("Contenido (For Each):")
        For Each numero As Integer In datos
            lstSalida.Items.Add(numero)
        Next

        lstSalida.Items.Add("-------------------")

        ' --- 3. LLENADO CON CICLO WHILE ---
        ' Sobrescribimos el vector con múltiplos de 10
        Dim j As Integer = 0
        While j < datos.Length
            datos(j) = (j + 1) * 10
            j += 1
        End While
        lstSalida.Items.Add("Vector sobrescrito con WHILE.")

        ' --- 4. RECORRIDO CON CICLO FOR (Para lectura) ---
        lstSalida.Items.Add("Contenido (For):")
        For k As Integer = 0 To datos.GetUpperBound(0)
            lstSalida.Items.Add("Posición " & k & ": " & datos(k))
        Next

    End Sub
End Class

```
# Matrices (Vectores Bidimensionales) en VB.NET

Una **matriz** es una estructura de datos organizada en filas y columnas, similar a una cuadrícula o tabla.

### 1. Sintaxis y Declaración

| Acción | Sintaxis |
| --- | --- |
| **Declaración vacía** | `Dim matriz(filas, columnas) As Tipo` |
| **Declaración con valores** | `Dim matriz(,) As Integer = {{1, 2}, {3, 4}}` |
| **Acceso a elemento** | `matriz(fila, columna) = valor` |
| **Límite de filas** | `matriz.GetUpperBound(0)` |
| **Límite de columnas** | `matriz.GetUpperBound(1)` |

---

### 2. Formas de Recorrido

* **Ciclos For Anidados:** Es el estándar. El ciclo externo recorre filas y el interno columnas.
* **For Each:** Recorre todos los elementos de forma secuencial (como si fuera una sola fila larga), pero no permite conocer la posición (fila/columna) actual.

---

### 3. Ejemplo 1: Matriz con Valores Iniciales (Lectura)

En este formulario usamos una matriz ya definida para mostrar su contenido en un `ListBox`.

```vb
Public Class FormLectura

    Private Sub btnMostrar_Click(sender As Object, e As EventArgs) Handles btnMostrar.Click
        ' Matriz de 3 filas y 2 columnas
        Dim notas(,) As Integer = {{85, 90}, {70, 80}, {95, 100}}
        lstSalida.Items.Clear()

        ' Recorrido con For Anidado
        For f As Integer = 0 To notas.GetUpperBound(0)
            Dim filaTexto As String = ""
            For c As Integer = 0 To notas.GetUpperBound(1)
                filaTexto &= " [" & notas(f, c) & "] "
            Next
            lstSalida.Items.Add("Fila " & f & ":" & filaTexto)
        Next
    End Sub

End Class

```

---

### 4. Ejemplo 2: Matriz Vacía (Llenado Dinámico y Lectura)

En este formulario declaramos una matriz de 3x3, la llenamos con una secuencia numérica usando ciclos y luego la leemos.

```vb
Public Class FormLlenado

    Private Sub btnEjecutar_Click(sender As Object, e As EventArgs) Handles btnEjecutar.Click
        Dim tabla(2, 2) As Integer ' Matriz 3x3
        Dim contador As Integer = 1
        lstSalida.Items.Clear()

        ' --- LLENADO CON CICLOS ---
        For f As Integer = 0 To 2
            For c As Integer = 0 To 2
                tabla(f, c) = contador
                contador += 1
            Next
        Next

        ' --- RECORRIDO CON FOR EACH (Solo lectura rápida) ---
        lstSalida.Items.Add("Recorrido secuencial:")
        For Each valor As Integer In tabla
            lstSalida.Items.Add("Valor: " & valor)
        Next

        lstSalida.Items.Add("-------------------")

        ' --- RECORRIDO CON WHILE (Manual) ---
        lstSalida.Items.Add("Recorrido con While:")
        Dim i As Integer = 0
        While i <= tabla.GetUpperBound(0)
            Dim j As Integer = 0
            Dim linea As String = ""
            While j <= tabla.GetUpperBound(1)
                linea &= tabla(i, j) & "  "
                j += 1
            End While
            lstSalida.Items.Add(linea)
            i += 1
        End While
    End Sub

End Class

```
# Operaciones con matrices
Esta guía detalla las operaciones fundamentales (CRUD: Crear, Leer, Actualizar, Borrar) y la interacción entre vectores y matrices en Visual Basic .NET.

---

### 1. Procesos entre Vectores y Matrices

A menudo es necesario transferir datos de una estructura a otra.

* **De Vector a Matriz:** Se usa para llenar una fila o columna específica de la matriz con los datos de un vector.
* **De Matriz a Vector:** Se extrae una fila o columna completa para procesarla de forma independiente.

### 2. Guardar (Escritura)

Para guardar, se asigna un valor a una posición específica mediante índices.

* **Vector:** `vector(indice) = valor`
* **Matriz:** `matriz(fila, columna) = valor`

### 3. Mostrar (Lectura)

Se recorre la estructura para enviar los datos a un control (ListBox, DataGridView o Label).

* **Vector:** Un ciclo `For` simple.
* **Matriz:** Dos ciclos `For` anidados.

### 4. Consultar (Búsqueda)

Consiste en comparar un valor buscado con cada elemento.

* **Lógica:** Se usa una variable booleana (bandera) para saber si se encontró y una variable para guardar la posición.

### 5. Modificar (Actualización)

Primero se debe buscar la posición del elemento. Una vez obtenida la fila/columna o índice, se sobrescribe el valor:

* `vector(posicionEncontrada) = nuevoValor`

### 6. Eliminar (Borrado)

En VB.NET, los arreglos tienen un tamaño fijo. "Eliminar" generalmente significa:

* **Limpiar:** Poner el valor en blanco (`""`) o cero (`0`).
* **Desplazar:** En vectores, mover los elementos siguientes una posición hacia atrás y reducir el tamaño con `ReDim Preserve`.

---

### Ejemplo Práctico en Windows Forms

Este ejemplo utiliza un vector para nombres de alumnos y una matriz para sus notas (2 notas por alumno).

```vb
Public Class FormGestion

    ' Declaración global de estructuras
    Dim alumnos(4) As String ' Vector para 5 alumnos
    Dim notas(4, 1) As Integer ' Matriz de 5 filas (alumnos) y 2 columnas (materias)
    Dim cantidadRegistrada As Integer = 0 ' Contador de registros actuales

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        ' --- 2. GUARDAR ---
        If cantidadRegistrada <= alumnos.GetUpperBound(0) Then
            ' Guardar en Vector
            alumnos(cantidadRegistrada) = txtNombre.Text
            
            ' Guardar en Matriz (Nota 1 y Nota 2)
            notas(cantidadRegistrada, 0) = CInt(txtNota1.Text)
            notas(cantidadRegistrada, 1) = CInt(txtNota2.Text)
            
            cantidadRegistrada += 1
            MessageBox.Show("Datos guardados correctamente.")
            LimpiarCampos()
        Else
            MessageBox.Show("Estructura llena.")
        End If
    End Sub

    Private Sub btnMostrar_Click(sender As Object, e As EventArgs) Handles btnMostrar.Click
        ' --- 3. MOSTRAR ---
        lstSalida.Items.Clear()
        ' Recorremos el vector y la matriz simultáneamente
        For i As Integer = 0 To cantidadRegistrada - 1
            Dim info As String = $"Alumno: {alumnos(i)} | " &
                                 $"Nota 1: {notas(i, 0)} | " &
                                 $"Nota 2: {notas(i, 1)}"
            lstSalida.Items.Add(info)
        Next
    End Sub

    Private Sub btnConsultar_Click(sender As Object, e As EventArgs) Handles btnConsultar.Click
        ' --- 4. CONSULTAR ---
        Dim buscado As String = txtNombre.Text
        Dim encontrado As Boolean = False

        For i As Integer = 0 To cantidadRegistrada - 1
            If alumnos(i).ToLower() = buscado.ToLower() Then
                ' Si lo encuentra, muestra los datos de la matriz en los TextBox
                txtNota1.Text = notas(i, 0).ToString()
                txtNota2.Text = notas(i, 1).ToString()
                MessageBox.Show($"Alumno encontrado en índice: {i}")
                encontrado = True
                Exit For
            End If
        Next

        If Not encontrado Then MessageBox.Show("No se encontró el alumno.")
    End Sub

    Private Sub btnModificar_Click(sender As Object, e As EventArgs) Handles btnModificar.Click
        ' --- 5. MODIFICAR ---
        ' Primero buscamos el índice por nombre
        Dim nombreMod As String = txtNombre.Text
        For i As Integer = 0 To cantidadRegistrada - 1
            If alumnos(i).ToLower() = nombreMod.ToLower() Then
                ' Actualizamos los valores en la matriz
                notas(i, 0) = CInt(txtNota1.Text)
                notas(i, 1) = CInt(txtNota2.Text)
                MessageBox.Show("Notas actualizadas.")
                Return
            End If
        Next
    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        ' --- 6. ELIMINAR (Lógica de desplazamiento) ---
        Dim nombreEliminar As String = txtNombre.Text
        Dim indiceEliminar As Integer = -1

        ' Buscar el índice
        For i As Integer = 0 To cantidadRegistrada - 1
            If alumnos(i).ToLower() = nombreEliminar.ToLower() Then
                indiceEliminar = i
                Exit For
            End If
        Next

        If indiceEliminar <> -1 Then
            ' Desplazar elementos hacia arriba para "tapar" el hueco
            For i As Integer = indiceEliminar To cantidadRegistrada - 2
                alumnos(i) = alumnos(i + 1)
                notas(i, 0) = notas(i + 1, 0)
                notas(i, 1) = notas(i + 1, 1)
            Next
            
            ' Limpiar la última posición
            alumnos(cantidadRegistrada - 1) = ""
            notas(cantidadRegistrada - 1, 0) = 0
            notas(cantidadRegistrada - 1, 1) = 0
            
            cantidadRegistrada -= 1
            MessageBox.Show("Registro eliminado.")
        End If
    End Sub

    Private Sub LimpiarCampos()
        txtNombre.Clear()
        txtNota1.Clear()
        txtNota2.Clear()
    End Sub

End Class

```

### Explicación de puntos clave:

1. **Sincronización:** El índice `i` sirve para el vector `alumnos` y para la fila de la matriz `notas`.
2. **GetUpperBound(0):** Se usa para no exceder el límite físico de los arreglos declarados.
3. **Desplazamiento en Eliminación:** Cuando eliminas a "Alumno B" en la posición 1, el "Alumno C" de la posición 2 debe pasar a la 1 para mantener la integridad de la lista.

# Ordenamiento burbuja
El **Ordenamiento Burbuja** (Bubble Sort) es uno de los algoritmos más sencillos para ordenar datos. Su nombre proviene de la forma en que los valores más pequeños (o grandes) "flotan" hacia la parte superior de la lista, tal como lo hacen las burbujas en el agua.

---

## 1. ¿Cómo funciona la lógica?

El algoritmo recorre el vector varias veces, comparando **elementos adyacentes** (uno al lado del otro). Si están en el orden incorrecto, los intercambia.

### Pasos del proceso:

1. Compara el primer elemento con el segundo.
2. Si el primero es mayor que el segundo (para orden ascendente), se intercambian.
3. Se pasa al siguiente par (segundo con tercero) y se repite.
4. Al terminar la primera vuelta, el número más grande habrá "flotado" hasta la última posición.
5. Se repite el proceso para el resto de los elementos, ignorando los que ya están ordenados al final.

---

## 2. La Complejidad

Técnicamente, este algoritmo tiene una complejidad de:


$$O(n^2)$$


Esto significa que si el número de elementos ($n$) aumenta, el tiempo de ejecución crece de forma cuadrática. Por ello, se recomienda solo para listas pequeñas.

---

## 3. El Intercambio (Swap)

Para intercambiar dos valores sin perder ninguno, necesitamos una **variable auxiliar** (o temporal):

* `temporal = A` (Guardamos el valor de A)
* `A = B` (Ponemos el valor de B en A)
* `B = temporal` (Ponemos el valor que guardamos originalmente en B)

---

## 4. Ejemplo en Windows Forms

En este ejemplo, tomamos un vector de números desordenados y los ordenamos al presionar un botón.

```vb
Public Class FormBurbuja

    Private Sub btnOrdenar_Click(sender As Object, e As EventArgs) Handles btnOrdenar.Click
        ' 1. Declaramos un vector con números desordenados
        Dim numeros() As Integer = {50, 10, 40, 20, 30}
        
        lstOriginal.Items.Clear()
        lstOrdenado.Items.Clear()

        ' Mostrar lista original
        For Each n In numeros
            lstOriginal.Items.Add(n)
        Next

        ' --- PROCESO DE ORDENAMIENTO BURBUJA ---
        
        ' El ciclo externo controla cuántas pasadas haremos.
        ' Se requiere un máximo de (n-1) pasadas.
        For i As Integer = 0 To numeros.Length - 2
            
            ' El ciclo interno realiza las comparaciones de parejas.
            ' El "- i" es una optimización: en cada pasada, el último elemento ya queda ordenado,
            ' por lo que no hace falta volver a compararlo.
            For j As Integer = 0 To numeros.Length - i - 2
                
                ' Comparar el elemento actual con el siguiente
                If numeros(j) > numeros(j + 1) Then
                    
                    ' Si el de la izquierda es mayor, aplicamos el SWAP (Intercambio)
                    Dim temporal As Integer = numeros(j) ' Guardamos el valor actual
                    numeros(j) = numeros(j + 1)          ' Movemos el menor a la izquierda
                    numeros(j + 1) = temporal            ' Movemos el mayor a la derecha
                    
                End If
            Next
        Next

        ' --- MOSTRAR RESULTADO ---
        For Each n In numeros
            lstOrdenado.Items.Add(n)
        Next

        MessageBox.Show("Vector ordenado con éxito.")
    End Sub

End Class

```

### Explicación de los límites en los ciclos:

* `numeros.Length - 2`: Usamos -2 porque en el ciclo interno comparamos `j + 1`. Si llegáramos hasta el final (`Length - 1`), al intentar acceder a `j + 1` el programa daría un error de "índice fuera de rango".
* `numeros.Length - i - 2`: Al restar `i`, el código se vuelve más eficiente porque deja de revisar los números que ya quedaron fijos al final en las pasadas anteriores.
# Taller 4 - Bucles y vectores
## Formulario 1: Ciclo `For...Next`

**Objetivo:** Generar una secuencia numérica conocida.

### 1. Teoría

* **Qué es:** Un ciclo que repite un bloque de código un número específico de veces.
* **Cuándo usarlo:** Cuando sabes exactamente cuántas veces quieres repetir algo (ej. "repetir 10 veces").
* **Sintaxis:**
```vb
For contador = Inicio To Fin [Step Paso]
    ' Código
Next

```


* **Regla:** No manipules la variable `contador` dentro del ciclo manualmente; deja que el `Next` lo haga.

### 2. Práctica (Diseño del Form1)

* **Controles:**
* 1 Botón (`btnEjecutar`) -> Texto: "Ver Tabla del 5"
* 1 ListBox (`lstResultados`)


* **Código:**

```vb
Private Sub btnEjecutar_Click(sender As Object, e As EventArgs) Handles btnEjecutar.Click
    lstResultados.Items.Clear()
    
    ' Repite del 1 al 10
    For i As Integer = 1 To 10
        Dim resultado As Integer = 5 * i
        lstResultados.Items.Add("5 x " & i & " = " & resultado)
    Next
End Sub

```

---

## Formulario 2: Ciclo `Do While...Loop`

**Objetivo:** Simular un proceso que continúa mientras una condición sea válida.

### 1. Teoría

* **Qué es:** Un ciclo que evalúa una condición **antes** de ejecutar. Si es `True`, entra.
* **Cuándo usarlo:** Cuando no sabes cuántas veces se repetirá, pero sabes que debe continuar **mientras** algo sea verdad.
* **Sintaxis:**
```vb
Do While Condicion
    ' Código
Loop

```


* **Regla:** Asegúrate de que algo dentro del ciclo cambie la condición a `False`, o tendrás un bucle infinito (se trabará el programa).

### 2. Práctica (Diseño del Form2)

**Escenario:** Llenar un tanque de agua.

* **Controles:**
* 1 Botón (`btnLlenar`) -> Texto: "Llenar Tanque"
* 1 ProgressBar (`prgTanque`) -> Propiedad `Maximum` = 100
* 1 Label (`lblEstado`)


* **Código:**

```vb
Private Sub btnLlenar_Click(sender As Object, e As EventArgs) Handles btnLlenar.Click
    Dim litros As Integer = 0
    prgTanque.Value = 0

    ' Mientras los litros sean menores a 100, sigue llenando
    Do While litros < 100
        litros += 10
        prgTanque.Value = litros
        lblEstado.Text = "Llenando: " & litros & "%"
        
        ' Pausa visual para ver el efecto
        Application.DoEvents()
        System.Threading.Thread.Sleep(200) 
    Loop
    
    MessageBox.Show("¡Tanque lleno!")
End Sub

```

---

## Formulario 3: Ciclo `Do Until...Loop`

**Objetivo:** Repetir una acción hasta lograr un objetivo.

### 1. Teoría

* **Qué es:** Lo opuesto al While. Repite el código **hasta que** la condición se vuelva `True`.
* **Cuándo usarlo:** Cuando quieres ejecutar algo repetidamente esperando que suceda un evento específico de parada.
* **Sintaxis:**
```vb
Do Until Condicion
    ' Código
Loop

```


* **Regla:** La condición de salida debe ser alcanzable.

### 2. Práctica (Diseño del Form3)

**Escenario:** Encontrar un número de la suerte aleatorio (el 7).

* **Controles:**
* 1 Botón (`btnBuscar`) -> Texto: "Buscar el 7"
* 1 ListBox (`lstIntentos`)


* **Código:**

```vb
Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
    lstIntentos.Items.Clear()
    Dim aleatorio As New Random()
    Dim numero As Integer = 0
    Dim contador As Integer = 0

    ' Repetir HASTA QUE el número sea 7
    Do Until numero = 7
        numero = aleatorio.Next(1, 11) ' Genera entre 1 y 10
        contador += 1
        lstIntentos.Items.Add("Intento " & contador & ": Salió el " & numero)
    Loop

    MessageBox.Show("¡Encontrado en " & contador & " intentos!")
End Sub

```

---

## Formulario 4: `For Each...Next` con Vectores

**Objetivo:** Recorrer un arreglo unidimensional de nombres.

### 1. Teoría

* **Qué es:** Un ciclo especializado para leer colecciones y vectores.
* **Cuándo usarlo:** Cuando quieres leer **todos** los datos de un vector uno por uno, sin preocuparte por los índices (0, 1, 2...).
* **Sintaxis:**
```vb
For Each elemento In Coleccion
    ' Código usando elemento
Next

```


* **Regla:** La variable `elemento` es de solo lectura (normalmente); no debes intentar cambiar el valor dentro del vector usando esta variable.

### 2. Práctica (Diseño del Form4)

* **Controles:**
* 1 Botón (`btnProcesar`) -> Texto: "Procesar Empleados"
* 1 ListBox (`lstEmpleados`)


* **Código:**

```vb
Private Sub btnProcesar_Click(sender As Object, e As EventArgs) Handles btnProcesar.Click
    ' 1. Declarar e inicializar el vector (Arreglo de 1 dimensión)
    Dim empleados() As String = {"Carlos", "Ana", "Luis", "Sofía", "Jorge"}

    lstEmpleados.Items.Clear()

    ' 2. Recorrer el vector con For Each
    ' La variable 'persona' tomará el valor de cada nombre automáticamente
    For Each persona As String In empleados
        lstEmpleados.Items.Add("Empleado: " & persona)
    Next
    
    MessageBox.Show("Se procesaron " & empleados.Length & " empleados.")
End Sub

```

## Vectores
Este ejemplo es un **"Gestor de Calificaciones Escolar"**. Es el escenario clásico y más efectivo para entender vectores porque usaremos **Arreglos Paralelos** (dos vectores que comparten el mismo índice para relacionar datos).

### Escenario

Vamos a almacenar hasta **5 alumnos**.

1. **Vector 1:** Nombres (String)
2. **Vector 2:** Calificaciones (Integer)

---

### 1. Diseño del Formulario (Controles)

Arrastra estos controles a tu Formulario y configúralos así:

| Control | Nombre (Name) | Propiedad Text / Settings |
| --- | --- | --- |
| **Label** | - | "Nombre del Alumno:" |
| **TextBox** | `txtNombre` | (Vacío) |
| **Label** | - | "Calificación (0-100):" |
| **NumericUpDown** | `numNota` | Minimum: 0, Maximum: 100 |
| **Button** | `btnAgregar` | "Guardar en Vector" |
| **ListBox** | `lstRegistro` | (Vacío) |
| **Button** | `btnCalcular` | "Calcular Promedios" |
| **Label** | `lblResultado` | "Promedio: 0 / Nota Alta: 0" |
| **ComboBox** | `cmbIndices` | (Vacío) - Lo llenaremos por código |
| **Button** | `btnVerIndice` | "Ver Dato por Índice" |

---

### 2. El Código Completo

Copia y pega este código en tu formulario. Lee los comentarios, explican el **por qué** de cada línea.

```vb
Public Class Form1

    ' 1. DECLARACIÓN DE VECTORES (Ámbito de Clase)
    ' Creamos 5 espacios (del 0 al 4).
    ' Son "paralelos" porque el índice 0 de nombres corresponde al índice 0 de notas.
    Dim vectorNombres(4) As String
    Dim vectorNotas(4) As Integer

    ' Variable para saber en qué casilla vamos (Índice actual)
    Dim indiceActual As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Llenamos el ComboBox con los índices disponibles (0, 1, 2, 3, 4)
        ' Esto es visual, para que entiendas las posiciones.
        For i As Integer = 0 To 4
            cmbIndices.Items.Add("Posición " & i)
        Next
        cmbIndices.SelectedIndex = 0
    End Sub

    ' 2. LLENADO DEL VECTOR (Escritura)
    Private Sub btnAgregar_Click(sender As Object, e As EventArgs) Handles btnAgregar.Click
        ' Regla de Oro: Verificar que no nos salgamos del límite del vector
        If indiceActual > 4 Then
            MessageBox.Show("¡El vector está lleno! Solo caben 5 alumnos.")
            Return
        End If

        ' Guardamos los datos en la posición actual
        vectorNombres(indiceActual) = txtNombre.Text
        vectorNotas(indiceActual) = CInt(numNota.Value)

        ' Mostramos visualmente lo que pasó
        lstRegistro.Items.Add("Índice [" & indiceActual & "] - " & txtNombre.Text & " (" & numNota.Value & ")")

        ' IMPORTANTE: Aumentamos el índice para la siguiente vez
        indiceActual += 1

        ' Limpiamos cajas
        txtNombre.Clear()
        numNota.Value = 0
        txtNombre.Focus()
    End Sub

    ' 3. RECORRIDO DEL VECTOR (Lectura masiva con For)
    Private Sub btnCalcular_Click(sender As Object, e As EventArgs) Handles btnCalcular.Click
        Dim suma As Integer = 0
        Dim notaMasAlta As Integer = 0
        Dim mejorAlumno As String = ""

        ' Usamos un ciclo For porque sabemos exactamente cuántos alumnos registramos
        ' Recorremos desde 0 hasta (indiceActual - 1) para no leer casillas vacías
        For i As Integer = 0 To indiceActual - 1
            
            ' Acumulamos la suma para el promedio
            suma += vectorNotas(i)

            ' Lógica para encontrar el mayor
            If vectorNotas(i) > notaMasAlta Then
                notaMasAlta = vectorNotas(i)
                mejorAlumno = vectorNombres(i)
            End If
        Next

        ' Cálculo final
        Dim promedio As Double = 0
        If indiceActual > 0 Then promedio = suma / indiceActual

        lblResultado.Text = "Promedio: " & promedio.ToString("0.0") & 
                            " | Mejor Nota: " & mejorAlumno & " (" & notaMasAlta & ")"
    End Sub

    ' 4. ACCESO DIRECTO (Lectura por Índice específico)
    Private Sub btnVerIndice_Click(sender As Object, e As EventArgs) Handles btnVerIndice.Click
        ' Recuperamos el índice que el usuario eligió en el ComboBox (0 al 4)
        Dim indiceSeleccionado As Integer = cmbIndices.SelectedIndex

        ' Verificamos si ya hay datos ahí (para no mostrar vacíos)
        If indiceSeleccionado >= indiceActual Then
            MessageBox.Show("Esa posición (" & indiceSeleccionado & ") aún está vacía.")
        Else
            ' Accedemos directamente a la casilla sin recorrer todo
            Dim nombreRecuperado As String = vectorNombres(indiceSeleccionado)
            Dim notaRecuperada As Integer = vectorNotas(indiceSeleccionado)

            MessageBox.Show("En la posición " & indiceSeleccionado & " vive: " & nombreRecuperado)
        End If
    End Sub

End Class

```

---

### Conceptos Clave en este Ejemplo

1. **Declaración Global:** Los vectores se declaran *fuera* de los botones (`Dim vectorNombres(4)`). Si los declaras dentro del botón, se borrarán cada vez que hagas clic.
2. **Índice (Pointer):** Usamos la variable `indiceActual` para controlar en qué "cajita" guardamos el siguiente dato.
3. **Límite (Bounds):** El `If indiceActual > 4` protege tu programa de errores. Si intentas escribir en la posición 5 de un vector de tamaño 4, el programa explota (*IndexOutOfRangeException*).
4. **Vectores Paralelos:** Observa cómo `vectorNombres(i)` y `vectorNotas(i)` están conectados solo por el número `i`. Si ordenaras uno sin ordenar el otro, los datos se mezclarían erróneamente.
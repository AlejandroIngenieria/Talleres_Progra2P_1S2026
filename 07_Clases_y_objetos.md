> **Proyecto:** Clase7 — Windows Forms con VB.NET (.NET 10)
>
> Esta guía explica paso a paso los conceptos de **Programación Orientada a Objetos (POO)** y el uso del control **DataGridView**, tal como se aplican en este proyecto.

---

## 📑 Tabla de Contenidos

1. [¿Qué es una Clase?](#1--qué-es-una-clase)
2. [¿Qué es un Objeto?](#2--qué-es-un-objeto)
3. [La Clase `Persona` de este proyecto](#3--la-clase-persona-de-este-proyecto)
4. [Propiedades](#4--propiedades)
5. [Constructor (`Sub New`)](#5--constructor-sub-new)
6. [Arreglos de Objetos](#6--arreglos-de-objetos)
7. [¿Qué es un DataGridView?](#7--qué-es-un-datagridview)
8. [Form1 — Uso básico del DataGridView](#8--form1--uso-básico-del-datagridview)
9. [Form2 — CRUD completo (Buscar, Actualizar, Eliminar)](#9--form2--crud-completo-buscar-actualizar-eliminar)
10. [Form3 — Controles dinámicos con Panel y FlowLayoutPanel](#10--form3--controles-dinámicos-con-panel-y-flowlayoutpanel)
11. [Resumen de conceptos clave](#11--resumen-de-conceptos-clave)
12. [Glosario](#12--glosario)

---

## 1. 🧱 ¿Qué es una Clase?

Una **clase** es como un **molde** o **plano** que define las características (propiedades) y comportamientos (métodos) que tendrán los objetos creados a partir de ella.

### Analogía del mundo real

Piensa en una clase como el **plano de una casa**:

- El plano dice que toda casa tiene: paredes, techo, puertas, ventanas.
- Pero el plano **no es** una casa, es solo la **descripción** de cómo será.

En programación:

```
Clase = Plano / Molde / Plantilla
```

### Estructura básica de una clase en VB.NET

```vb
Public Class NombreDeLaClase
    ' Aquí van las propiedades (características)
    ' Aquí van los métodos (acciones)
End Class
```

- `Public` → Significa que la clase es accesible desde cualquier parte del proyecto.
- `Class` → Palabra clave para declarar una clase.
- `End Class` → Marca el final de la definición de la clase.

---

## 2. 📦 ¿Qué es un Objeto?

Un **objeto** es una **instancia concreta** de una clase. Es decir, es una "copia" real creada a partir del molde (la clase).

### Siguiendo la analogía

- **Clase** = Plano de la casa
- **Objeto** = La casa ya construida

Puedes construir **muchas casas** (objetos) a partir del **mismo plano** (clase), y cada una puede tener valores diferentes (color, tamaño, etc.).

### Cómo crear un objeto en VB.NET

```vb
' La palabra clave "New" crea una nueva instancia (objeto) de la clase
Dim miPersona As New Persona("Ana", 25, "Guatemala")
```

| Parte del código | ¿Qué hace? |
|---|---|
| `Dim miPersona` | Declara una variable llamada `miPersona` |
| `As New Persona(...)` | Crea un nuevo objeto de tipo `Persona` |
| `"Ana", 25, "Guatemala"` | Son los valores que se envían al constructor |

---

## 3. 👤 La Clase `Persona` de este proyecto

Esta es la clase principal del proyecto. Se encuentra en el archivo `Persona.vb`:

```vb
Public Class Persona

    ' Propiedades públicas
    Public Property Nombre As String
    Public Property Edad As Integer
    Public Property Ciudad As String
    Public Property UrlImagen As String

    ' Constructor
    Public Sub New(nombre As String, edad As Integer, ciudad As String, Optional url As String = Nothing)
        Me.Nombre = nombre
        Me.Edad = edad
        Me.Ciudad = ciudad
        Me.UrlImagen = url
    End Sub

End Class
```

### Desglose línea por línea

| Línea | Explicación |
|---|---|
| `Public Class Persona` | Inicia la definición de la clase `Persona`. |
| `Public Property Nombre As String` | Crea una propiedad llamada `Nombre` que almacena texto. |
| `Public Property Edad As Integer` | Crea una propiedad llamada `Edad` que almacena un número entero. |
| `Public Property Ciudad As String` | Crea una propiedad llamada `Ciudad` que almacena texto. |
| `Public Property UrlImagen As String` | Crea una propiedad para guardar la URL de una imagen. |
| `Public Sub New(...)` | Es el **constructor**: el método que se ejecuta al crear un objeto. |
| `Me.Nombre = nombre` | Asigna el valor recibido por parámetro a la propiedad del objeto. |
| `Optional url As String = Nothing` | Parámetro **opcional**: si no se envía, su valor será `Nothing` (nulo). |
| `End Sub` | Fin del constructor. |
| `End Class` | Fin de la clase. |

---

## 4. 🏷️ Propiedades

Las **propiedades** son las **características** que describe una clase. Son como las "variables" que pertenecen a cada objeto.

```vb
Public Property Nombre As String
```

### ¿Qué significa `Public Property`?

- `Public` → Cualquier parte del código puede leer y escribir esta propiedad.
- `Property` → Palabra clave que indica que es una propiedad (no una variable común).

### ¿Cómo se usan las propiedades?

```vb
' Crear un objeto
Dim p As New Persona("Ana", 25, "Guatemala")

' LEER una propiedad
Dim elNombre As String = p.Nombre    ' Resultado: "Ana"

' ESCRIBIR (cambiar) una propiedad
p.Nombre = "María"                    ' Ahora el nombre es "María"
```

### Tipos de datos usados en este proyecto

| Tipo | Descripción | Ejemplo |
|---|---|---|
| `String` | Texto (cadena de caracteres) | `"Guatemala"` |
| `Integer` | Número entero | `25` |

---

## 5. 🔨 Constructor (`Sub New`)

El **constructor** es un método especial que se ejecuta **automáticamente** cuando creas un nuevo objeto con la palabra `New`.

```vb
Public Sub New(nombre As String, edad As Integer, ciudad As String, Optional url As String = Nothing)
    Me.Nombre = nombre
    Me.Edad = edad
    Me.Ciudad = ciudad
    Me.UrlImagen = url
End Sub
```

### ¿Para qué sirve?

Sirve para **inicializar** el objeto con valores desde el momento en que se crea, en lugar de asignarlos uno por uno después.

### Sin constructor (más código):

```vb
Dim p As New Persona()
p.Nombre = "Ana"
p.Edad = 25
p.Ciudad = "Guatemala"
```

### Con constructor (más limpio):

```vb
Dim p As New Persona("Ana", 25, "Guatemala")
```

### ¿Qué es `Me`?

`Me` hace referencia al **objeto actual** (el que se está creando). Se usa para diferenciar entre el **parámetro** y la **propiedad** cuando tienen el mismo nombre:

```vb
Me.Nombre = nombre
'  ↑              ↑
'  Propiedad      Parámetro
'  del objeto     que recibimos
```

### ¿Qué es `Optional`?

Un parámetro `Optional` no es obligatorio al crear el objeto:

```vb
' Sin URL (usa el valor por defecto Nothing)
Dim p1 As New Persona("Ana", 25, "Guatemala")

' Con URL
Dim p2 As New Persona("Ana", 25, "Guatemala", "https://ejemplo.com/foto.jpg")
```

---

## 6. 📚 Arreglos de Objetos

Un **arreglo** (array) es una colección de elementos del mismo tipo, almacenados en posiciones numeradas (índices).

### Declaración de un arreglo de objetos

```vb
' Crea un arreglo que puede almacenar 3 objetos Persona (posiciones 0, 1, 2)
Private personas(2) As Persona
```

> ⚠️ **Importante:** En VB.NET, el número entre paréntesis es el **índice máximo**, no la cantidad.
> `personas(2)` crea **3 posiciones**: índice 0, 1 y 2.

### Llenar el arreglo

```vb
personas(0) = New Persona("Ana", 25, "Guatemala")
personas(1) = New Persona("Luis", 30, "México")
personas(2) = New Persona("María", 22, "El Salvador")
```

### Visualización del arreglo en memoria

```
Índice:     0              1              2
         ┌──────────┐  ┌──────────┐  ┌──────────┐
         │  Ana     │  │  Luis    │  │  María   │
         │  25      │  │  30      │  │  22      │
         │ Guatemala│  │  México  │  │ El Salv. │
         └──────────┘  └──────────┘  └──────────┘
```

### Recorrer un arreglo con `For`

```vb
For i As Integer = 0 To personas.Length - 1
    ' personas(i) accede al objeto en la posición i
    Console.WriteLine(personas(i).Nombre)
Next
```

| Expresión | Significado |
|---|---|
| `personas.Length` | Cantidad total de elementos del arreglo (3 en este caso) |
| `personas.Length - 1` | Último índice válido (2 en este caso) |
| `personas(i)` | Accede al objeto en la posición `i` |
| `personas(i).Nombre` | Accede a la propiedad `Nombre` del objeto en la posición `i` |

---

## 7. 📊 ¿Qué es un DataGridView?

El **DataGridView** es un control de Windows Forms que muestra datos en forma de **tabla** (filas y columnas), similar a una hoja de Excel.

```
┌──────────────────────────────────────────────┐
│  Nombre       │  Edad        │  Ciudad       │  ← Encabezados (columnas)
├──────────────────────────────────────────────┤
│  Ana          │  25          │  Guatemala    │  ← Fila 0
│  Luis         │  30          │  México       │  ← Fila 1
│  María        │  22          │  El Salvador  │  ← Fila 2
└──────────────────────────────────────────────┘
```

### Conceptos clave del DataGridView

| Concepto | Descripción |
|---|---|
| **Columna** (`Column`) | Cada categoría de datos (Nombre, Edad, Ciudad) |
| **Fila** (`Row`) | Cada registro o entrada de datos |
| **Celda** (`Cell`) | La intersección de una fila y una columna (un dato individual) |
| **`CurrentCell`** | La celda actualmente seleccionada por el usuario |
| **`CurrentRow`** | La fila actualmente seleccionada por el usuario |

### Operaciones básicas del DataGridView

```vb
' Limpiar todo el contenido
dgvPersonas.Columns.Clear()    ' Elimina todas las columnas
dgvPersonas.Rows.Clear()       ' Elimina todas las filas

' Que las columnas se ajusten al ancho disponible
dgvPersonas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill

' Agregar columnas manualmente
' Parámetro 1: nombre interno    Parámetro 2: texto del encabezado
dgvPersonas.Columns.Add("Nombre", "Nombre")
dgvPersonas.Columns.Add("Edad", "Edad")
dgvPersonas.Columns.Add("Ciudad", "Ciudad")

' Agregar una fila con datos
dgvPersonas.Rows.Add("Ana", 25, "Guatemala")
```

---

## 8. 📋 Form1 — Uso básico del DataGridView

El `Form1` demuestra las operaciones fundamentales: **cargar datos**, **leer una celda** y **leer una fila completa**.

### Controles del formulario

| Control | Tipo | Función |
|---|---|---|
| `dgvPersonas` | `DataGridView` | Tabla donde se muestran los datos |
| `btnCargar` | `Button` | Carga los datos de personas en la tabla |
| `btnMostrarCelda` | `Button` | Muestra el valor de la celda seleccionada |
| `btnMostrarFila` | `Button` | Muestra todos los datos de la fila seleccionada |
| `Button1` | `Button` | Navega al Form2 |

### Código completo explicado

#### Declaración del arreglo

```vb
Public Class Form1
    ' Se declara un arreglo de 3 posiciones (0, 1, 2) de tipo Persona
    Private personas(2) As Persona
```

- `Private` → Solo se puede usar dentro de `Form1`.
- `personas(2)` → Arreglo con 3 espacios.

#### Botón "Cargar" — Llenar el DataGridView

```vb
Private Sub btnCargar_Click(sender As Object, e As EventArgs) Handles btnCargar.Click

    ' 1. Crear objetos Persona y guardarlos en el arreglo
    personas(0) = New Persona("Ana", 25, "Guatemala")
    personas(1) = New Persona("Luis", 30, "México")
    personas(2) = New Persona("María", 22, "El Salvador")

    ' 2. Limpiar el DataGridView (por si ya tenía datos de una carga anterior)
    dgvPersonas.Columns.Clear()
    dgvPersonas.Rows.Clear()
    dgvPersonas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill

    ' 3. Crear las columnas manualmente
    dgvPersonas.Columns.Add("Nombre", "Nombre")
    dgvPersonas.Columns.Add("Edad", "Edad")
    dgvPersonas.Columns.Add("Ciudad", "Ciudad")

    ' 4. Recorrer el arreglo y agregar cada persona como una fila
    For i As Integer = 0 To personas.Length - 1
        dgvPersonas.Rows.Add(
            personas(i).Nombre,
            personas(i).Edad,
            personas(i).Ciudad
        )
    Next

End Sub
```

**Flujo paso a paso:**

```
1. Se crean 3 objetos Persona → se guardan en el arreglo
2. Se limpia la tabla por si ya tenía datos
3. Se crean 3 columnas: Nombre, Edad, Ciudad
4. Se recorre el arreglo con For y se agrega cada persona como fila
```

#### Botón "Mostrar Celda" — Leer una celda seleccionada

```vb
Private Sub btnMostrarCelda_Click(sender As Object, e As EventArgs) Handles btnMostrarCelda.Click

    ' Primero verificamos que haya una celda seleccionada
    If dgvPersonas.CurrentCell IsNot Nothing Then

        Dim fila As Integer = dgvPersonas.CurrentCell.RowIndex       ' Número de fila
        Dim columna As Integer = dgvPersonas.CurrentCell.ColumnIndex  ' Número de columna
        Dim valor As Object = dgvPersonas.CurrentCell.Value           ' Valor de la celda

        MessageBox.Show(
            "Fila: " & fila &
            vbCrLf & "Columna: " & columna &
            vbCrLf & "Valor: " & valor.ToString()
        )

    Else
        MessageBox.Show("No hay ninguna celda seleccionada.")
    End If

End Sub
```

**Conceptos importantes:**

| Expresión | Significado |
|---|---|
| `CurrentCell` | La celda donde el usuario hizo clic |
| `IsNot Nothing` | Verifica que el valor **no sea nulo** (que sí exista) |
| `RowIndex` | Índice numérico de la fila (empieza en 0) |
| `ColumnIndex` | Índice numérico de la columna (empieza en 0) |
| `.Value` | El dato almacenado en esa celda |
| `vbCrLf` | Salto de línea (para que el mensaje se vea en varias líneas) |

#### Botón "Mostrar Fila" — Leer una fila completa

```vb
Private Sub btnMostrarFila_Click(sender As Object, e As EventArgs) Handles btnMostrarFila.Click

    If dgvPersonas.CurrentRow IsNot Nothing Then

        Dim filaSeleccionada = dgvPersonas.CurrentRow

        ' Acceder a cada celda por el NOMBRE de la columna
        Dim nombre = filaSeleccionada.Cells("Nombre").Value.ToString
        Dim edad = filaSeleccionada.Cells("Edad").Value.ToString
        Dim ciudad = filaSeleccionada.Cells("Ciudad").Value.ToString

        MessageBox.Show(
            "Nombre: " & nombre &
            vbCrLf & "Edad: " & edad &
            vbCrLf & "Ciudad: " & ciudad
        )

    Else
        MessageBox.Show("No hay ninguna fila seleccionada.")
    End If

End Sub
```

**Diferencia clave:** Aquí se usa `Cells("Nombre")` para acceder a las celdas **por nombre de columna**, en lugar de usar el índice numérico. Esto hace el código más legible.

```
filaSeleccionada.Cells("Nombre").Value  →  Accede por NOMBRE
filaSeleccionada.Cells(0).Value         →  Accede por ÍNDICE (también funciona)
```

---

## 9. 🔄 Form2 — CRUD completo (Buscar, Actualizar, Eliminar)

**CRUD** = **C**reate (Crear), **R**ead (Leer), **U**pdate (Actualizar), **D**elete (Eliminar).

El `Form2` implementa operaciones de lectura, actualización y eliminación sobre el arreglo de personas.

### Controles del formulario

| Control | Tipo | Función |
|---|---|---|
| `dgvPersonas` | `DataGridView` | Tabla de datos |
| `txtNombreBuscar` | `TextBox` | Campo para escribir el nombre a buscar/eliminar |
| `btnBuscar` | `Button` | Busca una persona por nombre |
| `btnEliminar` | `Button` | Elimina una persona del arreglo |
| `txtNombre` | `TextBox` | Campo para editar el nombre |
| `txtEdad` | `TextBox` | Campo para editar la edad |
| `txtCiudad` | `TextBox` | Campo para editar la ciudad |
| `btnActualizar` | `Button` | Actualiza los datos de la persona buscada |

### Declaración y carga inicial

```vb
Public Class Form2

    ' Arreglo fijo de 5 posiciones (índices 0 a 4)
    Private personas(4) As Persona

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Se cargan solo 3 personas; las posiciones 3 y 4 quedan vacías (Nothing)
        personas(0) = New Persona("Ana", 25, "Guatemala")
        personas(1) = New Persona("Luis", 30, "México")
        personas(2) = New Persona("Maria", 22, "El Salvador")

        MostrarDatos()  ' Se llama a la función que llena el DataGridView

    End Sub
```

> 📝 **Nota:** `Handles MyBase.Load` significa que este método se ejecuta **automáticamente** cuando el formulario se abre. Es el evento "Load" del formulario.

### Función `ObtenerCantidad` — Contar elementos no vacíos

```vb
Private Function ObtenerCantidad() As Integer

    Dim contador As Integer = 0

    For i As Integer = 0 To personas.Length - 1
        If personas(i) IsNot Nothing Then
            contador += 1
        End If
    Next

    Return contador

End Function
```

**¿Por qué es necesaria?** Porque el arreglo tiene 5 posiciones, pero no todas tienen datos. Esta función cuenta **solo las posiciones que tienen un objeto** (no son `Nothing`).

```
Arreglo:  [Ana] [Luis] [María] [Nothing] [Nothing]
Índice:     0      1      2        3          4

ObtenerCantidad() devuelve → 3
```

| Concepto | Explicación |
|---|---|
| `Function` | A diferencia de `Sub`, una `Function` **devuelve un valor**. |
| `As Integer` | El tipo de dato que retorna (un número entero). |
| `Return contador` | Devuelve el resultado al código que llamó a la función. |
| `IsNot Nothing` | Verifica que la posición del arreglo **sí tenga** un objeto. |
| `contador += 1` | Equivale a `contador = contador + 1`. |

### Función `MostrarDatos` — Refrescar el DataGridView

```vb
Private Sub MostrarDatos()

    ' Limpiar la tabla
    dgvPersonas.Rows.Clear()
    dgvPersonas.Columns.Clear()
    dgvPersonas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill

    ' Crear columnas
    dgvPersonas.Columns.Add("Nombre", "Nombre")
    dgvPersonas.Columns.Add("Edad", "Edad")
    dgvPersonas.Columns.Add("Ciudad", "Ciudad")

    ' Recorrer solo las posiciones con datos
    For i As Integer = 0 To ObtenerCantidad() - 1
        If personas(i) IsNot Nothing Then
            dgvPersonas.Rows.Add(
                personas(i).Nombre,
                personas(i).Edad,
                personas(i).Ciudad
            )
        End If
    Next

End Sub
```

> Esta función se reutiliza cada vez que los datos cambian (después de actualizar o eliminar) para **refrescar** la tabla.

### Botón "Buscar" — Encontrar una persona por nombre

```vb
Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

    Dim nombreBuscado As String = txtNombreBuscar.Text
    Dim encontrado As Boolean = False

    For i As Integer = 0 To ObtenerCantidad() - 1

        ' Comparar en minúsculas para que no importe si escriben "ana", "Ana" o "ANA"
        If personas(i).Nombre.ToLower() = nombreBuscado.ToLower() Then

            ' Llenar los TextBox con los datos encontrados
            txtNombre.Text = personas(i).Nombre
            txtEdad.Text = personas(i).Edad.ToString()
            txtCiudad.Text = personas(i).Ciudad

            encontrado = True
            Exit For  ' Salir del For porque ya encontramos lo que buscábamos

        End If

    Next

    If Not encontrado Then
        MessageBox.Show("Persona no encontrada.")
    End If

End Sub
```

**Flujo:**

```
1. Leer lo que el usuario escribió en txtNombreBuscar
2. Recorrer el arreglo comparando nombres (sin importar mayúsculas/minúsculas)
3. Si se encuentra → llenar los TextBox de edición con los datos
4. Si no se encuentra → mostrar mensaje de error
```

| Concepto | Explicación |
|---|---|
| `ToLower()` | Convierte el texto a minúsculas para comparar sin importar mayúsculas |
| `Exit For` | Sale del ciclo `For` inmediatamente |
| `Boolean` | Tipo de dato que solo puede ser `True` o `False` |

### Botón "Actualizar" — Modificar datos

```vb
Private Sub btnActualizar_Click(sender As Object, e As EventArgs) Handles btnActualizar.Click

    Dim nombreBuscado As String = txtNombreBuscar.Text

    For i As Integer = 0 To ObtenerCantidad() - 1

        If personas(i).Nombre.ToLower() = nombreBuscado.ToLower() Then

            ' Modificar las propiedades del objeto directamente
            personas(i).Nombre = txtNombre.Text
            personas(i).Edad = Integer.Parse(txtEdad.Text)
            personas(i).Ciudad = txtCiudad.Text

            MessageBox.Show("Datos actualizados.")
            MostrarDatos()  ' Refrescar la tabla para ver los cambios
            Exit Sub        ' Salir del método completo

        End If

    Next

    MessageBox.Show("Persona no encontrada.")

End Sub
```

| Concepto | Explicación |
|---|---|
| `Integer.Parse(txtEdad.Text)` | Convierte texto `"25"` a número entero `25` |
| `Exit Sub` | Sale del método completo (no solo del `For`) |
| `MostrarDatos()` | Se llama para **refrescar** la tabla con los datos nuevos |

> 💡 **Concepto clave:** Cuando modificamos `personas(i).Nombre`, estamos cambiando el **objeto en el arreglo**, no el DataGridView directamente. Por eso llamamos `MostrarDatos()` para que la tabla refleje los cambios.

### Botón "Eliminar" — Quitar una persona del arreglo

```vb
Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click

    Dim nombreBuscado As String = txtNombreBuscar.Text
    Dim posicion As Integer = -1

    ' PASO 1: Buscar la posición de la persona
    For i As Integer = 0 To ObtenerCantidad() - 1
        If personas(i).Nombre.ToLower() = nombreBuscado.ToLower() Then
            posicion = i
            Exit For
        End If
    Next

    ' Si no se encontró, la posición sigue siendo -1
    If posicion = -1 Then
        MessageBox.Show("Persona no encontrada.")
        Exit Sub
    End If

    ' PASO 2: Desplazar todos los elementos una posición a la izquierda
    For i As Integer = posicion To ObtenerCantidad() - 2
        personas(i) = personas(i + 1)
    Next

    ' PASO 3: Limpiar la última posición (que quedó duplicada)
    personas(ObtenerCantidad() - 1) = Nothing

    MessageBox.Show("Persona eliminada.")
    MostrarDatos()

End Sub
```

**¿Cómo funciona la eliminación?** Los arreglos en VB.NET tienen un tamaño fijo, así que no puedes "quitar" una posición. En su lugar, se **desplazan** los elementos:

```
Estado inicial:     [Ana] [Luis] [María] [Nothing] [Nothing]

Eliminar "Luis" (posición 1):

Paso 1 - Desplazar:  [Ana] [María] [María] [Nothing] [Nothing]
                             ↑ María se copió a la izquierda

Paso 2 - Limpiar:    [Ana] [María] [Nothing] [Nothing] [Nothing]
                                      ↑ Se limpia la posición duplicada

Resultado final:     [Ana] [María] [Nothing] [Nothing] [Nothing]
```

---

## 10. 🎨 Form3 — Controles dinámicos con Panel y FlowLayoutPanel

El `Form3` muestra cómo crear controles de la interfaz (paneles, imágenes, etiquetas) **desde el código** en lugar de arrastrarlos en el diseñador visual.

### Controles del formulario

| Control | Tipo | Función |
|---|---|---|
| `panelContenedor` | `Panel` | Contenedor donde se agregan tarjetas manualmente (posición manual) |
| `flpPanel` | `FlowLayoutPanel` | Contenedor que acomoda las tarjetas automáticamente |

### Carga de datos con URL de imagen

```vb
Private Sub FormGaleria_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    personas(0) = New Persona(
        "Ana", 25, "Guatemala",
        "https://randomuser.me/api/portraits/women/1.jpg"    ' ← Se usa el parámetro Optional
    )
    ' ... más personas ...

    CrearControlesDinamicos()

End Sub
```

> Aquí se usa el cuarto parámetro del constructor (`url`), que en Form1 y Form2 no se usaba porque era `Optional`.

### Crear controles dinámicamente (Panel normal)

```vb
Private Sub CrearControlesDinamicos()

    Dim posicionY As Integer = 10   ' Posición vertical inicial

    For i As Integer = 0 To personas.Length - 1

        If personas(i) IsNot Nothing Then

            ' ===== CREAR UN PANEL (tarjeta) POR CADA PERSONA =====
            Dim panelPersona As New Panel
            panelPersona.Width = 200
            panelPersona.Height = 300
            panelPersona.Left = 10          ' Posición horizontal
            panelPersona.Top = posicionY    ' Posición vertical (cambia en cada iteración)
            panelPersona.BorderStyle = BorderStyle.FixedSingle

            ' ===== CREAR UN PICTUREBOX (imagen) =====
            Dim pic As New PictureBox
            pic.Width = 180
            pic.Height = 150
            pic.Top = 10
            pic.Left = 10
            pic.SizeMode = PictureBoxSizeMode.StretchImage

            Try
                pic.Load(personas(i).UrlImagen)    ' Intentar cargar la imagen desde la URL
            Catch ex As Exception
                pic.BackColor = Color.Gray          ' Si falla, mostrar un fondo gris
            End Try

            ' ===== CREAR LABELS (etiquetas de texto) =====
            Dim lblNombre As New Label
            lblNombre.Text = "Nombre: " & personas(i).Nombre
            lblNombre.Top = 170
            lblNombre.Left = 10
            lblNombre.AutoSize = True

            ' ... similar para lblEdad y lblCiudad ...

            ' ===== AGREGAR todo al panel de la persona =====
            panelPersona.Controls.Add(pic)
            panelPersona.Controls.Add(lblNombre)
            panelPersona.Controls.Add(lblEdad)
            panelPersona.Controls.Add(lblCiudad)

            ' ===== AGREGAR el panel al contenedor principal =====
            panelContenedor.Controls.Add(panelPersona)

            posicionY += 320   ' Mover la posición Y hacia abajo para la siguiente tarjeta

        End If
    Next
```

### Visualización de la estructura de controles

```
panelContenedor (Panel principal con scroll)
├── panelPersona 1 (Top = 10)
│   ├── PictureBox (foto de Ana)
│   ├── Label "Nombre: Ana"
│   ├── Label "Edad: 25"
│   └── Label "Ciudad: Guatemala"
│
├── panelPersona 2 (Top = 330)
│   ├── PictureBox (foto de Luis)
│   ├── Label "Nombre: Luis"
│   ├── Label "Edad: 30"
│   └── Label "Ciudad: México"
│
└── panelPersona 3 (Top = 650)
    ├── PictureBox (foto de María)
    ├── Label "Nombre: Maria"
    ├── Label "Edad: 22"
    └── Label "Ciudad: El Salvador"
```

### Panel normal vs FlowLayoutPanel

El Form3 crea las mismas tarjetas dos veces, pero en contenedores diferentes:

| Panel normal | FlowLayoutPanel |
|---|---|
| Debes calcular `Top` y `Left` manualmente | Se acomodan **automáticamente** uno tras otro |
| `panelPersona.Top = posicionY` | No necesitas definir posición |
| `posicionY += 320` (mover manualmente) | El FlowLayoutPanel lo hace solo |
| Más control sobre la posición exacta | Más fácil y menos código |

### Try-Catch — Manejo de errores

```vb
Try
    pic.Load(personas(i).UrlImagen)    ' Intentar cargar la imagen
Catch ex As Exception
    pic.BackColor = Color.Gray          ' Si hay error, mostrar gris
End Try
```

| Bloque | Función |
|---|---|
| `Try` | "Intenta ejecutar este código" |
| `Catch ex As Exception` | "Si ocurre un error, ejecuta este código en su lugar" |
| `End Try` | Fin del bloque de manejo de errores |

> Sin `Try-Catch`, si una imagen no se puede cargar (URL inválida, sin internet, etc.), el programa se detendrá con un error. Con `Try-Catch`, el error se "atrapa" y el programa continúa normalmente.

---

## 11. 📌 Resumen de conceptos clave

### Programación Orientada a Objetos (POO)

| Concepto | En este proyecto | Descripción |
|---|---|---|
| **Clase** | `Persona` | Molde que define Nombre, Edad, Ciudad, UrlImagen |
| **Objeto** | `New Persona("Ana", 25, "Guatemala")` | Instancia concreta con datos reales |
| **Propiedad** | `Nombre`, `Edad`, `Ciudad`, `UrlImagen` | Características del objeto |
| **Constructor** | `Sub New(...)` | Método que inicializa el objeto al crearlo |
| **`Me`** | `Me.Nombre = nombre` | Referencia al objeto actual |

### DataGridView

| Operación | Código |
|---|---|
| Limpiar columnas | `dgv.Columns.Clear()` |
| Limpiar filas | `dgv.Rows.Clear()` |
| Agregar columna | `dgv.Columns.Add("nombre", "Encabezado")` |
| Agregar fila | `dgv.Rows.Add(valor1, valor2, valor3)` |
| Celda seleccionada | `dgv.CurrentCell.Value` |
| Fila seleccionada | `dgv.CurrentRow` |
| Celda por nombre | `fila.Cells("Nombre").Value` |
| Ajustar columnas | `dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill` |

### Arreglos

| Operación | Código |
|---|---|
| Declarar | `Private personas(2) As Persona` |
| Asignar | `personas(0) = New Persona(...)` |
| Acceder | `personas(i).Nombre` |
| Tamaño | `personas.Length` |
| Verificar vacío | `personas(i) IsNot Nothing` |

### Flujo general de la aplicación

```
Form1 (Básico)                Form2 (CRUD)              Form3 (Dinámico)
┌─────────────────┐     ┌─────────────────────┐    ┌──────────────────────┐
│ Cargar datos     │     │ Buscar por nombre    │    │ Crear controles      │
│ Mostrar en tabla │ ──► │ Actualizar datos     │ ──►│ desde código         │
│ Leer celdas/filas│     │ Eliminar del arreglo │    │ Mostrar imágenes     │
└─────────────────┘     └─────────────────────┘    └──────────────────────┘
```

---

## 12. 📖 Glosario

| Término | Definición |
|---|---|
| **`Dim`** | Declara una variable local |
| **`Private`** | Accesible solo dentro de la clase actual |
| **`Public`** | Accesible desde cualquier parte |
| **`New`** | Crea una nueva instancia (objeto) de una clase |
| **`Nothing`** | Valor nulo (sin datos, vacío) |
| **`IsNot Nothing`** | Verifica que algo NO sea nulo |
| **`Sub`** | Procedimiento que no devuelve valor |
| **`Function`** | Procedimiento que SÍ devuelve un valor |
| **`Handles`** | Conecta un método con un evento (clic, carga, etc.) |
| **`Me`** | Referencia al objeto actual |
| **`Exit For`** | Sale de un ciclo `For` |
| **`Exit Sub`** | Sale de un procedimiento `Sub` |
| **`vbCrLf`** | Salto de línea en texto |
| **`ToString()`** | Convierte cualquier valor a texto |
| **`Integer.Parse()`** | Convierte texto a número entero |
| **`ToLower()`** | Convierte texto a minúsculas |
| **`Try-Catch`** | Manejo de errores (evita que el programa se detenga) |
| **`Optional`** | Parámetro no obligatorio con valor por defecto |
| **`Controls.Add()`** | Agrega un control hijo a un contenedor |

---

> 📝 **Nota final:** Este proyecto usa arreglos fijos. En proyectos más avanzados se usan **`List(Of T)`** que permiten agregar y quitar elementos sin preocuparse por el tamaño fijo. Pero los arreglos son fundamentales para entender cómo funcionan las colecciones en programación.

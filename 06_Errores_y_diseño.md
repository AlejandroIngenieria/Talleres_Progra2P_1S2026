# try-catch / try-catch-finally

La idea será:

* Un `TextBox` donde el usuario escribe su edad.
* Un botón que convierte el texto a número.
* Mostrar un mensaje personalizado.
* Manejar errores comunes:

  * Campo vacío
  * Texto en vez de número
  * Número negativo
* Usar `Finally` para limpiar el formulario.

---

## 🟢 Ejemplo sencillo: Validar Edad

### 🎯 Objetivo

Aprender:

* Qué hace `Try`
* Cómo funcionan los `Catch`
* Para qué sirve `Finally`

---

### 🧩 Diseño del Form

Agrega al formulario:

* `TextBox` → `txtEdad`
* `Button` → `btnValidar`
* `Label` → "Ingrese su edad:"

---

### 💻 Código Completo

```vb
Public Class FormEjemploBasico

    Private Sub btnValidar_Click(sender As Object, e As EventArgs) Handles btnValidar.Click

        Try
            ' ----------------------------------------------------
            ' Intentamos convertir el texto a número entero
            ' Si el usuario escribe letras, aquí ocurre el error
            ' ----------------------------------------------------
            Dim edad As Integer = Integer.Parse(txtEdad.Text)

            ' ----------------------------------------------------
            ' Validación lógica (no es excepción del sistema)
            ' Lanzamos una excepción manualmente
            ' ----------------------------------------------------
            If edad < 0 Then
                Throw New Exception("La edad no puede ser negativa.")
            End If

            ' ----------------------------------------------------
            ' Si todo sale bien, mostramos mensaje
            ' ----------------------------------------------------
            If edad >= 18 Then
                MessageBox.Show("Eres mayor de edad.")
            Else
                MessageBox.Show("Eres menor de edad.")
            End If

        Catch ex As FormatException
            ' Ocurre si escriben texto en vez de número
            MessageBox.Show("Error: Debes escribir solo números.")

        Catch ex As Exception
            ' Captura errores personalizados u otros errores
            MessageBox.Show("Error: " & ex.Message)

        Finally
            ' ----------------------------------------------------
            ' Este bloque SIEMPRE se ejecuta
            ' Haya error o no
            ' ----------------------------------------------------
            txtEdad.Clear()
            txtEdad.Focus()
        End Try

    End Sub

End Class
```

---

## 🔎 ¿Qué está pasando aquí?

### 🔹 1. `Try`

Contiene el código que **puede fallar**.

```vb
Dim edad As Integer = Integer.Parse(txtEdad.Text)
```

Si el usuario escribe `"hola"`, el programa genera un:

* `FormatException`

---

### 🔹 2. `Catch`

Atrapa el error y evita que el programa se cierre.

Sin `Try-Catch`, la aplicación mostraría un error crítico y podría detenerse.

---

### 🔹 3. `Throw`

Aquí estamos generando un error manual:

```vb
Throw New Exception("La edad no puede ser negativa.")
```

Esto es útil cuando:

* El sistema no detecta el error automáticamente
* Pero tú quieres forzar una validación

---

### 🔹 4. `Finally`

Siempre se ejecuta.
Perfecto para:

* Limpiar campos
* Cerrar conexiones
* Liberar recursos

---

## 🟡 Extra: Validar Solo Números en el TextBox

Si quieres evitar que escriban letras desde el inicio:

```vb
Private Sub txtEdad_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEdad.KeyPress

    ' Permite solo números y Backspace
    If Not Char.IsDigit(e.KeyChar) AndAlso
       Not Char.IsControl(e.KeyChar) Then

        e.Handled = True
        MessageBox.Show("Solo se permiten números.")
    End If

End Sub
```

Esto reduce errores antes de llegar al `Try`.

---

## 📌 Resumen Simple

| Parte   | Función                 |
| ------- | ----------------------- |
| Try     | Código que puede fallar |
| Catch   | Maneja el error         |
| Finally | Se ejecuta siempre      |
| Throw   | Genera un error manual  |

# Diseño

## Código para colocar una imagen de fondo desde una URL (recomiendo pinterest)

```vb
' Imagen de fondo del form desde url
Dim url As String = "https://i.pinimg.com/1200x/6d/a6/43/6da64326bbd71d4552a2226b58259661.jpg"
Dim request As System.Net.WebRequest = System.Net.WebRequest.Create(url)
Dim response As System.Net.WebResponse = request.GetResponse()
Dim stream As System.IO.Stream = response.GetResponseStream()
BackgroundImage = System.Drawing.Image.FromStream(stream)
BackgroundImageLayout = ImageLayout.Stretch
```

# Controles dinamicos - Ejemplo 1
### Objetivo de la práctica

Al finalizar esta será capaz de:

* Comprender que los controles son objetos.
* Crear controles en tiempo de ejecución.
* Configurar propiedades desde código.
* Asignar eventos dinámicamente.
* Agregar controles a un contenedor.
* Utilizar la propiedad `Tag` para almacenar información adicional.

---

### Fundamento Teórico

En **Windows Forms**, todos los controles (Button, TextBox, Label, etc.) heredan de:

```vb
System.Windows.Forms.Control
```

Esto significa que comparten propiedades comunes como:

* `Name`
* `Text`
* `Location`
* `Size`
* `BackColor`
* `Anchor`
* `Dock`
* `Tag`
* `Parent`

Cuando usamos el diseñador visual, Visual Studio genera código automáticamente.
En esta práctica, nosotros escribiremos ese código manualmente.

---

### Propiedades Fundamentales de un Control

| Propiedad  | Descripción                             |
| ---------- | --------------------------------------- |
| `Name`     | Identificador interno del control       |
| `Text`     | Texto visible para el usuario           |
| `Location` | Posición (X,Y) dentro del contenedor    |
| `Size`     | Dimensiones en píxeles (Ancho, Alto)    |
| `Parent`   | Contenedor que lo aloja                 |
| `Anchor`   | Mantiene posición al redimensionar      |
| `Dock`     | Ajuste automático dentro del contenedor |
| `Tag`      | Almacena datos adicionales              |

---

### Preparación del Proyecto

1. Crear un proyecto **Windows Forms App (.NET Framework)**.
2. Agregar un botón al formulario llamado:

```
btnCrear
```

3. Cambiar su propiedad `Text` a:

```
Crear Botón
```

---

### Código Completo de la Práctica

```vb
Public Class FormDinamico

    ' Variable que permitirá numerar los botones creados
    Private contador As Integer = 1

    Private Sub btnCrear_Click(sender As Object, e As EventArgs) Handles btnCrear.Click

        ' ==========================================
        ' PASO 1: Instanciar el control
        ' ==========================================
        Dim nuevoBoton As New Button()

        ' ==========================================
        ' PASO 2: Configurar propiedades básicas
        ' ==========================================
        nuevoBoton.Name = "btnExtra" & contador
        nuevoBoton.Text = "Botón " & contador

        ' Tamaño del botón
        nuevoBoton.Size = New Size(120, 40)

        ' Posición en el formulario
        nuevoBoton.Location = New Point(20, 20 + (contador * 50))

        ' Color de fondo
        nuevoBoton.BackColor = Color.LightBlue

        ' Mantener alineado arriba-izquierda al redimensionar
        nuevoBoton.Anchor = AnchorStyles.Top Or AnchorStyles.Left

        ' Guardar información adicional
        nuevoBoton.Tag = "Identificador interno: " & contador

        ' ==========================================
        ' PASO 3: Asignar eventos dinámicamente
        ' ==========================================
        AddHandler nuevoBoton.Click, AddressOf EventoComun_Click
        AddHandler nuevoBoton.MouseDown, AddressOf EventoClickDerecho

        ' ==========================================
        ' PASO 4: Agregar el control al contenedor
        ' ==========================================
        Me.Controls.Add(nuevoBoton)

        contador += 1

    End Sub


    ' Evento que manejará el click
    Private Sub EventoComun_Click(sender As Object, e As EventArgs)

        ' Convertimos sender en Button
        Dim botonPresionado As Button = DirectCast(sender, Button)

        MessageBox.Show(
            "Has presionado: " & botonPresionado.Text &
            vbCrLf & botonPresionado.Tag.ToString()
        )

    End Sub


    ' Evento que permite eliminar el botón con clic derecho
    Private Sub EventoClickDerecho(sender As Object, e As MouseEventArgs)

        If e.Button = MouseButtons.Right Then
            Dim boton As Button = DirectCast(sender, Button)
            Me.Controls.Remove(boton)
        End If

    End Sub

End Class
```

---

## Explicación Paso a Paso

### 🔹 Paso 1 – Instanciación

```vb
Dim nuevoBoton As New Button()
```

Aquí se crea el objeto en memoria.
Todavía no es visible en pantalla.

---

### 🔹 Paso 2 – Configuración

Se asignan propiedades antes de mostrarlo:

* `Name` → Identificador interno.
* `Text` → Texto visible.
* `Size` → Dimensiones.
* `Location` → Posición.
* `Tag` → Información adicional.
* `Anchor` → Comportamiento al redimensionar.

---

### 🔹 Paso 3 – Asignación de Eventos

```vb
AddHandler nuevoBoton.Click, AddressOf EventoComun_Click
```

Permite que múltiples botones compartan el mismo método.

Concepto clave:

* `sender` representa el objeto que disparó el evento.

---

### 🔹 Paso 4 – Agregar al Contenedor

```vb
Me.Controls.Add(nuevoBoton)
```

En este momento el botón se vuelve visible.

Si no se agrega al contenedor, no aparece.

---

### Conceptos Importantes para Reflexión

✔ Los controles son objetos.
✔ Se crean dinámicamente en tiempo de ejecución.
✔ Los eventos pueden asignarse desde código.
✔ `sender` permite identificar qué control ejecutó el evento.
✔ `Tag` permite asociar datos personalizados.
✔ Todo control necesita un contenedor.

---

### Ejercicios Propuestos

1. Modificar el ejemplo para crear `TextBox` dinámicos.
2. Hacer que cada botón tenga un color aleatorio.
3. Crear botones dentro de un `Panel` en lugar del formulario.
4. Agregar un botón que elimine todos los controles dinámicos.

---

### Conclusión

Crear controles dinámicamente:

* Refuerza el modelo orientado a objetos.
* Permite interfaces flexibles.
* Es útil cuando la interfaz depende de datos.
* Es una práctica común en aplicaciones empresariales.

# Controles dinamicos - Ejemplo 2

### 🎯 Objetivo

* Crear controles desde código.
* Configurar propiedades básicas.
* Usar valores ingresados por el usuario para posicionar controles.
* Comprender que los controles son objetos.

---

### Preparación del Formulario

En el diseñador agrega:

#### 🔹 Controles fijos (desde el diseñador)

| Control | Name            | Text          |
| ------- | --------------- | ------------- |
| Button  | btnCrearBoton   | Crear Botón   |
| Button  | btnCrearLabel   | Crear Label   |
| Button  | btnCrearTextBox | Crear TextBox |
| Label   | lblX            | Posición X    |
| Label   | lblY            | Posición Y    |
| TextBox | txtX            | (vacío)       |
| TextBox | txtY            | (vacío)       |

Estos controles NO son dinámicos.
Servirán para generar otros controles.

---

### Código Completo

```vb
Public Class FormDinamico

    ' Contadores para identificar los controles creados
    Private contadorBoton As Integer = 1
    Private contadorLabel As Integer = 1
    Private contadorTextBox As Integer = 1

    ' ==========================================
    ' MÉTODO PARA OBTENER POSICIÓN
    ' ==========================================
    Private Function ObtenerPosicion() As Point

        Dim x As Integer
        Dim y As Integer

        ' Intentamos convertir los valores ingresados
        If Integer.TryParse(txtX.Text, x) AndAlso
           Integer.TryParse(txtY.Text, y) Then

            Return New Point(x, y)

        Else
            MessageBox.Show("Ingrese valores numéricos válidos para la posición.")
            Return New Point(0, 0)
        End If

    End Function


    ' ==========================================
    ' CREAR BOTÓN DINÁMICO
    ' ==========================================
    Private Sub btnCrearBoton_Click(sender As Object, e As EventArgs) Handles btnCrearBoton.Click

        Dim posicion As Point = ObtenerPosicion()

        ' 1. Instanciar el botón
        Dim nuevoBoton As New Button()

        ' 2. Configurar propiedades
        nuevoBoton.Name = "btnDinamico" & contadorBoton
        nuevoBoton.Text = "Botón " & contadorBoton
        nuevoBoton.Size = New Size(120, 40)
        nuevoBoton.Location = posicion

        ' 3. Agregar al formulario
        Me.Controls.Add(nuevoBoton)

        contadorBoton += 1

    End Sub


    ' ==========================================
    ' CREAR LABEL DINÁMICO
    ' ==========================================
    Private Sub btnCrearLabel_Click(sender As Object, e As EventArgs) Handles btnCrearLabel.Click

        Dim posicion As Point = ObtenerPosicion()

        Dim nuevoLabel As New Label()

        nuevoLabel.Name = "lblDinamico" & contadorLabel
        nuevoLabel.Text = "Label " & contadorLabel
        nuevoLabel.AutoSize = True
        nuevoLabel.Location = posicion

        Me.Controls.Add(nuevoLabel)

        contadorLabel += 1

    End Sub


    ' ==========================================
    ' CREAR TEXTBOX DINÁMICO
    ' ==========================================
    Private Sub btnCrearTextBox_Click(sender As Object, e As EventArgs) Handles btnCrearTextBox.Click

        Dim posicion As Point = ObtenerPosicion()

        Dim nuevoTextBox As New TextBox()

        nuevoTextBox.Name = "txtDinamico" & contadorTextBox
        nuevoTextBox.Size = New Size(150, 25)
        nuevoTextBox.Location = posicion

        Me.Controls.Add(nuevoTextBox)

        contadorTextBox += 1

    End Sub

End Class
```

---

## Explicación Paso a Paso

---

### 🔹 1. ¿Qué hace `New Button()`?

```vb
Dim nuevoBoton As New Button()
```

Aquí se crea un objeto en memoria.
Aún no aparece en pantalla.

---

### 🔹 2. ¿Qué hace `Location`?

```vb
nuevoBoton.Location = New Point(x, y)
```

Define la posición:

* X → Horizontal
* Y → Vertical

Se obtiene de los TextBox `txtX` y `txtY`.

---

### 🔹 3. ¿Qué hace `Me.Controls.Add()`?

```vb
Controls.Add(nuevoBoton)
```

Agrega el control al formulario.

Sin esta línea:
El control NO se muestra.

---

### 🔹 4. ¿Por qué usamos contadores?

Para que cada control tenga un nombre distinto:

```
btnDinamico1
btnDinamico2
btnDinamico3
```

Evita duplicados.

---

### Conceptos Fundamentales que Deben Entender

✔ Un control es un objeto.
✔ Primero se crea (instancia).
✔ Luego se configuran propiedades.
✔ Finalmente se agrega al contenedor.
✔ Las posiciones se definen con coordenadas.

---

###  Flujo Mental Correcto

Cuando presionan "Crear Botón":

1. Se leen las posiciones.
2. Se crea el objeto Button.
3. Se configuran propiedades.
4. Se agrega al formulario.
5. Aparece en pantalla.

---

### Conclusión

Este ejercicio demuestra que:

* El diseñador solo genera código automáticamente.
* Podemos crear interfaces dinámicas.
* Los controles no son "mágicos", son objetos.
* La programación visual sigue siendo programación orientada a objetos.

# Controles dinamicos - Ejemplo 3
### Objetivo

* Declarar un arreglo unidimensional.
* Recorrerlo con un ciclo `For`.
* Crear botones dinámicamente.
* Asignar eventos desde código.
* Usar `sender` para identificar qué botón fue presionado.

---

### Preparación del Formulario

Desde el diseñador agrega:

| Control | Name       | Text            |
| ------- | ---------- | --------------- |
| Button  | btnGenerar | Generar Botones |

Nada más.
Los demás botones se crearán desde código.

---

### Concepto Importante: Arreglo Unidimensional

Un arreglo es una estructura que almacena varios valores del mismo tipo.

Ejemplo:

```vb id="x1nhre"
Dim nombres() As String = {"Ana", "Luis", "Carlos", "María", "Sofía"}
```

Esto crea un arreglo con 5 nombres.

---

### Código Completo del Ejemplo

```vb
Public Class FormArregloBotones

    ' Arreglo unidimensional de nombres
    Private nombres() As String = {"Ana", "Luis", "Carlos", "María", "Sofía"}

    Private Sub btnGenerar_Click(sender As Object, e As EventArgs) Handles btnGenerar.Click

        Dim posicionY As Integer = 60

        ' Recorremos el arreglo con un ciclo For
        For i As Integer = 0 To nombres.Length - 1

            ' 1️⃣ Crear el botón
            Dim nuevoBoton As New Button()

            ' 2️⃣ Configurar propiedades
            nuevoBoton.Size = New Size(120, 40)
            nuevoBoton.Location = New Point(30, posicionY)
            nuevoBoton.Text = nombres(i)

            ' Guardamos el nombre en la propiedad Tag
            nuevoBoton.Tag = nombres(i)

            ' 3️⃣ Asignar evento
            AddHandler nuevoBoton.Click, AddressOf Saludar

            ' 4️⃣ Agregar al formulario
            Me.Controls.Add(nuevoBoton)

            ' Aumentar posición vertical
            posicionY += 50

        Next

    End Sub


    ' Método que se ejecuta cuando se presiona cualquier botón
    Private Sub Saludar(sender As Object, e As EventArgs)

        Dim botonPresionado As Button = DirectCast(sender, Button)

        MessageBox.Show("Hola " & botonPresionado.Tag.ToString() & " 👋")

    End Sub

End Class
```

---

### Explicación Paso a Paso

---

#### 🔹 1. Declaración del arreglo

```vb
Private nombres() As String = {"Ana", "Luis", "Carlos", "María", "Sofía"}
```

Aquí se almacenan los nombres que usaremos para crear los botones.

---

#### 🔹 2. Recorrer el arreglo

```vb
For i As Integer = 0 To nombres.Length - 1
```

* `Length` indica cuántos elementos tiene el arreglo.
* Restamos 1 porque los arreglos comienzan en posición 0.

---

#### 🔹 3. Crear el botón

```vb
Dim nuevoBoton As New Button()
```

Se crea el objeto en memoria.

---

#### 🔹 4. Usar la propiedad `Tag`

```vb
nuevoBoton.Tag = nombres(i)
```

`Tag` permite guardar información adicional.
En este caso, el nombre correspondiente.

---

#### 🔹 5. Asignar evento dinámico

```vb
AddHandler nuevoBoton.Click, AddressOf Saludar
```

Todos los botones usarán el mismo método `Saludar`.

---

#### 🔹 6. Identificar qué botón fue presionado

```vb
Dim botonPresionado As Button = DirectCast(sender, Button)
```

`sender` representa el botón que se presionó.

---

### Flujo de Ejecución

1. Se presiona "Generar Botones".
2. Se recorre el arreglo.
3. Se crea un botón por cada nombre.
4. Cada botón muestra un saludo al presionarlo.

---

### Resultado Esperado

Se crean botones como:

```
Ana
Luis
Carlos
María
Sofía
```

Al presionar "Carlos":

```
Hola Carlos 👋
```

---

### Conceptos que Refuerza Esta Práctica

✔ Arreglos unidimensionales
✔ Ciclo `For`
✔ Programación orientada a objetos
✔ Eventos dinámicos
✔ Uso de `sender`
✔ Propiedad `Tag`

---

### Ejercicios Propuestos

1. Agregar más nombres al arreglo.
2. Cambiar el color de cada botón.
3. Crear los botones horizontalmente en lugar de vertical.
4. Agregar un botón que elimine todos los botones creados.

---

### Conclusión

Este ejemplo demuestra que:

* Un arreglo puede generar una interfaz dinámica.
* Un solo método puede manejar múltiples eventos.
* Los controles pueden representar datos.
* La interfaz puede construirse desde datos almacenados.

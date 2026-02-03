# Parte 1: Sentencias Condicionales

### Ejercicio 1: Acceso al Sistema (If...Then...Else con Operador `And`)

* **Descripción:** Valida dos condiciones simultáneas (Usuario y Contraseña) para permitir el acceso.

* **Controles:** 2 `TextBox` (txtUsuario, txtPass), 1 `Button` (btnEntrar).

* **Código:**

```vbnet
Dim usuarioValido As String = "admin"
Dim passValida As String = "1234"

' Uso de operador lógico AND
If txtUsuario.Text = usuarioValido And txtPass.Text = passValida Then
    MsgBox("Acceso Concedido")
Else
    MsgBox("Usuario o contraseña incorrectos")
End If
```

### Ejercicio 2: Validación de Rango (If...ElseIf con Operador `Or`)

* **Descripción:** Verifica si un número está fuera de un rango permitido (menor a 0 o mayor a 100).

* **Controles:** 1 `TextBox` (txtValor), 1 `Button` (btnValidar).

* **Código:**

```vbnet
Dim numero As Integer = CInt(txtValor.Text)

' Uso de operador lógico OR para detectar extremos
If numero < 0 Or numero > 100 Then
    MsgBox("Error: El valor está fuera del rango permitido (0-100).")
ElseIf numero >= 0 And numero <= 50 Then
    MsgBox("Rango bajo.")
Else
    MsgBox("Rango alto.")
End If
```

### Ejercicio 3: Estado de Interruptor (Variable Booleana + IIf)

* **Descripción:** Cambia el texto de una etiqueta basándose en el estado de una variable booleana (bandera), alternando entre "Encendido" y "Apagado".

* **Controles:** 1 `Button` (btnSwitch), 1 `Label` (lblEstado).

* **Código:**

```vbnet
' Variable booleana a nivel de formulario
Dim esEncendido As Boolean = False

Private Sub btnSwitch_Click(...) Handles btnSwitch.Click
    esEncendido = Not esEncendido ' Invierte el valor booleano

    ' Uso de IIf para asignar texto según el booleano
    lblEstado.Text = IIf(esEncendido, "Sistema ENCENDIDO", "Sistema APAGADO")

    ' Cambio de color opcional usando la misma lógica
    lblEstado.ForeColor = IIf(esEncendido, Color.Green, Color.Red)
End Sub
```

### Ejercicio 4: Categoría de Cliente (Select Case)

* **Descripción:** Determina el descuento basado en una letra de categoría (A, B, C).

* **Controles:** 1 `TextBox` (txtCategoria), 1 `Button` (btnVerificar).

* **Código:**

```vbnet
Dim categoria As String = txtCategoria.Text.ToUpper()

Select Case categoria
    Case "A"
        MsgBox("Cliente VIP: 20% Descuento")
    Case "B"
        MsgBox("Cliente Frecuente: 10% Descuento")
    Case "C"
        MsgBox("Cliente Nuevo: 5% Descuento")
    Case Else
        MsgBox("Sin descuento asignado")
End Select
```

* * *

# Parte 2: Subrutinas y Funciones

### Ejercicio 5: Limpieza de Formulario (Subrutina)

* **Descripción:** Una subrutina que se llama desde un botón para limpiar múltiples campos de texto a la vez (reutilización de código de mantenimiento).

* **Controles:** 3 `TextBox` (txt1, txt2, txt3), 1 `Button` (btnLimpiar).

* **Código:**

```vbnet
' Declaración de la Subrutina
Sub LimpiarCampos()
    txt1.Text = ""
    txt2.Text = ""
    txt3.Text = ""
    txt1.Focus()
End Sub

' Llamada desde el botón
Private Sub btnLimpiar_Click(...) Handles btnLimpiar.Click
    LimpiarCampos()
End Sub
```

### Ejercicio 6: Calculadora de IVA (Función)

* **Descripción:** Una función que recibe un monto, calcula el impuesto y devuelve el total.

* **Controles:** 1 `TextBox` (txtSubtotal), 1 `Button` (btnCalcularTotal), 1 `Label` (lblTotal).

* **Código:**

```vbnet
' Declaración de la Función
Function CalcularTotalConImpuesto(ByVal monto As Double) As Double
    Const TASA_IMPUESTO As Double = 0.12 ' 12%
    Return monto * (1 + TASA_IMPUESTO)
End Function

' Llamada y uso del valor devuelto
Private Sub btnCalcularTotal_Click(...) Handles btnCalcularTotal.Click
    Dim subtotal As Double = CDbl(txtSubtotal.Text)
    Dim total As Double = CalcularTotalConImpuesto(subtotal)
    lblTotal.Text = "Total a pagar: " & total.ToString("C")
End Sub
```

* * *

# Parte 3: Módulos

### Ejercicio 7: Variables Globales y Funciones Compartidas (Module)

* **Descripción:** Crear un módulo aparte para almacenar una variable que sea accesible desde cualquier formulario del proyecto (ej. Nombre de Usuario Logueado) y una función matemática general.

* **Componentes:** 1 Archivo `Module` (ej. `MdlGeneral.vb`), 1 `Formulario`.

* **Código en Módulo (MdlGeneral.vb):**

```vbnet
Module MdlGeneral
    ' Variable accesible desde todo el proyecto
    Public UsuarioActual As String = ""

    ' Función accesible desde todo el proyecto
    Public Function ConvertirDolaresAQuetzales(dolares As Double) As Double
        Return dolares * 7.8
    End Function
End Module
```

* **Código en Formulario (Uso del Módulo):**

```vbnet
Private Sub btnGuardar_Click(...) Handles btnGuardar.Click
    ' Guardamos en la variable global del módulo
    MdlGeneral.UsuarioActual = txtUsuario.Text
    MsgBox("Bienvenido " & MdlGeneral.UsuarioActual)
End Sub
```

# Enunciado

### Distribuidora de Textiles "El Arcoíris"

**Instrucciones:**

Diseñe un formulario y desarrolle el código para un programa que controle las ventas de playeras por mayor. El programa debe leer el nombre del cliente y el número de NIT. Debe ingresar por cada color de playera (Roja, Azul, Verde, Negra) la cantidad de **decenas** (paquetes de 10 unidades). Si no se vende algún color, debe ingresar el valor 0.

**Precios por Decena:**

| **Color de Playera** | **Precio por Decena** |
| -------------------- | --------------------- |
| Roja                 | Q150.00               |
| Azul                 | Q165.50               |
| Verde                | Q140.00               |
| Negra                | Q175.25               |

**Requerimientos de Cálculo:**

1. **Subtotal:** Calcular el valor de las decenas solicitadas multiplicadas por su precio correspondiente.
   
   * _Ejemplo:_ 2 decenas Rojas y 1 decena Verde = $(150.00 * 2) + (140.00 * 1) = 440.00$.

2. **Transporte:** Se cobra un recargo del **5% (0.05)** sobre el Subtotal calculado anteriormente.

3. **Total a Pagar:** Suma del Subtotal más el recargo de Transporte.

**Salida de Datos:**

El programa debe mostrar en pantalla:

* El Subtotal (antes del transporte).

* El monto calculado del Transporte.

* El Total Final a pagar.

**Botones / Métodos a programar:**

1. **Calcular:** Ejecuta las operaciones matemáticas y muestra los resultados.

2. **Limpiar:** Reinicia todos los cuadros de texto y etiquetas a sus valores iniciales.

3. **Confirmar:** Despliega un mensaje emergente (_MessageBox_) confirmando que la venta fue procesada con el nombre del cliente y el monto final. _(Método adicional solicitado)_.

4. **Salir:** Cierra la aplicación.

# Pasos para resolver

## Fase de diseño: ¿Qué necesito?

* **Labels (Etiquetas):**
  
  * Para títulos: "Distribuidora El Arcoíris", "Nombre", "NIT".
  
  * Para los productos: "Decenas Rojas", "Decenas Azules", "Decenas Verdes", "Decenas Negras".
  
  * Para resultados: "Subtotal:", "Transporte (5%):", "Total a Pagar:".
  
  * **Importante:** 3 Labels vacíos para mostrar los resultados numéricos (ej. `lblSubtotal`, `lblTransporte`, `lblTotal`).

* **Textboxes (Cajas de texto):**
  
  * 2 para datos del cliente (`txtNombre`, `txtNit`).
  
  * 4 para ingresar las cantidades de decenas (`txtRojas`, `txtAzules`, `txtVerdes`, `txtNegras`).

* **Buttons (Botones):**
  
  * `btnCalcular`: Para procesar los datos.
  
  * `btnLimpiar`: Para borrar los campos.
  
  * `btnConfirmar`: Para mostrar el mensaje de venta exitosa.
  
  * `btnSalir`: Para cerrar el programa.

## Fase lógica: ¿Cómo lo hago?

1. **Definir Constantes:** Declara los precios de cada color (150.00, 165.50, 140.00, 175.25) y la tasa de transporte (0.05) al inicio de la clase.

2. **Capturar Datos:** En el botón Calcular, guarda el valor de los Textboxes de cantidades en variables enteras usando `Val()` para evitar errores con espacios vacíos.

3. **Calcular Subtotal:** Multiplica cada cantidad por su precio constante y suma los resultados.

4. **Calcular Transporte:** Multiplica el Subtotal obtenido por 0.05.

5. **Calcular Total:** Suma el Subtotal más el Transporte.

6. **Mostrar Resultados:** Asigna los valores calculados a las propiedades `.Text` de los Labels de resultado (usa `Format` para los decimales).

7. **Botón Limpiar:** Usa el método `.Clear()` en los Textboxes y restablece los Labels de resultados a "Q0.00".

8. **Botón Confirmar:** Usa un `If` para verificar que el total no sea 0; si es correcto, muestra un `MessageBox` con el nombre y el monto.

9. **Botón Salir:** Usa `MsgBox` para preguntar confirmación y `Me.Close()` para terminar.

## Codificar

Aquí tienes la explicación detallada del código, dividida por secciones lógicas para que entiendas qué hace cada parte.

### 1. Definición de variables

Esta sección define valores de las variables que al ser fijos se usaran constantes lo que significa que no cambiarán durante la ejecución del programa.

```visual-basic
' Constantes de precios 
Const PRECIO_ROJA As Double = 150.00 
Const PRECIO_AZUL As Double = 165.50 
Const PRECIO_VERDE As Double = 140.00 
Const PRECIO_NEGRA As Double = 175.25 
Const TASA_TRANSPORTE As Double = 0.05

```

* **`Const`**: Define un valor constante. Se usa para que, si el precio cambia en el futuro, solo tengas que modificarlo aquí y no en todo el código.

* **`Double`**: Es el tipo de dato para números con decimales (dinero).

* **`TASA_TRANSPORTE`**: Representa el 5% (0.05) que se cobrará por envío.

* * *

### 2. Botón Calcular (`btnCalcular`)

Este es el "cerebro" del programa. Realiza las matemáticas.

```visual-basic
Private Sub btnCalcular_Click(...) Handles btnCalcular.Click 
' Declaración de variables 
Dim cantRojas, cantAzules, cantVerdes, cantNegras As Integer 
Dim subtotal, pagoTransporte, pagoFinal As Double 
' Captura de datos 
cantRojas = Val(txtRojas.Text) 
cantAzules = Val(txtAzules.Text) 
cantVerdes = Val(txtVerdes.Text) 
cantNegras = Val(txtNegras.Text)
```

VB.Net
**`Dim`**: Reserva espacio en memoria para las variables (cantidades enteras y montos decimales).

* **`Val(...)`**: Esta función es vital. Convierte el texto de la caja (`.Text`) en un número. Si la caja está vacía, `Val` devuelve un `0` en lugar de dar error, evitando que el programa se cierre inesperadamente.
  
  

```visual-basic
        ' Cálculo del Subtotal
        subtotal = (cantRojas * PRECIO_ROJA) + (cantAzules * PRECIO_AZUL) + ...
        ' Cálculo del Transporte y Total
        pagoTransporte = subtotal * TASA_TRANSPORTE
        pagoFinal = subtotal + pagoTransporte
            ' Mostrar resultados
            lblSubtotal.Text = "Q" & Format(subtotal, "0.00")
            lblTransporte.Text = "Q" & Format(pagoTransporte, "0.00")
            lblTotal.Text = "Q" & Format(pagoFinal, "0.00")
        End Sub
```

* **Lógica matemática**: Multiplica cantidad por precio y suma todo. Luego calcula el 5% extra.

* **`lbl...Text =`**: Envía el resultado a la etiqueta en pantalla para que el usuario lo vea.

* **`Format(..., "0.00")`**: Asegura que el número siempre muestre dos decimales (ejemplo: muestra `150.50` en lugar de `150.5`).

* * *

### 3. Botón Limpiar (`btnLimpiar`)

Sirve para borrar todo y dejar el formulario listo para un nuevo cliente.

```visual-basic
    Private Sub btnLimpiar_Click(...) Handles btnLimpiar.Click
        ' Limpiar Cajas de Texto
        txtNombre.Clear()
        txtNit.Clear()
        txtRojas.Text = "0"
        ...
        ' Limpiar Resultados
        lblSubtotal.Text = "Q0.00"
        ...
        ' Enfocar
        txtNombre.Focus()
    End Sub
```

* **`.Clear()`**: Borra totalmente el texto (útil para Nombre y NIT).

* **`.Text = "0"`**: En las cantidades, es mejor dejar un "0" visible que dejarlo vacío, para indicar al usuario que ahí van números.

* **`.Focus()`**: Pone el cursor (la barrita que parpadea) automáticamente en la casilla del Nombre, ahorrándole un clic al usuario.

* * *

### 4. Botón Confirmar (`btnConfirmar`)

Este método valida que haya una venta antes de confirmar.

```visual-basic
    Private Sub btnConfirmar_Click(...) Handles btnConfirmar.Click
        If lblTotal.Text = "Q0.00" Then
            MessageBox.Show("Primero debe realizar el cálculo...", ...)
        Else
            MessageBox.Show("Venta procesada con éxito... " & txtNombre.Text ...)
        End If
    End Sub
```

* **`If lblTotal.Text = "Q0.00"`**: Verifica si el total es cero. Si es cero, significa que no se ha calculado nada, por lo que muestra un error (`MessageBoxIcon.Warning`).

* **`Else`**: Si el total es mayor a cero, muestra el mensaje de éxito.

* **`vbCrLf`**: Es un código especial que significa "Salto de línea" (Enter), para que el mensaje no salga todo seguido en una sola línea.

* * *

### 5. Botón Salir (`btnSalir`)

Cierra la aplicación de forma segura preguntando primero.

```visual-basic
    Private Sub btnSalir_Click(...) Handles btnSalir.Click
        Dim respuesta As MsgBoxResult
        respuesta = MsgBox("¿Desea salir...?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Salir")
        If respuesta = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub
```

* **`MsgBoxStyle.YesNo`**: Muestra el cuadro de diálogo con botones de "Sí" y "No".

* **`If respuesta = MsgBoxResult.Yes`**: Solo si el usuario presiona "Sí", se ejecuta la siguiente línea.

* **`Me.Close()`**: Cierra el formulario actual y termina el programa.

### Resultado esperado

```visual-basic
Public Class Form2
    ' Constantes de precios
    Const PRECIO_ROJA As Double = 150.00
    Const PRECIO_AZUL As Double = 165.50
    Const PRECIO_VERDE As Double = 140.00
    Const PRECIO_NEGRA As Double = 175.25
    Const TASA_TRANSPORTE As Double = 0.05
    ' Método Botón Calcular
    Private Sub btnCalcular_Click(sender As Object, e As EventArgs) Handles btnCalcular.Click
        ' Declaración de variables
        Dim cantRojas, cantAzules, cantVerdes, cantNegras As Integer
        Dim subtotal, pagoTransporte, pagoFinal As Double
        ' Captura de datos (Val evita errores si el campo está vacío)
        cantRojas = Val(txtRojas.Text)
        cantAzules = Val(txtAzules.Text)
        cantVerdes = Val(txtVerdes.Text)
        cantNegras = Val(txtNegras.Text)

        ' Cálculo del Subtotal
        subtotal = (cantRojas * PRECIO_ROJA) +
                   (cantAzules * PRECIO_AZUL) +
                   (cantVerdes * PRECIO_VERDE) +
                   (cantNegras * PRECIO_NEGRA)

        ' Cálculo del Transporte y Total
        pagoTransporte = subtotal * TASA_TRANSPORTE
        pagoFinal = subtotal + pagoTransporte

        ' Mostrar resultados en las etiquetas (Labels)
        lblSubtotal.Text = "Q" & Format(subtotal, "0.00")
        lblTransporte.Text = "Q" & Format(pagoTransporte, "0.00")
        lblTotal.Text = "Q" & Format(pagoFinal, "0.00")
    End Sub

    ' Método Botón Limpiar (Inicializar ingresos)
    Private Sub btnLimpiar_Click(sender As Object, e As EventArgs) Handles btnLimpiar.Click
        ' Limpiar Cajas de Texto
        txtNombre.Clear()
        txtNit.Clear()
        txtRojas.Text = "0"
        txtAzules.Text = "0"
        txtVerdes.Text = "0"
        txtNegras.Text = "0"

        ' Limpiar Resultados
        lblSubtotal.Text = "Q0.00"
        lblTransporte.Text = "Q0.00"
        lblTotal.Text = "Q0.00"

        ' Enfocar en el primer campo
        txtNombre.Focus()
    End Sub

    ' Método Botón Confirmar (El método extra solicitado)
    Private Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
        If lblTotal.Text = "Q0.00" Then
            MessageBox.Show("Primero debe realizar el cálculo de la venta.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MessageBox.Show("Venta procesada con éxito para el cliente: " & txtNombre.Text & vbCrLf &
                            "Monto Final: " & lblTotal.Text, "Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    ' Método Botón Salir
    Private Sub btnSalir_Click(sender As Object, e As EventArgs) Handles btnSalir.Click
        Dim respuesta As MsgBoxResult
        respuesta = MsgBox("¿Desea salir del programa?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Salir")
        If respuesta = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub
End Class
```

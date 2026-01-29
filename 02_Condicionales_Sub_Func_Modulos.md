# Parte 1: Sentencias Condicionales

#### Ejercicio 1: Acceso al Sistema (If...Then...Else con Operador `And`)

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

#### Ejercicio 2: Validación de Rango (If...ElseIf con Operador `Or`)

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

#### Ejercicio 3: Estado de Interruptor (Variable Booleana + IIf)

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

#### Ejercicio 4: Categoría de Cliente (Select Case)

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

#### Ejercicio 5: Limpieza de Formulario (Subrutina)

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

#### Ejercicio 6: Calculadora de IVA (Función)

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

#### Ejercicio 7: Variables Globales y Funciones Compartidas (Module)

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



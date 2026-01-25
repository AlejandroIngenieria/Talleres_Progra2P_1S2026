# Ejercicio 1: Operaciones y Mensajes

Para este ejercicio, solo necesitas un **Button** en tu formulario. Haz doble clic en él y pega el siguiente código:

```VB.Net
        ' --- 1. DECLARACIÓN Y MENSAJES SIMPLES ---
        Dim numero1 As Integer = 10
        Dim numero2 As Integer = 2
        ' Mostramos cada variable en un mensaje básico
        MessageBox.Show(numero1)
        MessageBox.Show(numero2)

        ' --- 2. OPERACIONES MATEMÁTICAS ---
        ' Suma
        Dim suma As Integer = numero1 + numero2
        MessageBox.Show("La suma es: " & suma)

        ' Resta
        Dim resta As Integer = numero1 - numero2
        MessageBox.Show("La resta es: " & resta)

        ' Multiplicación
        Dim multiplicacion As Integer = numero1 * numero2
        MessageBox.Show("La multiplicación es: " & multiplicacion)

        ' División (Usamos Double porque el resultado puede tener decimales)
        Dim division As Double = numero1 / numero2
        MessageBox.Show("La división es: " & division)

        ' --- 3. STRINGS Y MESSAGEBOX COMPUESTOS ---
        Dim nombre As String = "Usuario"
        Dim mensaje As String = "Operaciones completadas con éxito"

        ' MessageBox con Título, Botones e Icono de Información
        MessageBox.Show(mensaje, "Sistema de " & nombre, MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' MessageBox con Icono de Pregunta y botones Sí/No
        MessageBox.Show("¿Deseas realizar otra operación?", "Consulta", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

```

* * *

### Resumen de lo aprendido en este ejercicio:

* **Variables Numéricas:** Usamos `Integer` para números enteros y `Double` para resultados que puedan dar decimales (como la división).

* **Concatenación:** Usamos el símbolo **`&`** para unir texto con valores de variables en los mensajes.

* **Operadores:**
  
  * Suma: `+`
  
  * Resta: `-`
  
  * Multiplicación: `*`
  
  * División: `/`

* **MessageBox:** Aprendiste a pasar de un mensaje simple a uno con **Título**, **Botones** (`MessageBoxButtons`) e **Iconos** (`MessageBoxIcon`).

# Ejercicio 2: Ficha de personaje

Este ejercicio crea una "Ficha de Personaje". Usaremos el **TextBox** para el nombre y definiremos el resto de los datos directamente en el código para evitar funciones de conversión.

### 1. Elementos necesarios

Coloca estos controles en tu formulario:

| **Control** | **Nombre (Name)** | **Propósito**                     |
| ----------- | ----------------- | --------------------------------- |
| **TextBox** | `txtNombre`       | Escribir el nombre del personaje. |
| **Button**  | `btnCrear`        | Ejecutar el código.               |

* * *

### 2. El Código

Haz doble clic en el botón y pega este código:

```VB.Net
        ' --- DECLARACIÓN DE VARIABLES ---

        ' String: Obtiene el texto directamente del TextBox
        Dim nombre As String = txtNombre.Text

        ' Integer: Números enteros (ej. Puntos de Vida)
        Dim vida As Integer = 100

        ' Double: Números con decimales (ej. Velocidad)
        Dim velocidad As Double = 10.5

        ' Decimal: Gran precisión (ej. Monedas de oro)
        Dim oro As Decimal = 500.75

        ' Byte: Números pequeños de 0 a 255 (ej. Nivel)
        Dim nivel As Byte = 1

        ' --- COMENTARIO: CONSTRUCCIÓN DEL MENSAJE ---
        ' Unimos los datos usando el símbolo &

        Dim ficha As String
        ficha = "Personaje: " & nombre & vbCrLf &
                "Nivel: " & nivel & vbCrLf &
                "Vida: " & vida & vbCrLf &
                "Velocidad: " & velocidad & vbCrLf &
                "Oro: " & oro

        ' Mostrar resultado
        MessageBox.Show(ficha, "Ficha de Personaje")
```

### Resumen de tipos utilizados

* **`String`**: Para palabras o frases.

* **`Integer`**: Para números sin decimales.

* **`Double`**: Para decimales comunes.

* **`Decimal`**: Para dinero o valores que requieren exactitud total.

* **`Byte`**: Para números muy pequeños (ahorra memoria).

# Ejercicio 3: Simulador de Validación de Seguridad

Para este ejercicio, usaremos variables con valores fijos (hardcoded) para enfocarnos puramente en la lógica de los `If`. Solo necesitas un **Button**.

* * *

### Código del Ejercicio

Copia este código dentro del evento `Click` de tu botón:

```VB.Net
        ' --- 1. DECLARACIÓN DE DATOS ---
        Dim usuario As String = "Admin"
        Dim nivelAcceso As Integer = 5
        Dim saldo As Double = 150.75
        Dim intentos As Byte = 1
        Dim cuentaActiva As Boolean = True
        ' --- 2. SERIE DE "IF" SIMPLES ---
        ' Verificación de Nombre
        If usuario = "Admin" Then
            MessageBox.Show("Bienvenido, Administrador.", "Usuario", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' Verificación de Nivel (Número entero)
        If nivelAcceso >= 3 Then
            MessageBox.Show("Nivel de acceso elevado detectado.", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

        ' Verificación de Saldo (Double)
        If saldo > 100 Then
            MessageBox.Show("Cuenta con fondos suficientes.", "Finanzas")
        End If

        ' Verificación de Intentos (Byte)
        If intentos = 1 Then
            MessageBox.Show("Es tu primer intento del día.", "Logs")
        End If

        ' Verificación de Estado (Boolean)
        If cuentaActiva = True Then
            MessageBox.Show("La cuenta está vigente.", "Estado", MessageBoxButtons.OK, MessageBoxIcon.Check)
        End If

        ' --- 3. IF CON STRING Y MESSAGEBOX COMPUESTO ---
        ' Verificamos si el usuario es Admin para preguntar si quiere cerrar sesión
        If usuario = "Admin" Then
            MessageBox.Show("¿Desea cerrar el sistema de seguridad?", "Cierre", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        End If
```

* * *

### Puntos clave de la lógica `If`

* **Independencia:** Al usar solo `If` (sin `Else`), el programa evalúa **cada condición una por una**. Si todas son verdaderas, se mostrarán todos los mensajes.

* **Operadores de comparación:**
  
  * `=`: Igual a.
  
  * `>`: Mayor que.
  
  * `>=`: Mayor o igual que.

* **Booleano:** Para los valores `True/False`, no es estrictamente necesario poner `= True`, pero en nivel principiante ayuda a leer mejor el código.

* * *

### Notas de apoyo

> **Recordatorio:** En Visual Basic, los `If` siempre deben terminar con la instrucción `End If` si el código a ejecutar está en una línea diferente a la condición.

# Ejercicios con If simple

## Ejercicio 1: Edad mínima para votar

Solicita al usuario su edad. Si tiene 18 años o más, mostrar el mensaje "Puedes votar".

Controles sugeridos:

* TextBox para ingresar la edad
* Button para verificar
* Label para mostrar el resultado

Sentencias condicionales sugeridas:

* If edad >= 18 Then
  
  

## Ejercicio 2: Número positivo o negativo

Solicita un número al usuario e indica si es positivo.

Controles sugeridos:

* TextBox para ingresar el número
* Button para verificar
* Label para mostrar si es positivo

Sentencias condicionales sugeridas:

* If numero > 0 Then



## Ejercicio 3: Verificación de contraseña

Solicita una contraseña. Si el usuario escribe "1234", mostrar "Acceso concedido".

Controles sugeridos:

* TextBox con propiedad PasswordChar activada
* Button para validar
* Label para mostrar el mensaje

Sentencias condicionales sugeridas:

* If contrasena = "1234" Then



## Ejercicio 4: Verificación de número par

Solicita un número. Si es par, mostrar "Es par".

Controles sugeridos:

* TextBox para ingresar el número
* Button para comprobar
* Label para mostrar el resultado

Sentencias condicionales sugeridas:

* If numero Mod 2 = 0 Then



## Ejercicio 5: Calificación aprobatoria

Pide al usuario su calificación. Si es mayor a 60, mostrar "Aprobado".

Controles sugeridos:

* TextBox para ingresar la calificación
* Button para evaluar
* Label para mostrar el resultado

Sentencias condicionales sugeridas:

* If nota > 60 Then





























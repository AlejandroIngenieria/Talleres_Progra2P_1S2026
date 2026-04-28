# 🏥 Sistema de Gestión de Citas Médicas
### VB.NET + Windows Forms + MySQL

---

## 📋 Descripción del Problema

Una clínica médica te contrata para crear un sistema que gestione citas de pacientes. El sistema debe:

- Leer desde la base de datos los **médicos disponibles**, sus **especialidades** y su **tarifa por consulta**.
- El usuario ingresa los datos de la cita (paciente, médico, días de tratamiento, cantidad de sesiones).
- Con **lógica de programación** (vectores, ciclos, estadísticas) se calculan los cobros y se generan reportes.

---

## 🗄️ PARTE 1 — Base de Datos MySQL

### 1.1 Crear la base de datos y tabla

```sql
-- Crear base de datos
CREATE DATABASE clinica_db;
USE clinica_db;

-- Tabla de médicos (solo lectura desde VB)
CREATE TABLE medicos (
    id_medico   INT          PRIMARY KEY AUTO_INCREMENT,
    nombre      VARCHAR(100) NOT NULL,
    especialidad VARCHAR(50) NOT NULL,
    tarifa      DECIMAL(10,2) NOT NULL
);

-- Insertar datos de prueba
INSERT INTO medicos (nombre, especialidad, tarifa) VALUES
('Dr. Carlos López',    'General',      Q250.00 -- usar 250.00),
('Dra. María Pérez',   'Pediatría',    350.00),
('Dr. Juan García',    'Cardiología',  500.00);
```

> ⚠️ En MySQL no uses `Q` para Quetzales, escribe solo el número: `250.00`

```sql
INSERT INTO medicos (nombre, especialidad, tarifa) VALUES
('Dr. Carlos López',    'General',     250.00),
('Dra. María Pérez',   'Pediatría',   350.00),
('Dr. Juan García',    'Cardiología', 500.00);
```

---

## 💻 PARTE 2 — Proyecto Visual Basic .NET (Windows Forms)

### 2.1 Configuración del proyecto

| Parámetro | Valor |
|---|---|
| Tipo de proyecto | Windows Forms App (.NET Framework) |
| Nombre | `CLINICA` |
| Framework | .NET Framework 4.7.2 |
| Paquete NuGet requerido | `MySql.Data` |

---

### 2.2 Módulo de variables globales (`ModuloDatos.vb`)

```vb
Module ModuloDatos

    ' ── Cadena de conexión ──────────────────────────────────────────
    Public ConexionStr As String =
        "Server=localhost;Database=clinica_db;Uid=root;Pwd=TuPassword;"

    ' ── Constante: número de citas a registrar ──────────────────────
    Public Const MAX_CITAS As Integer = 7

    ' ── Vectores (arreglos paralelos) ───────────────────────────────
    Public idCliente(MAX_CITAS - 1)      As String
    Public nombreMedico(MAX_CITAS - 1)   As String
    Public especialidad(MAX_CITAS - 1)   As String
    Public diasTratamiento(MAX_CITAS - 1) As Integer
    Public numSesiones(MAX_CITAS - 1)    As Integer
    Public tarifaDia(MAX_CITAS - 1)      As Decimal
    Public totalPagar(MAX_CITAS - 1)     As Decimal

    ' ── Índice de registro actual ────────────────────────────────────
    Public indiceActual As Integer = 0

    ' ── Función: calcular total ──────────────────────────────────────
    Public Function CalcularTotal(dias As Integer,
                                  sesiones As Integer,
                                  tarifa As Decimal) As Decimal
        Return Math.Round(dias * sesiones * tarifa, 2)
    End Function

    ' ── Función: limpiar todos los vectores ─────────────────────────
    Public Sub LimpiarVectores()
        For i As Integer = 0 To MAX_CITAS - 1
            idCliente(i)       = ""
            nombreMedico(i)    = ""
            especialidad(i)    = ""
            diasTratamiento(i) = 0
            numSesiones(i)     = 0
            tarifaDia(i)       = 0
            totalPagar(i)      = 0
        Next
        indiceActual = 0
    End Sub

End Module
```

---

### 2.3 Formulario principal (`Form1.vb`)

#### Controles necesarios en el diseñador:

| Control | Nombre (Name) | Descripción |
|---|---|---|
| `TextBox` | `txtIdCliente` | ID del paciente |
| `ComboBox` | `cboMedico` | Lista médicos (cargada desde DB) |
| `NumericUpDown` | `nudDias` | Días de tratamiento |
| `NumericUpDown` | `nudSesiones` | Número de sesiones |
| `Label` | `lblTarifa` | Muestra tarifa del médico seleccionado |
| `DataGridView` | `dgvCitas` | Muestra el contenido de los vectores |
| `Label` | `lblEstadisticas` | Muestra resultados estadísticos |
| `MenuStrip` | `menuPrincipal` | Menú colgante |

#### Ítems del `MenuStrip`:

```
Opciones
 ├── Aceptar
 ├── Mostrar Vectores
 ├── Limpiar Vectores
 └── Estadísticas
```

---

#### Código del formulario (`Form1.vb`)

```vb
Imports MySql.Data.MySqlClient
Imports System.Data

Public Class Form1

    ' Guarda la tarifa del médico seleccionado en el ComboBox
    Private tarifaSeleccionada As Decimal = 0

    ' ════════════════════════════════════════════════════════════════
    '  CARGA DEL FORMULARIO — obtiene médicos desde MySQL
    ' ════════════════════════════════════════════════════════════════
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CargarMedicosDesdeDB()
        ConfigurarDataGridView()
    End Sub

    ' ────────────────────────────────────────────────────────────────
    '  Conexión y consulta a MySQL (solo lectura)
    ' ────────────────────────────────────────────────────────────────
    Private Sub CargarMedicosDesdeDB()
        Try
            Using conn As New MySqlConnection(ConexionStr)
                conn.Open()
                Dim query As String = "SELECT id_medico, nombre, especialidad, tarifa FROM medicos"
                Dim cmd As New MySqlCommand(query, conn)
                Dim reader As MySqlDataReader = cmd.ExecuteReader()

                cboMedico.Items.Clear()

                While reader.Read()
                    ' Guardamos los datos en un objeto anónimo dentro del ComboBox
                    Dim item As New With {
                        .Id          = reader("id_medico").ToString(),
                        .Nombre      = reader("nombre").ToString(),
                        .Especialidad = reader("especialidad").ToString(),
                        .Tarifa      = Convert.ToDecimal(reader("tarifa")),
                        .Display     = $"{reader("nombre")} — {reader("especialidad")}"
                    }
                    cboMedico.Items.Add(item)
                End While

                reader.Close()
            End Using

            ' Mostrar texto en ComboBox usando la propiedad Display
            cboMedico.DisplayMember = "Display"

        Catch ex As Exception
            MessageBox.Show("Error al conectar con la base de datos: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ────────────────────────────────────────────────────────────────
    '  Cuando el usuario elige un médico, muestra su tarifa
    ' ────────────────────────────────────────────────────────────────
    Private Sub cboMedico_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles cboMedico.SelectedIndexChanged

        If cboMedico.SelectedIndex >= 0 Then
            Dim item = cboMedico.SelectedItem
            tarifaSeleccionada = item.Tarifa
            lblTarifa.Text = $"Tarifa por sesión/día: Q{tarifaSeleccionada:F2}"
        End If
    End Sub

    ' ════════════════════════════════════════════════════════════════
    '  MENÚ → ACEPTAR  (guardar en vectores)
    ' ════════════════════════════════════════════════════════════════
    Private Sub mnuAceptar_Click(sender As Object, e As EventArgs) Handles mnuAceptar.Click

        If indiceActual >= MAX_CITAS Then
            MessageBox.Show("Ya se registraron las 7 citas permitidas.", "Aviso")
            Return
        End If

        If cboMedico.SelectedIndex < 0 Then
            MessageBox.Show("Seleccione un médico.", "Aviso")
            Return
        End If

        ' Leer datos de la interfaz
        Dim medico     = cboMedico.SelectedItem
        Dim dias       As Integer = CInt(nudDias.Value)
        Dim sesiones   As Integer = CInt(nudSesiones.Value)

        ' Guardar en vectores
        idCliente(indiceActual)       = txtIdCliente.Text.Trim()
        nombreMedico(indiceActual)    = medico.Nombre
        especialidad(indiceActual)    = medico.Especialidad
        diasTratamiento(indiceActual) = dias
        numSesiones(indiceActual)     = sesiones
        tarifaDia(indiceActual)       = tarifaSeleccionada

        ' Calcular total usando la función del módulo
        totalPagar(indiceActual) = CalcularTotal(dias, sesiones, tarifaSeleccionada)

        indiceActual += 1

        MessageBox.Show($"Cita #{indiceActual} registrada. Total: Q{totalPagar(indiceActual - 1):F2}",
                        "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' Limpiar controles para la siguiente entrada
        txtIdCliente.Clear()
        cboMedico.SelectedIndex = -1
        nudDias.Value    = 1
        nudSesiones.Value = 1
        lblTarifa.Text   = "Tarifa: —"
    End Sub

    ' ════════════════════════════════════════════════════════════════
    '  MENÚ → MOSTRAR VECTORES
    ' ════════════════════════════════════════════════════════════════
    Private Sub mnuMostrar_Click(sender As Object, e As EventArgs) Handles mnuMostrar.Click
        MostrarVectores()
    End Sub

    Private Sub MostrarVectores()
        dgvCitas.Rows.Clear()

        ' Ciclo For para recorrer los vectores
        For i As Integer = 0 To indiceActual - 1
            dgvCitas.Rows.Add(
                i + 1,
                idCliente(i),
                nombreMedico(i),
                especialidad(i),
                diasTratamiento(i),
                numSesiones(i),
                $"Q{tarifaDia(i):F2}",
                $"Q{totalPagar(i):F2}"
            )
        Next
    End Sub

    ' ════════════════════════════════════════════════════════════════
    '  MENÚ → LIMPIAR VECTORES
    ' ════════════════════════════════════════════════════════════════
    Private Sub mnuLimpiar_Click(sender As Object, e As EventArgs) Handles mnuLimpiar.Click
        LimpiarVectores()          ' Función del módulo (usa For internamente)
        dgvCitas.Rows.Clear()
        lblEstadisticas.Text = ""
        MessageBox.Show("Vectores limpiados. Listo para nuevas citas.", "Limpiar")
    End Sub

    ' ════════════════════════════════════════════════════════════════
    '  MENÚ → ESTADÍSTICAS  (usa ciclo While)
    ' ════════════════════════════════════════════════════════════════
    Private Sub mnuEstadisticas_Click(sender As Object, e As EventArgs) Handles mnuEstadisticas.Click

        If indiceActual = 0 Then
            MessageBox.Show("No hay datos registrados.", "Aviso")
            Return
        End If

        ' ── Variables acumuladoras ───────────────────────────────────
        Dim totalSesionesGeneral    As Integer = 0
        Dim totalDineroCardiologia  As Decimal = 0
        Dim citasPediatria          As Integer = 0

        ' ── Ciclo While ─────────────────────────────────────────────
        Dim i As Integer = 0
        While i < indiceActual

            ' Estadística 1: total de sesiones de especialidad General
            If especialidad(i) = "General" Then
                totalSesionesGeneral += numSesiones(i)
            End If

            ' Estadística 2: total en dinero de Cardiología
            If especialidad(i) = "Cardiología" Then
                totalDineroCardiologia += totalPagar(i)
            End If

            ' Estadística 3: cuántas citas de Pediatría
            If especialidad(i) = "Pediatría" Then
                citasPediatria += 1
            End If

            i += 1
        End While

        ' ── Mostrar resultados ───────────────────────────────────────
        Dim resultado As String =
            $"📊 ESTADÍSTICAS:{Environment.NewLine}" &
            $"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Environment.NewLine}" &
            $"1. Sesiones en Medicina General : {totalSesionesGeneral}{Environment.NewLine}" &
            $"2. Total dinero en Cardiología  : Q{totalDineroCardiologia:F2}{Environment.NewLine}" &
            $"3. Citas de Pediatría           : {citasPediatria}"

        lblEstadisticas.Text = resultado
    End Sub

    ' ════════════════════════════════════════════════════════════════
    '  Configurar columnas del DataGridView
    ' ════════════════════════════════════════════════════════════════
    Private Sub ConfigurarDataGridView()
        dgvCitas.Columns.Clear()
        dgvCitas.Columns.Add("col1", "#")
        dgvCitas.Columns.Add("col2", "ID Paciente")
        dgvCitas.Columns.Add("col3", "Médico")
        dgvCitas.Columns.Add("col4", "Especialidad")
        dgvCitas.Columns.Add("col5", "Días")
        dgvCitas.Columns.Add("col6", "Sesiones")
        dgvCitas.Columns.Add("col7", "Tarifa/Ses.")
        dgvCitas.Columns.Add("col8", "Total Q")
        dgvCitas.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

End Class
```

---

## 📐 PARTE 3 — Diseño de la interfaz (referencia)

```
┌─────────────────────────────────────────────────────────┐
│  Opciones ▼                                             │
│  ┌──────────┬──────────────┬──────────────┬──────────┐  │
│  │ Aceptar  │ Mostrar Vec. │ Limpiar Vec. │ Estadíst │  │
│  └──────────┴──────────────┴──────────────┴──────────┘  │
│                                                         │
│  ID Paciente: [______________]                          │
│  Médico:      [ComboBox ▼   ]  Tarifa: Q 000.00        │
│  Días:        [NumericUpDown]                           │
│  Sesiones:    [NumericUpDown]                           │
│                                                         │
│  ┌─────────────────────────────────────────────────┐   │
│  │  #  │ ID │ Médico │ Espec. │ Días │ Ses │ Total │   │
│  │  1  │    │        │        │      │     │       │   │
│  └─────────────────────────────────────────────────┘   │
│                                                         │
│  [lblEstadisticas — resultados estadísticos aquí]       │
└─────────────────────────────────────────────────────────┘
```

---

## 🔢 Fórmula de cálculo

```
Total a pagar = Días × Sesiones × Tarifa por sesión
```

> La función `CalcularTotal` en el módulo aplica `Math.Round(..., 2)` para redondeo a 2 decimales.
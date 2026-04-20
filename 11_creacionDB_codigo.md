Para crear la base de datos **InventoryPack** en MySQL siguiendo el diseño de tu archivo, sigue este orden lógico y técnico:

---

## 1. Estructura Inicial y Tipos de Datos

Antes de crear tablas, debes definir el contenedor. El comando `CREATE DATABASE` crea el espacio, y `USE` le indica a MySQL que todas las acciones siguientes se aplicarán a esa base de datos.

### Tipos de Datos Comunes en SQL
* **INT**: Números enteros. Se usa para IDs.
* **VARCHAR(N)**: Texto de longitud variable hasta **N** caracteres. Ideal para nombres y contraseñas.
* **DECIMAL(M,D)**: Números con decimales exactos (M = total de dígitos, D = decimales). Es el estándar para dinero.
* **DATETIME**: Almacena fecha y hora.
* **TEXT**: Para descripciones largas que superan los 255 caracteres.

---

## 2. Creación de Tablas y Relaciones

El orden de creación es crítico: **primero crea las tablas "padre"** (las que no dependen de nadie) y **luego las tablas "hijo"** (las que tienen llaves foráneas).

### Definiciones Clave
* **Primary Key (PK)**: El identificador único de la tabla.
* **Foreign Key (FK)**: Una columna que referencia a la PK de otra tabla para crear el vínculo.
* **AUTO_INCREMENT**: MySQL genera el número del ID automáticamente.

### Guía de comandos SQL

```sql
-- 1. Crear y usar la base de datos
CREATE DATABASE InventoryPack;
USE InventoryPack;

-- 2. Tablas Independientes (Padres)
CREATE TABLE usuarios (
    id_usuario INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(50) UNIQUE NOT NULL, -- Restricción de unicidad
    password VARCHAR(255) NOT NULL,
    rol VARCHAR(20) NOT NULL, -- 'Administrador' o 'Empleado'
    nombre_completo VARCHAR(100)
);

CREATE TABLE productos (
    id_producto INT AUTO_INCREMENT PRIMARY KEY,
    nombre VARCHAR(100) NOT NULL,
    descripcion TEXT,
    precio DECIMAL(10,2) NOT NULL,
    stock INT NOT NULL -- Debe validarse antes de pedidos
);

-- 3. Tablas Dependientes (Hijos)
CREATE TABLE pedidos (
    id_pedido INT AUTO_INCREMENT PRIMARY KEY,
    id_usuario INT,
    fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP,
    total DECIMAL(10,2),
    estado VARCHAR(50),
    FOREIGN KEY (id_usuario) REFERENCES usuarios(id_usuario) -- Relación con usuarios
);

CREATE TABLE detalle_pedido (
    id_detalle INT AUTO_INCREMENT PRIMARY KEY,
    id_pedido INT,
    id_producto INT,
    cantidad INT NOT NULL,
    subtotal DECIMAL(10,2),
    FOREIGN KEY (id_pedido) REFERENCES pedidos(id_pedido),
    FOREIGN KEY (id_producto) REFERENCES productos(id_producto)
);

-- 4. Tabla de Capacidad Estricta
CREATE TABLE despacho_cuadricula (
    id_espacio INT PRIMARY KEY, -- Del 1 al 9 manualmente
    id_pedido INT UNIQUE, -- Solo un pedido por espacio
    estado_espacio VARCHAR(20) DEFAULT 'Libre',
    ultima_actualizacion DATETIME ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (id_pedido) REFERENCES pedidos(id_pedido)
);
```

---

## 3. Población Inicial (Data Seeding)

Según las reglas de tu proyecto, hay datos que deben existir desde el inicio:

1.  **Administrador por defecto**: El sistema requiere credenciales iniciales (`admin`/`admin`).
2.  **Cuadrícula de Despacho**: Debe tener exactamente 9 registros pre-insertados.

```sql
-- Insertar administrador inicial
INSERT INTO usuarios (username, password, rol, nombre_completo) 
VALUES ('admin', 'admin', 'Administrador', 'Administrador del Sistema');

-- Insertar los 9 espacios físicos del despacho
INSERT INTO despacho_cuadricula (id_espacio, estado_espacio) VALUES 
(1, 'Libre'), (2, 'Libre'), (3, 'Libre'), 
(4, 'Libre'), (5, 'Libre'), (6, 'Libre'), 
(7, 'Libre'), (8, 'Libre'), (9, 'Libre');
```

---

## Resumen de Reglas de Negocio a programar en VB.NET
* **Usuarios**: Tu código debe contar los registros en `usuarios` antes de un `INSERT`. Si el conteo es $\ge 4$, debe bloquear el registro.
* **Stock**: Antes de insertar en `detalle_pedido`, debes consultar si el `stock` en `productos` es suficiente.
* **Despacho**: Al crear un pedido, debes buscar un `id_espacio` cuyo `estado_espacio` sea 'Libre'.
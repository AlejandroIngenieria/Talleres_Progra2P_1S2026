Esta guía técnica aplica los conceptos de **SQL Avanzado** vistos en clase  directamente a tu base de datos `ZooManagement`. Utiliza estos comandos para manipular, consultar y optimizar la información de tu sistema de gestión.

# Script de la base de datos

```sql
CREATE DATABASE IF NOT EXISTS ZooManagement;
USE ZooManagement;

-- 1. Tabla de Personal (Equivalente a usuarios)
CREATE TABLE personal_zoo (
    id_personal INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(50) NOT NULL UNIQUE,
    password VARCHAR(255) NOT NULL,
    rol VARCHAR(20) NOT NULL, -- 'Admin' o 'Cuidador'
    nombre_completo VARCHAR(100) NOT NULL
);

-- 2. Tabla de Animales (Equivalente a productos)
CREATE TABLE animales (
    id_animal INT AUTO_INCREMENT PRIMARY KEY,
    nombre VARCHAR(100) NOT NULL,
    especie VARCHAR(100) NOT NULL,
    precio_mantenimiento DECIMAL(10,2) NOT NULL,
    salud_puntos INT NOT null, -- Validar estado antes de asignar
    image_url Text
);

-- 3. Tabla de Bitácora (Equivalente a pedidos)
CREATE TABLE bitacora_cuidados (
    id_bitacora INT AUTO_INCREMENT PRIMARY KEY,
    id_personal INT NOT NULL,
    fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP,
    costo_total_insumos DECIMAL(10,2) NOT NULL,
    tipo_actividad VARCHAR(50) NOT NULL, -- 'Alimentación', 'Medicina', etc.
    FOREIGN KEY (id_personal) REFERENCES personal_zoo(id_personal)
);

-- 4. Tabla de Detalle (Equivalente a detalle_pedido)
CREATE TABLE detalle_cuidados (
    id_detalle INT AUTO_INCREMENT PRIMARY KEY,
    id_bitacora INT NOT NULL,
    id_animal INT NOT NULL,
    cantidad_alimento INT NOT NULL,
    subtotal_insumos DECIMAL(10,2) NOT NULL,
    FOREIGN KEY (id_bitacora) REFERENCES bitacora_cuidados(id_bitacora),
    FOREIGN KEY (id_animal) REFERENCES animales(id_animal)
);

-- 5. Tabla de Recintos (Equivalente a despacho_cuadricula)
CREATE TABLE recintos_mapa (
    id_recinto INT PRIMARY KEY,
    id_animal INT UNIQUE NULL,
    estado_limpieza VARCHAR(20) NOT NULL, -- 'Limpio' / 'Sucio'
    ultima_inspeccion DATETIME,
    FOREIGN KEY (id_animal) REFERENCES animales(id_animal)
);
```

---

## 1. DML: Manipulación de Datos (CRUD)
El **DML** permite gestionar los registros de tus tablas.

### Inserción de Datos (`INSERT`)
Puedes insertar registros individuales o múltiples de forma eficiente. Recuerda que las cadenas de texto siempre van entre comillas simples (`''`).

```sql
-- Insertar un nuevo cuidador
INSERT INTO personal_zoo (username, password, rol, nombre_completo)
VALUES ('mlopez', 'pass123', 'Cuidador', 'Mario López');

-- Inserción múltiple de animales
INSERT INTO animales (nombre, especie, precio_mantenimiento, salud_puntos)
VALUES 
('Simba', 'León', 500.00, 100),
('Melman', 'Jirafa', 350.00, 95),
('Gloria', 'Hipopótamo', 450.00, 90);
```

### Actualización y Eliminación (`UPDATE` & `DELETE`)
**Advertencia Crítica:** Siempre utiliza la cláusula `WHERE`. Sin ella, modificarás o borrarás **todos** los registros de la tabla.

```sql
-- Actualizar la salud de un animal específico
UPDATE animales 
SET salud_puntos = 85 
WHERE id_animal = 1;

-- Eliminar un animal que ya no está en el zoo
DELETE FROM animales 
WHERE id_animal = 10;
```

---

## 2. DQL: Consultas y Filtrado
El **DQL** se utiliza para recuperar información específica mediante filtros y ordenamientos.

### Selección y Filtrado (`WHERE`)
Puedes filtrar por rangos, listas o condiciones lógicas.

```sql
-- Animales con mantenimiento entre 100 y 500
SELECT nombre, especie, precio_mantenimiento 
FROM animales 
WHERE precio_mantenimiento BETWEEN 100 AND 500;

-- Buscar cuidadores cuyo nombre empiece con 'A'
SELECT * FROM personal_zoo 
WHERE nombre_completo LIKE 'A%';
```

### Ordenamiento (`ORDER BY`)
Ordena tus resultados de forma ascendente (`ASC`) o descendente (`DESC`).

```sql
-- Lista de animales por salud (de mayor a menor)
SELECT nombre, salud_puntos 
FROM animales 
ORDER BY salud_puntos DESC;
```

---

## 3. Consultas Multitabla (JOINS)
Los **JOINS** permiten unir tablas aprovechando las **llaves foráneas** (FK) y sus cardinalidades.


| Tipo de JOIN | Función en ZooManagement | Ejemplo de Aplicación |
| :--- | :--- | :--- |
| **INNER JOIN** | Retorna registros con coincidencia en ambas tablas. | Ver qué animal está asignado a qué jaula. |
| **LEFT JOIN** | Todos los de la izquierda y coincidentes de la derecha. | Ver todos los animales, incluso los que no tienen jaula asignada. |
| **RIGHT JOIN** | Todos los de la derecha y coincidentes de la izquierda. | Ver todas las jaulas, incluso las que están vacías (NULL). |

**Ejemplo de Consulta Compleja:**
```sql
-- Ver animales y sus respectivos recintos
SELECT a.nombre AS Animal, r.id_recinto AS Jaula, r.estado_limpieza
FROM animales a
INNER JOIN recintos_mapa r ON a.id_animal = r.id_animal;
```

---

## 4. Normalización
La normalización asegura un diseño limpio y sin redundancias.


1.  **1FN (Atomicidad):** Cada celda de tu tabla `animales` tiene un solo valor (ej. un solo nombre, una sola especie). No guardes una lista de vacunas en una sola columna.
2.  **2FN (Dependencia Completa):** En `detalle_cuidados`, la cantidad de alimento depende de la combinación de la bitácora y el animal (PK Compuesta). 
3.  **3FN (Sin Dependencias Transitivas):** Tu tabla `personal_zoo` es correcta porque el `nombre_completo` depende del `id_personal`, no de otro atributo no clave.

---

> **Consejo Pro:** Antes de ejecutar un `UPDATE` o `DELETE` masivo, realiza un `SELECT` con el mismo `WHERE` para verificar exactamente qué filas se verán afectadas.
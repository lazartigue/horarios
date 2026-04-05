# Lazartigue — Sistema de Horarios V7.1

## Modelo General

- **Horario de atención**: Lunes a Viernes 09:00–20:00, Sábados 09:00–14:00.
- **Turno AM**: 09:00–18:30 (9.5 horas).
- **Turno PM**: 10:30–20:00 (9.5 horas).
- **Turno Sábado**: 09:00–14:00 (5 horas).
- **Domingos**: cerrado.
- **Equipo**: 21 personas distribuidas en 7 áreas.

---

## Áreas y Personas

| Área | Personas | Rol |
|---|---|---|
| **Coloristas** | Marcela Carrion, María Paz Hodges, Paola Mella, Daniel Olave | Daniel siempre PM (10:30). Las otras 3 rotan AM/PM. |
| **Aplicadoras** | Patricia Toledo, Cecilia Flores, Karen Leiva | Rotan AM/PM entre ellas. |
| **Corte / Peinado** | Lorena Cecilia Pacheco, Ingrid Higuera | Rotan AM/PM entre ellas. |
| **Ayudantes** | Isabel San Martín, Tamara Bustos, Uberlinda Urra, Carol Silva | Rotan AM/PM. 1 PM (Lun-Mié), 2 PM (Jue-Vie). |
| **Lavapelo y Secado** | Eugenia Uribe, Carolina Madariaga, Alma Arévalo, Francis | Rotan AM/PM. 2 PM por día. |
| **Recepción** | Antonio San Martín, Francisca Brocco, Valeria Otegui, Raquel Hasson | Antonio siempre AM. Las otras 3 rotan 1 PM por día. |
| **Back Office** | Pamela Hernández | Siempre PM (10:30). |

---

## Turnos Lunes a Viernes

Cada área tiene una rotación de turnos AM/PM que se repite en un **ciclo de 3 semanas**. Esto define quién entra a las 09:00 (AM) y quién a las 10:30 (PM) cada día de Martes a Viernes.

**Lunes, Martes y Miércoles** tienen una asignación dinámica de PM: el sistema determina quién hace PM esos días considerando quién tiene libre esa semana, para asegurar que siempre haya suficiente personal en apertura (09:00).

**Jueves y Viernes** son días peak y siguen la rotación fija de turnos, donde hay máxima dotación en la mañana.

---

## Sábados — Equipo Reducido con Rotación Individual

En lugar de que un grupo completo trabaje cada sábado, se arma un **equipo mínimo** seleccionando **1 persona por área** mediante rotación:

| Área | Quiénes rotan | Frecuencia individual |
|---|---|---|
| Coloristas | Marcela → María Paz → Paola → Daniel | Cada 4 semanas |
| Aplicadoras | Patricia → Cecilia → Karen | Cada 3 semanas |
| Ayudantes | Isabel → Tamara → Uberlinda → Carol | Cada 4 semanas |
| Lavapelo | Eugenia → Carolina → Alma → Francis | Cada 4 semanas |
| Recepción | Francisca → Valeria → Raquel | Cada 3 semanas |
| Back Office | Pamela | Cada 2 semanas (alterno) |
| Corte | Lorena | Cada 3 semanas (no todos) |

**Resultado**: cada sábado trabajan entre 5 y 7 personas (dependiendo de si le toca a Pamela y/o Lorena esa semana).

**Excluidos de sábados**: Antonio San Martín e Ingrid Higuera nunca trabajan sábado.

---

## Día Libre Post-Sábado

**Solo la persona que trabajó el sábado recibe un día libre** la semana siguiente (Lunes, Martes o Miércoles). El día se asigna según el área:

- Cada área tiene un día base (por ejemplo: coloristas parten en Lunes, aplicadoras en Martes, recepción en Miércoles).
- Ese día base **rota cada semana**, de modo que no siempre cae en el mismo día.
- Así se distribuyen los libres de forma pareja: aproximadamente 2–3 personas libres por día, en vez de 7–8 como sería con un grupo completo.

**Quienes no trabajaron el sábado no reciben libre entre semana.** Esto es la principal ventaja del sistema: menos ausencias de Lunes a Viernes, especialmente en Jueves y Viernes (días peak).

---

## Restricciones Consideradas

### Dotación mínima en apertura (09:00)
- **Coloristas**: mínimo 2 personas AM todos los días.
- **Aplicadoras**: mínimo 1 aplicadora por cada 2 coloristas en AM.
- **Ayudantes**: mínimo 2 presentes (idealmente 1 por cada colorista en AM).
- **Lavapelo**: mínimo 2 presentes.
- **Recepción**: mínimo 3 en apertura (Antonio + 2 de las rotativas).

Si asignar un turno PM a alguien dejaría su área por debajo del mínimo de apertura, el sistema no le asigna PM ese día.

### Dotación mínima en cierre (hasta 20:00)
- Siempre al menos 2 coloristas PM (Daniel fijo + 1 rotativa).
- Al menos 1 aplicadora PM, 1 corte PM.
- **Ayudantes PM**: mínimo 1 (Lun-Mié), **mínimo 2 (Jue-Vie)** por ser días peak.
- Al menos 2 lavapelo PM.
- Al menos 1 recepción PM.

### Personas con horario especial
- **Antonio San Martín**: siempre AM, sin sábados, sin libre.
- **Ingrid Higuera**: rota AM/PM, sin sábados, sin libre.
- **Daniel Olave**: siempre PM entre semana. Participa en rotación de sábados (09:00–14:00).
- **Pamela Hernández**: siempre PM entre semana. Sábado cada 2 semanas.
- **Francisca Brocco**: contrato Art. 22, rota turnos normalmente.

### Equidad
- La rotación de sábados es predecible: cada persona sabe con anticipación cuándo le toca.
- Los días libres rotan semanalmente para que no siempre caigan el mismo día de la semana.
- Los turnos PM rotan en ciclo de 3 semanas para que nadie cierre todos los días.

---

## Resumen de Beneficios

1. **Menos personas en sábado**: 5–7 en vez de 7–8. Menor costo operativo.
2. **Más cobertura Lunes a Viernes**: solo 2–3 libres por día en vez de ~8. Especialmente importante en días peak (Jueves/Viernes).
3. **Equidad en sábados**: cada persona trabaja según el tamaño de su pool (cada 3–4 semanas), no cada 3 semanas fijo para todos.
4. **Predecibilidad**: toda la rotación es calculable. Cualquier persona puede saber su calendario con meses de anticipación.

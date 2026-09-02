# AGRO — Dashboard Streamlit para construir la nueva malla

## Objetivo

Aplicación interactiva para:
- construir la propuesta de malla por semestres;
- añadir materias desde el catálogo internacional normalizado o crear una materia EPN propia;
- mover/eliminar materias en un tablero tipo Kanban;
- analizar cada semestre con radar + deltas frente al baseline EPN;
- medir exposición práctica, digitalización, prevalencia benchmark, emergentes y profundidad;
- comparar cada semestre con otra universidad usando **avance curricular relativo (%)**, no número nominal de semestre;
- mostrar solo familias/contenidos diferenciales ("ganamos / nos falta") según los datos disponibles;
- mostrar estado oficial EUR-ACE por separado de la contribución potencial de la propuesta;
- exportar la malla propuesta a Excel.

## 1. Estructura de carpetas

Coloque el script y una carpeta `data` así:

```text
proyecto/
├─ agro_curriculum_dashboard.py
└─ data/
   ├─ Base_Maestra_AGRO_v1_2_CAPAS_EVIDENCIA_NORMALIZADAS_AUDITADAS.xlsx
   ├─ EPN_AGRO_Dataset_Comparativo_Integrable_v1_0.xlsx
   ├─ Matriz_Comparativa_AGRO_v1_0.xlsx
   ├─ Matriz_Brechas_EPN_Agro_vs_Benchmark_v1_0.xlsx
   └─ CAEE_AGRO_v1_0_NORMALIZADA_AUDITADA.xlsx
```

## 2. Instalar

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate

pip install -r requirements.txt
```

## 3. Ejecutar

```bash
streamlit run agro_curriculum_dashboard.py
```

## 4. Modelo de datos mínimo

La aplicación **no carga todo el proyecto en pantalla**. Usa:

- Base Maestra:
  - `09_Asignaturas`
  - `11_Contenidos`
  - `33_Diccionario_Familias_Norm`
- EPN:
  - `04_Asignaturas_60`
  - `05_Carga_52`
  - `08_Prerrequisitos`
- Benchmark:
  - `02_Universos`
  - `04_Presencia_Curriculos`
  - `06_Posicion_Core`
  - `07_Profundidad_Core`
  - `08_Resumen_Areas`
- Brechas:
  - `02_Cobertura_Familias`
  - `08_Senales_Stakeholders`
- CAEE:
  - `02_Taxonomia_Temas`
  - `03_EURACE_Criterios`
  - `14_Cobertura_Temas`

## 5. Reglas metodológicas implementadas

1. EPN no entra en denominadores del benchmark.
2. `NR` nunca se transforma en cero.
3. Comparación temporal internacional por avance curricular relativo.
4. Los créditos de una universidad extranjera no se convierten automáticamente a créditos EPN.
5. El estado oficial EUR-ACE permanece inmutable.
6. Los ejes ejecutivos y el puente AGRO-NORM → CAEE-NORM son **capas derivadas de visualización** y no reescriben los datos auditados.
7. Si una universidad no tiene microcontenidos recuperados, el dashboard informa "sin evidencia disponible"; no declara ausencia curricular.

## 6. Siguiente mejora recomendada

Cuando la Junta estabilice una primera propuesta:
- incorporar validador de prerrequisitos;
- añadir control automático de 15 créditos por semestre / 135 totales;
- incorporar escenarios A/B/C y comparación Pareto;
- guardar versiones de propuesta con fecha/autor;
- añadir modo "Junta" (presentación) y modo "Analista" (detalle).

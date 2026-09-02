# AGRO Dashboard v0.3

## Cambios principales

- Búsqueda fuzzy con RapidFuzz y fallback interno.
- Panel desplegable de detalle para la materia seleccionada.
- Créditos EPN únicamente en números enteros.
- Malla Kanban compacta, con hasta 5 semestres por fila en escritorio.
- Tarjetas coloreadas por tipología curricular ejecutiva derivada de AGRO-NORM.
- Leyenda visual de tipos.
- Nueva pestaña `Análisis por tipos` con radar, barras y diferencias frente a un referente internacional.
- Se mantiene comparación internacional por avance curricular relativo.
- Se mantiene la regla NR != 0 y EUR-ACE oficial separado de contribución potencial.

## Tipos visuales

- Básicas — amarillo
- Ciencias agro/biológicas — verde claro
- Ingeniería y procesos — azul
- Agroalimentos y calidad — verde
- Digital — violeta
- Sostenibilidad — turquesa
- Gestión/administración — naranja
- Social/profesional — rosa
- Práctica/integración — azul grisáceo
- Optativas — gris

La clasificación visual es una capa derivada para el dashboard y no sustituye la taxonomía AGRO-NORM original.

## Instalación

```bash
pip install -r requirements.txt
streamlit run agro_curriculum_dashboard.py
```

Mantenga los cinco Excel aprobados dentro de la carpeta `data/` con los nombres usados en versiones anteriores.


## Cambios v0.4

- Un solo buscador fuzzy con `streamlit-searchbox`.
- Busca sobre todo el catálogo habilitado por nombre de materia, universidad, dataset, nombre/familia normalizada, área y código.
- `REVISAR` está incluido por defecto; puede ocultarse con el checkbox explicado mediante tooltip.
- La universidad de origen se recupera desde `04_Curriculo_Institucion` y se agrega sin duplicar materias de programas conjuntos.
- Se incorporan tooltips en controles y métricas principales.
- Se añaden paneles `Cómo interpretar...` para radar/barras, tipos, comparación internacional y CAEE/EUR-ACE.
- Los microcontenidos diferenciales en comparación internacional son desplegables.

Después de reemplazar esta versión, ejecute nuevamente:

```bash
pip install -r requirements.txt
streamlit run agro_curriculum_dashboard.py
```

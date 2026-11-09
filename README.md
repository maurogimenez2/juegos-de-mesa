# ¿A qué jugamos?

Lista de juegos de mesa con filtros por cantidad de jugadores, duración, complejidad, azar, modo de juego, género y precio.

## Cómo se actualiza

La página se genera a partir del Excel del grupo:

```
python build.py "ruta/al/Juegos de mesa.xlsx"
```

Eso regenera `index.html` y después alcanza con subirlo de nuevo al repo.

- Las columnas del Excel tienen que mantener sus nombres (Juego, Precio, Género, etc.).
- La imagen de cada juego es la miniatura del video de YouTube de la columna "Link reseña". Si no hay link, se muestran las iniciales.
- Si aparece un género nuevo, el script avisa y lo pone en "Otros". Para asignarle familia se agrega en `FAMILIAS` dentro de `build.py`.
- Las correcciones de nombres (por ejemplo "Carcasone" → "Carcassonne") están en `NOMBRES` dentro de `build.py`.
- Necesita `openpyxl` (`pip install openpyxl`).

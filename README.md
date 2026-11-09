# ¿A qué jugamos?

Lista de juegos de mesa con filtros por cantidad de jugadores, duración, complejidad, azar, modo de juego, género y precio.

## Cómo se actualiza

La página se genera a partir del Excel del grupo:

```
python build.py "ruta/al/Juegos de mesa.xlsx"
```

En Windows, si `python` no está en el PATH: `"C:\Program Files\Python313\python.exe" build.py "..."`.

Eso regenera `index.html`. Para publicarlo, en la web de GitHub: **Add file → Upload files**, arrastrar `index.html` (y `build.py` / `template.html` si cambiaron) y **Commit changes**. A veces el primer clic en "Commit changes" no hace nada: esperar a que aparezca "Processing your files…" o volver a hacer clic.

- Las columnas del Excel tienen que mantener sus nombres (Juego, Precio, Género, etc.).
- La imagen de cada juego es la miniatura del video de YouTube de la columna "Link reseña". Si no hay link, se muestran las iniciales.
- Si aparece un género nuevo, el script avisa y lo pone en "Otros". Para asignarle familia se agrega en `FAMILIAS` dentro de `build.py`.
- Si aparece una modalidad nueva, el chip de filtro se crea solo. El script avisa para que se agregue en `MODALIDAD` si hay que agruparla con otra.
- Las correcciones de nombres (por ejemplo "Carcasone" → "Carcassonne") están en `NOMBRES`, y las de ortografía de la temática en `TEMATICA_ORTOGRAFIA`, ambas dentro de `build.py`.
- Si un número está escrito como texto ("1,5", "$50.000", "90 min"), el script lo convierte. Si no puede, avisa.
- Los topes del filtro de precio están en `TOPES_PRECIO` dentro de `template.html`. Con la inflación conviene revisarlos de vez en cuando.
- Cada juego tiene su propio link (`#juego=nombre-del-juego`), que se comparte desde el botón "Compartir este juego" de la ficha.
- GitHub Pages guarda la página en caché unos 10 minutos: después de subir cambios, esperar antes de mandar el link.
- Necesita `openpyxl` (`pip install openpyxl`).

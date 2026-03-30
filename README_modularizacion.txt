Archivos incluidos:
- config.py
- normalization.py
- status_rules.py
- geo_utils.py
- routing_rules.py
- excel_export.py
- sugerir_nivelacion.py

Uso:
1. Sube todos estos archivos juntos a la raiz del repo.
2. Reemplaza el archivo principal anterior por sugerir_nivelacion.py
3. Mantén app.py importando sugerir_nivelacion.generate_suggestions

Nota:
La lógica principal sigue centralizada en generate_suggestions, pero los helpers, reglas y estilos ya quedaron separados para facilitar mantenimiento.

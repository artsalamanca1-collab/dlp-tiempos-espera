# Tiempos de espera Disneyland Paris

Recogida automática cada 15 minutos de los tiempos de espera de los dos parques,
almacenamiento en CSV e informe en Excel bajo demanda.

## Estructura

```
collect.py                        toma una muestra y la añade a data/YYYY-MM-DD.csv
build_excel.py                    genera el Excel de informe desde los CSV
data/                             un CSV por día (formato largo)
.github/workflows/collect.yml     cron cada 15 min en GitHub Actions
```

## Puesta en marcha

**1. Probar en local** (no escribe nada, solo imprime el top 15):

```bash
python collect.py --dry-run
```

Si esto funciona, el resto funciona. Si falla, la API ha cambiado y hay que mirarla.

**2. Subirlo a GitHub.** Repo **público** — en privado los 2000 min/mes gratuitos se
quedan muy justos con ~60 ejecuciones diarias. Luego, en Settings → Actions → General
→ Workflow permissions, marca *Read and write permissions*, o el bot no podrá commitear.

**3. Lanzar una ejecución manual** desde la pestaña Actions (`workflow_dispatch`) para
comprobar que el commit se hace bien. A partir de ahí va solo.

**4. Generar el informe** cuando quieras:

```bash
git pull
pip install openpyxl
python build_excel.py --dias 30
```

## El Excel

- **Resumen** — media, máximo, mínimo y número de muestras por atracción, más la
  media de Single Rider.
- **Perfil horario** — espera media por hora del día. Aquí es donde se ve de verdad
  cuándo merece la pena ir a cada atracción.
- **Datos** — todas las muestras en formato largo, con autofiltro.

Las dos primeras hojas son fórmulas vivas sobre `Datos`, no valores pegados.

## Fuentes

Principal: [themeparks.wiki](https://api.themeparks.wiki/v1) — sin clave, devuelve
standby, Single Rider y Premier Access por separado.

Respaldo automático: [Queue-Times.com](https://queue-times.com/) — solo standby. Si
publicas algo con estos datos, su licencia exige mostrar "Powered by Queue-Times.com"
enlazando a su web.

## Avisos

- El cron de GitHub se retrasa entre 5 y 20 minutos en horas punta. Cada fila guarda
  la hora real de la muestra, así que no afecta al análisis.
- GitHub desactiva los workflows programados tras 60 días sin actividad en el repo.
  Los commits del propio bot no siempre cuentan: si un mes ves el repo parado, entra
  y reactívalo desde la pestaña Actions.
- Volumen aproximado: unas 60 ejecuciones diarias × ~150 filas ≈ 3 MB al mes.
  A partir del año conviene comprimir los CSV antiguos.
- Los nombres de atracción vienen de la API en inglés y no coinciden literalmente con
  los de tus notas a mano. Si quieres cruzarlos, hace falta una tabla de equivalencias.

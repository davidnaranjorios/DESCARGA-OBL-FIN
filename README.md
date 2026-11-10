# DESCARGA-OBL-FIN

Repositorio de trabajo para consultar obligaciones financieras y consolidar deuda vencida de reportes generados por la plataforma DIAN.

## Estado

**En desarrollo.** El script de consulta de DIAN original no está disponible como fuente editable. Se dispone del ejecutable original, y se recuperó parcialmente su módulo principal como bytecode Python 3.11. **El programa integrado que consulte DIAN y produzca solo Excel todavía no está implementado ni validado.**

## Contenido

- `extraer_deuda_vencida.py`: analiza PDF previamente descargados y genera un consolidado en Excel y CSV.
- `docs/inspeccion_ejecutable.txt`: inspección técnica del ejecutable.
- `requirements.txt`: dependencias del extractor.

## Ejecutar el extractor existente

Instalar dependencias:

```powershell
py -m pip install -r requirements.txt
```

Colocar los documentos en esta estructura local:

```text
D:\CONFIRMACION OBL.FIN ENDOSOS\
  Obligaciones Financieras NIT - 98137715\
    98137715_REPORTE_CONSOLIDADO.pdf
```

Ejecutar:

```powershell
py extraer_deuda_vencida.py
```

El script utiliza por defecto la ruta `D:\CONFIRMACION OBL.FIN ENDOSOS` y produce `DEUDA_VENCIDA_CONSOLIDADO.xlsx` y archivos CSV de control.

## Objetivo de la siguiente versión

1. Autenticación y consulta de NIT en DIAN mediante el flujo existente.
2. Solicitud de generación del PDF por la plataforma.
3. Lectura del PDF generado sin conservarlo permanentemente (en memoria, si el mecanismo de entrega lo permite; de lo contrario, en un espacio temporal).
4. Extracción de `DEUDA VENCIDA`, obligación, año, valor y cantidad.
5. Un único Excel con consolidado, control por NIT y advertencias.

## Seguridad

Este repositorio es privado. No subir credenciales, cookies, datos de contribuyentes, reportes PDF reales ni resultados Excel con información tributaria. Los binarios compilados y archivos de extracción no se incluyen en el control de versiones. El archivo `.gitignore` excluye sus extensiones por defecto.

## Limitaciones

La extracción disponible se validó con un PDF de ejemplo. Otros formatos de reporte requieren verificación. El proyecto aún no incluye el código fuente recuperado de Selenium ni una integración funcional.

# Ekahau_Report_Macro
Copyright © "NAITEL Proveedor de Tecnología y Redes S DE RL DE CV" 2024

Macro de Encabezados para Informes de Ekahau
Descripción

Esta macro está diseñada para usuarios que generan informes utilizando Ekahau, una herramienta popular para la planificación y análisis de redes inalámbricas. En algunos casos, todos los encabezados en los informes exportados están formateados como "Normal". Esta inconsistencia causa problemas con la generación de la tabla de contenido (TOC) y el índice del documento, ya que depende de un formato de encabezado adecuado.
Propósito

El propósito de esta macro es automatizar el proceso de cambio de los estilos de encabezado en sus informes de Ekahau a los tipos de encabezado correctos, asegurando que el índice y la tabla de contenido se generen correctamente y reflejen la estructura del documento.
Cómo Funciona

    La macro escanea todos los párrafos en el documento.
    Verifica palabras clave o patrones específicos que corresponden a diferentes niveles de encabezado.
    Dependiendo del texto que coincida, aplica los estilos de encabezado apropiados:
        Encabezado 2 (Título 2) para patrones específicos (por ejemplo, "Punto de Acceso Asociado para...").
        Encabezado 3 (Título 3) para otros patrones especificados (por ejemplo, "Fuerza de Señal para...").

Uso

    El informe de Ekahau debe ser abierto en Microsoft Word.
    El editor de VBA debe ser abierto (Alt + F11).
    El código de la macro debe ser insertado en un nuevo módulo.
    El documento debe ser guardado como un archivo habilitado para macros.
    La macro debe ser ejecutada para actualizar los estilos de encabezado.
    Nota: Si se van a unir informes grandes, es mejor que los informes sean insertados a través del objeto en lugar de copiar y pegar para evitar que la computadora se congele.

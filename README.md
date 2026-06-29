# 📊 VBA Height and Weight Data Entry

## Descripción
Este repositorio contiene un script desarrollado en Visual Basic for Applications (VBA) para Microsoft Excel. El módulo automatiza la recopilación de datos físicos (altura y peso), asegurando la integridad de la información mediante procesos de validación antes de insertarla en la hoja de cálculo. Este proyecto es una muestra práctica de automatización de tareas y control de flujo en herramientas de análisis de datos.

## 🚀 Características Principales
* **Interfaz de Usuario Interactiva:** Utiliza la función `InputBox` para solicitar al usuario que ingrese su altura (en metros) y su peso[cite: 2].
* **Validación de Tipos de Datos:** Emplea la función `IsNumeric` para verificar rigurosamente que las entradas sean números, previniendo errores de ejecución o corrupción de datos en la hoja de cálculo.
* **Manejo de Errores y Notificaciones:** 
  * Despliega un mensaje de error (`vbCritical`) si el usuario introduce texto o valores no válidos[cite: 2].
  * Muestra una confirmación de éxito (`vbInformation`) cuando los datos se procesan adecuadamente[cite: 2].
* **Procesamiento de Precisión:** Convierte las entradas a variables de doble precisión (`CDbl`) para mantener la exactitud de los decimales (especialmente útil para la altura en metros) y las asigna automáticamente a las celdas `M17` y `M18`[cite: 2].

## 🛠️ Tecnologías Utilizadas
* **Lenguaje:** Visual Basic for Applications (VBA)[cite: 2]
* **Entorno:** Microsoft Excel (Estructura XML/Binaria)[cite: 1]

## ⚙️ Instalación y Uso
1. Abre un libro de Excel y presiona `Alt + F11` para abrir el Editor de Visual Basic (VBE).
2. Inserta un nuevo módulo (`Insertar > Módulo`).
3. Copia y pega el código del procedimiento `Sub Height_and_Weight()`[cite: 2].
4. Ejecuta la macro (presionando `F5` o asignándola a un botón en la hoja).
5. Ingresa los datos solicitados en las ventanas emergentes[cite: 2].
6. Revisa las celdas `M17` y `M18` para confirmar que los datos se han insertado correctamente[cite: 2].

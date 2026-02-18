Documentación Técnica: Sistema RPA de Recaudo Digital (EAAB)
1. Resumen Ejecutivo
Este desarrollo consiste en un agente de Automatización de Procesos Robóticos (RPA) diseñado para la extracción masiva de facturación desde el portal transaccional de la Empresa de Acueducto y Alcantarillado de Bogotá (EAAB). El sistema integra navegación web automatizada, persistencia de datos en Microsoft Excel y un motor de ejecución en segundo plano (Headless).
2. Arquitectura del Sistema
El software sigue un patrón modular para facilitar el mantenimiento y la escalabilidad:
•	Capa de Orquestación (main.py): Punto de entrada que gestiona el ciclo de vida de la ejecución.
•	Capa de Configuración (Config/settings.py): Centraliza variables de entorno, rutas de directorio y parámetros de espera (timeouts).
•	Capa de Lógica de Negocio (modules/flujo_eaab.py): Contiene el motor de decisiones, gestión de frames y lógica de descarga.
•	Capa de Interfaz de Datos (modules/excel_reader.py / xlwings): Maneja el flujo de I/O con el archivo de control Excel.
________________________________________
3. Implementación del Modo "Headless" (Segundo Plano)
Uno de los mayores retos técnicos fue la ejecución invisible sin ser detectado por los sistemas de seguridad perimetral de la web (WAF).
3.1. Ingeniería de Invisibilidad
Para lograr la ejecución en segundo plano, se implementó el modo headless=new de Selenium. A diferencia del modo headless antiguo, este renderiza el DOM completo, permitiendo la interacción con elementos dinámicos.
Configuración de Stealth (Sigilo):
1.	User-Agent Spoofing: Se inyecta una cabecera de agente de usuario real para evitar que el servidor identifique la conexión como un proceso de Python.
2.	Exclusión de Switches de Automatización: Se eliminan las banderas enable-automation y el objeto navigator.webdriver mediante la inyección de JavaScript en el contexto del navegador.
3.	Virtual Viewport: Se define una resolución de ventana virtual (1920x1080) para asegurar que los elementos responsivos de la página se rendericen correctamente y sean "clicables".
Python
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("user-agent=Mozilla/5.0 ...")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
________________________________________
4. Flujo de Control de Descargas y Persistencia
El bot implementa un algoritmo de verificación de estado para garantizar la integridad de los datos.
4.1. Ciclo de Vida de Procesamiento
1.	Polling de Excel: Se filtran los registros cuya columna Descargada sea igual a NO.
2.	Gestión de Frames: Dado que el portal EAAB utiliza arquitecturas de iframes anidados, el bot realiza un "context switching" dinámico para acceder a los formularios de consulta.
3.	Validación de Descarga Física: El bot no asume que la descarga inició; monitorea el sistema de archivos local, comparando el estado del directorio antes y después del evento de clic, realizando el renombrado del archivo mediante el manejo de excepciones de OS.
________________________________________
5. Requisitos de Instalación y Despliegue
5.1. Stack Tecnológico
•	Lenguaje: Python 3.12+
•	Framework de Automatización: Selenium 4.x
•	Gestión de Drivers: webdriver-manager (Instalación binaria automática).
•	Interoperabilidad Excel: xlwings (Para manipulación de instancias de Excel en tiempo real).
5.2. Instalación de Dependencias
Bash
pip install selenium pandas xlwings webdriver-manager openpyxl
5.3. Configuración de la Macro (Trigger)
El bot se dispara mediante una macro VBA utilizando el objeto WScript.Shell, permitiendo que un usuario administrativo ejecute un proceso complejo de Python con un solo clic, manteniendo la consola (cmd.exe /K) visible para auditoría técnica.
________________________________________
6. Manejo de Errores y Resiliencia
•	Timeout Handling: Uso de WebDriverWait con estrategias de Expected Conditions (EC) para manejar la latencia de la red.
•	Auto-Recuperación: En caso de fallo en un contrato específico, el bot captura la excepción, marca el error en Excel y reinicia el flujo hacia el siguiente contrato sin detener el proceso global.
________________________________________
Desarrollado por: Anyi Daniela Manco Henao 
Versión: 1.0.0 
Licencia: Propietaria - Uso Interno

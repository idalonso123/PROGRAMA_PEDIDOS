#!/usr/bin/env python3
"""
Módulo EmailService - Servicio de envío de correos electrónicos
Este módulo gestiona el envío de correos electrónicos con archivos adjuntos
utilizando SMTP. Los destinatarios se configuran exclusivamente en config.json
y los nombres de los encargados se leen desde config/encargados.json.
Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-02-05
"""
import smtplib
import ssl
import os
import json
import logging
import pandas as pd
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from typing import Optional, List, Dict, Any
from pathlib import Path

# Configuración del logger
logger = logging.getLogger(__name__)


class EmailService:
    """
    Servicio de envío de correos electrónicos para el sistema de pedidos.
    
    Esta clase encapsula toda la lógica relacionada con el envío de correos:
    configuración SMTP, gestión de destinatarios, generación de mensajes
    y envío de adjuntos.
    
    Los destinatarios se configuran exclusivamente en config.json.
    Los nombres de los encargados se leen desde config/encargados.json.
    
    Attributes:
        config (dict): Configuración del sistema
        smtp_config (dict): Configuración del servidor SMTP
        remitente (dict): Información del remitente
        destinatarios (dict): Mapeo sección -> lista de destinatarios desde config.json
        plantilla_asunto (str): Template del asunto del email
        plantilla_cuerpo (str): Template del cuerpo del email
        encargados_por_seccion (dict): Mapeo de secciones a nombres de encargados
    """
    
    def __init__(self, config: dict):
        """
        Inicializa el EmailService con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
        """
        self.config = config
        self.smtp_config = {}
        self.remitente = {}
        self.destinatarios = {}
        self.plantilla_asunto = ""
        self.plantilla_cuerpo = ""
        self.encargados_por_seccion = {}
        
        # Cargar configuración
        self._cargar_configuracion()
        
        # Cargar nombres de encargados desde archivo JSON
        self._cargar_encargados()
        
        logger.info("EmailService inicializado correctamente")
        logger.info(f"Remitente: {self.remitente.get('email', 'No configurado')}")
        logger.info(f"Secciones con destinatarios: {list(self.destinatarios.keys())}")
    
    def _cargar_configuracion(self):
        """
        Carga la configuración del email desde el diccionario config.
        Extrae configuración SMTP, remitente, destinatarios y templates.
        Los destinatarios se leen exclusivamente desde config.json.
        """
        email_config = self.config.get('email', {})
        
        # Configuración SMTP
        self.smtp_config = {
            'servidor': email_config.get('servidor', 'smtp.serviciodecorreo.es'),
            'puerto': email_config.get('puerto', 465),
            'usar_ssl': email_config.get('usar_ssl', True),
            'usar_tls': email_config.get('usar_tls', False)
        }
        
        # Remitente
        self.remitente = {
            'email': email_config.get('remitente', {}).get('email', 'ivan.delgado@viveverde.es'),
            'nombre': email_config.get('remitente', {}).get('nombre', 'Sistema de Pedidos VIVEVERDE')
        }
        
        # Destinatarios - SE LEEN EXCLUSIVAMENTE DESDE CONFIG.JSON
        self.destinatarios = email_config.get('destinatarios', {})
        
        # Plantillas
        self.plantilla_asunto = email_config.get('plantillas', {}).get(
            'asunto', 
            'VIVEVERDE: Pedido de compra - Semana {semana} - {seccion}'
        )
        self.plantilla_cuerpo = email_config.get('plantillas', {}).get(
            'cuerpo', 
            'Buenos días {nombre_encargado}. Te adjunto el pedido de compra generado '
            'para la semana {semana} de la sección {seccion}. Atentamente, '
            'Sistema de Pedidos automáticos VIVEVERDE.'
        )
        
        logger.debug("Configuración de email cargada desde config.json")
    
    def _cargar_encargados(self):
        """
        Carga información de encargados desde archivo JSON.
        NOTA: Los nombres de los encargados ahora se leen desde config/encargados.json
        en lugar del archivo Excel. Esto permite cambios futuros sin modificar código.
        
        El archivo JSON tiene la siguiente estructura:
        {
            "version": "1.0",
            "descripcion": "Configuración de encargados por sección",
            "encargados": {
                "interior": "Iris",
                "mascotas_manufacturado": "María",
                ...
            }
        }
        """
        try:
            # Construir ruta al archivo de configuración de encargados
            directorio_base = self.config.get('rutas', {}).get('directorio_base', '.')
            ruta_archivo = Path(directorio_base) / 'config' / 'encargados.json'
            
            # Verificar que el archivo existe
            if not ruta_archivo.exists():
                logger.warning(f"Archivo de encargados no encontrado: {ruta_archivo}")
                logger.info("Se usarán nombres genéricos para los encargados")
                return
            
            # Leer el archivo JSON
            with open(ruta_archivo, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # Extraer el diccionario de encargados
            self.encargados_por_seccion = config_data.get('encargados', {})
            
            logger.info(f"Encargados cargados desde {ruta_archivo}: {len(self.encargados_por_seccion)} secciones")
            logger.debug(f"Mapping de encargados: {self.encargados_por_seccion}")
            
        except json.JSONDecodeError as e:
            logger.error(f"Error al parsear archivo de encargados (JSON inválido): {e}")
            logger.info("Se usarán nombres genéricos para los encargados")
        except Exception as e:
            logger.error(f"Error al leer archivo de encargados: {e}")
            logger.info("Se usarán nombres genéricos para los encargados")
    
    def enviar_resumen_gestion(self, semana: int, archivo_resumen: str) -> Dict[str, Any]:
        """
        Envía el resumen de pedidos de compra a los responsables de gestión.
        
        Destinatarios:
        - Sandra: ivan.delgado@viveverde.es
        - Ivan: ivan.delgado@viveverde.es
        - Pedro: ivan.delgado@viveverde.es
        
        Args:
            semana (int): Número de semana procesada
            archivo_resumen (str): Ruta al archivo de resumen consolidado
            
        Returns:
            Dict[str, Any]: Resultado del envío con estado y detalles
        """
        logger.info("\n" + "=" * 60)
        logger.info("ENVÍO DE RESUMEN A RESPONSABLES DE GESTIÓN")
        logger.info("=" * 60)
        
        # Destinatarios del resumen de gestión
        destinatarios_resumen = [
            {'nombre': 'Sandra', 'email': 'sandra.delgado@viveverde.es'},
            {'nombre': 'Ivan', 'email': 'ivan.delgado@viveverde.es'},
            {'nombre': 'Pedro', 'email': 'pedro.delgado@viveverde.es'}
        ]
        
        # Verificar que el archivo de resumen existe
        if not Path(archivo_resumen).exists():
            logger.warning(f"Archivo de resumen no encontrado: {archivo_resumen}")
            logger.info("Omitiendo envío de resumen a responsables de gestión")
            return {'enviado': False, 'razon': 'archivo_no_encontrado'}
        
        try:
            # Generar asunto y cuerpo del email
            asunto = f"Viveverde: Resumen de pedidos de compra de las secciones semana {semana}"
            
            resultados_envio = {}
            emails_enviados = 0
            emails_fallidos = 0
            
            for destinatario in destinatarios_resumen:
                nombre = destinatario['nombre']
                email = destinatario['email']
                
                # Generar cuerpo personalizado para cada destinatario
                cuerpo = (f"Buenos días {nombre}.\n\n"
                         f"Te adjunto el resumen de los pedidos de compra de cada sección "
                         f"de la semana {semana}.\n\n"
                         f"Atentamente,\n"
                         f"Sistema de Pedidos automáticos VIVEVERDE.")
                
                # Crear y enviar mensaje
                msg = self._crear_mensaje([email], asunto, cuerpo, [archivo_resumen])
                enviado = self._enviar_email(msg)
                
                if enviado:
                    logger.info(f"✓ Resumen enviado a {nombre} ({email})")
                    emails_enviados += 1
                    resultados_envio[nombre] = {'email': email, 'enviado': True}
                else:
                    logger.error(f"✗ Error al enviar resumen a {nombre} ({email})")
                    emails_fallidos += 1
                    resultados_envio[nombre] = {'email': email, 'enviado': False}
            
            logger.info("\n" + "=" * 60)
            logger.info("RESUMEN DE ENVÍO A RESPONSABLES DE GESTIÓN")
            logger.info("=" * 60)
            logger.info(f"Emails enviados: {emails_enviados}")
            logger.info(f"Emails fallidos: {emails_fallidos}")
            
            return {
                'enviado': emails_enviados > 0,
                'emails_enviados': emails_enviados,
                'emails_fallidos': emails_fallidos,
                'resultados': resultados_envio
            }
            
        except Exception as e:
            logger.error(f"Error al enviar resumen de gestión: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            return {'enviado': False, 'error': str(e)}
    
    def _normalizar_seccion(self, seccion: str) -> str:
        """
        Normaliza el nombre de una sección para buscar en el mapeo de encargados.
        
        Args:
            seccion (str): Nombre de la sección a normalizar
            
        Returns:
            str: Nombre de sección normalizado
        """
        return seccion.strip().lower().replace(' ', '_')
    
    def _obtener_password(self) -> str:
        """
        Obtiene la contraseña del remitente desde variable de entorno.
        
        Returns:
            str: Contraseña del correo remitente
            
        Raises:
            ValueError: Si la variable de entorno no está configurada
        """
        email_config = self.config.get('email', {})
        password_var = email_config.get('password_var', 'EMAIL_PASSWORD')
        
        password = os.environ.get(password_var)
        
        if not password:
            error_msg = (f"Variable de entorno '{password_var}' no configurada. "
                        "Configure la contraseña antes de enviar emails.")
            logger.error(error_msg)
            raise ValueError(error_msg)
        
        return password
    
    def _generar_asunto(self, semana: int, seccion: str) -> str:
        """
        Genera el asunto del email usando la plantilla configurada.
        
        Args:
            semana (int): Número de semana
            seccion (str): Nombre de la sección
            
        Returns:
            str: Asunto formateado
        """
        return self.plantilla_asunto.format(semana=semana, seccion=seccion)
    
    def _generar_cuerpo(self, semana: int, seccion: str, nombre_encargado: str) -> str:
        """
        Genera el cuerpo del email usando la plantilla configurada.
        
        Args:
            semana (int): Número de semana
            seccion (str): Nombre de la sección
            nombre_encargado (str): Nombre del encargado
            
        Returns:
            str: Cuerpo del mensaje formateado
        """
        return self.plantilla_cuerpo.format(
            semana=semana,
            seccion=seccion,
            nombre_encargado=nombre_encargado
        )
    
    def _crear_mensaje(self, destinatarios: List[str], asunto: str, 
                      cuerpo: str, archivos_adjuntos: List[str]) -> MIMEMultipart:
        """
        Crea el mensaje MIME con adjuntos.
        
        Args:
            destinatarios (List[str]): Lista de correos destinatarios
            asunto (str): Asunto del email
            cuerpo (str): Cuerpo del mensaje
            archivos_adjuntos (List[str]): Lista de rutas de archivos a adjuntar
            
        Returns:
            MIMEMultipart: Mensaje MIME listo para enviar
        """
        msg = MIMEMultipart()
        msg['From'] = f"{self.remitente['nombre']} <{self.remitente['email']}>"
        msg['To'] = ', '.join(destinatarios)
        msg['Subject'] = asunto
        
        # Adjuntar cuerpo en texto plano
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
        # Adjuntar archivos
        for archivo in archivos_adjuntos:
            if not Path(archivo).exists():
                logger.warning(f"Archivo no encontrado: {archivo}")
                continue
            
            try:
                # Determinar tipo de archivo
                extension = Path(archivo).suffix.lower()
                
                if extension in ['.xlsx', '.xls']:
                    mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                elif extension == '.csv':
                    mime_type = 'text/csv'
                elif extension == '.pdf':
                    mime_type = 'application/pdf'
                else:
                    mime_type = 'application/octet-stream'
                
                # Leer archivo
                with open(archivo, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                
                # Codificar
                encoders.encode_base64(part)
                
                # Añadir header
                filename = Path(archivo).name
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= "{filename}"'
                )
                part.add_header('Content-Type', mime_type)
                
                msg.attach(part)
                logger.debug(f"Adjunto añadido: {filename}")
                
            except Exception as e:
                logger.error(f"Error al adjuntar archivo {archivo}: {e}")
        
        return msg
    
    def _enviar_email(self, msg: MIMEMultipart) -> bool:
        """
        Envía el email a través del servidor SMTP.
        
        Args:
            msg (MIMEMultipart): Mensaje MIME a enviar
            
        Returns:
            bool: True si el envío fue exitoso, False en caso contrario
        """
        try:
            password = self._obtener_password()
            
            # Crear contexto SSL
            context = ssl.create_default_context()
            
            # Conectar al servidor
            if self.smtp_config.get('usar_ssl', True):
                with smtplib.SMTP_SSL(
                    self.smtp_config['servidor'],
                    self.smtp_config['puerto'],
                    context=context
                ) as server:
                    server.login(self.remitente['email'], password)
                    server.sendmail(
                        self.remitente['email'],
                        msg['To'].split(', '),
                        msg.as_string()
                    )
            else:
                with smtplib.SMTP(
                    self.smtp_config['servidor'],
                    self.smtp_config['puerto']
                ) as server:
                    if self.smtp_config.get('usar_tls', False):
                        server.starttls(context=context)
                    server.login(self.remitente['email'], password)
                    server.sendmail(
                        self.remitente['email'],
                        msg['To'].split(', '),
                        msg.as_string()
                    )
            
            logger.info(f"Email enviado exitosamente a: {msg['To']}")
            return True
            
        except smtplib.SMTPException as e:
            logger.error(f"Error SMTP al enviar email: {e}")
            return False
        except Exception as e:
            logger.error(f"Error al enviar email: {e}")
            return False
    
    def obtener_destinatarios_seccion(self, seccion: str) -> List[Dict[str, str]]:
        """
        Obtiene la lista de destinatarios para una sección específica.
        
        Los correos electrónicos se obtienen EXCLUSIVAMENTE desde config.json.
        Los nombres se obtienen desde config/encargados.json.
        
        Args:
            seccion (str): Nombre de la sección
            
        Returns:
            List[Dict[str, str]]: Lista de diccionarios con 'email' y 'nombre'
        """
        seccion_normalizada = self._normalizar_seccion(seccion)
        
        # Obtener correos desde config.json (única fuente)
        correos = self.destinatarios.get(seccion_normalizada, [])
        
        if isinstance(correos, str):
            correos = [c.strip() for c in correos.split(',')]
        
        if not correos:
            logger.warning(f"No hay destinatarios configurados para la sección: {seccion}")
            return []
        
        # Obtener nombre del encargado desde config/encargados.json
        nombre_encargado = self.encargados_por_seccion.get(seccion_normalizada, "")
        
        # Si no hay nombre en el JSON, usar un nombre genérico
        if not nombre_encargado:
            nombre_encargado = "Encargado"
        
        # Construir lista de destinatarios con nombres
        destinatarios = []
        for correo in correos:
            destinatarios.append({
                'email': correo.strip(),
                'nombre': nombre_encargado
            })
        
        logger.debug(f"Destinatarios para {seccion}: {destinatarios}")
        return destinatarios
    
    def enviar_pedido_por_seccion(self, semana: int, seccion: str, 
                                  archivos: List[str]) -> Dict[str, Any]:
        """
        Envía los archivos de pedido por email al responsable de la sección.
        
        Args:
            semana (int): Número de semana
            seccion (str): Nombre de la sección
            archivos (List[str]): Lista de rutas de archivos a adjuntar
            
        Returns:
            Dict[str, Any]: Resultado del envío con estado y detalles
        """
        resultado = {
            'seccion': seccion,
            'semana': semana,
            'enviado': False,
            'destinatarios': [],
            'archivos_adjuntos': [],
            'error': None
        }
        
        # Obtener destinatarios
        destinatarios = self.obtener_destinatarios_seccion(seccion)
        
        if not destinatarios:
            resultado['error'] = "No hay destinatarios configurados"
            logger.warning(f"No se puede enviar email para {seccion}: {resultado['error']}")
            return resultado
        
        # Filtrar archivos existentes
        archivos_existentes = [f for f in archivos if Path(f).exists()]
        resultado['archivos_adjuntos'] = [Path(f).name for f in archivos_existentes]
        
        if not archivos_existentes:
            resultado['error'] = "No hay archivos para adjuntar"
            logger.warning(f"No se puede enviar email para {seccion}: {resultado['error']}")
            return resultado
        
        # Generar asunto y cuerpo
        asunto = self._generar_asunto(semana, seccion)
        cuerpo = self._generar_cuerpo(semana, seccion, destinatarios[0]['nombre'])
        
        # Extraer solo los correos
        lista_correos = [d['email'] for d in destinatarios]
        resultado['destinatarios'] = lista_correos
        
        # Crear y enviar mensaje
        msg = self._crear_mensaje(lista_correos, asunto, cuerpo, archivos_existentes)
        resultado['enviado'] = self._enviar_email(msg)
        
        if resultado['enviado']:
            logger.info(f"Email enviado para sección {seccion} (semana {semana})")
        else:
            resultado['error'] = "Error en el envío"
        
        return resultado
    
    def enviar_resumen_centralizado(self, semana: int, 
                                    archivos: Dict[str, List[str]]) -> Dict[str, Any]:
        """
        Envía todos los pedidos de forma centralizada al responsable principal.
        
        Args:
            semana (int): Número de semana
            archivos (Dict[str, List[str]]): Archivos por sección
            
        Returns:
            Dict[str, Any]: Resultado del envío
        """
        email_config = self.config.get('email', {})
        email_centralizado = email_config.get('email_centralizado')
        
        if not email_centralizado:
            logger.info("No hay configuración de email centralizado, omitiendo envío")
            return {'enviado': False, 'razon': 'No configurado'}
        
        # Recopilar todos los archivos
        todos_archivos = []
        for seccion, archivos_seccion in archivos.items():
            todos_archivos.extend(archivos_seccion)
        
        if not todos_archivos:
            return {'enviado': False, 'razon': 'No hay archivos'}
        
        # Generar cuerpo del resumen
        cuerpo = f"Resumen de pedidos para la semana {semana}:\n\n"
        for seccion in archivos.keys():
            cuerpo += f"- {seccion}\n"
        cuerpo += "\nArchivos adjuntos.\n"
        
        asunto = f"VIVEVERDE: Resumen de Pedidos - Semana {semana}"
        
        # Filtrar archivos existentes
        archivos_existentes = [f for f in todos_archivos if Path(f).exists()]
        
        # Crear y enviar mensaje
        msg = self._crear_mensaje([email_centralizado], asunto, cuerpo, archivos_existentes)
        enviado = self._enviar_email(msg)
        
        return {
            'enviado': enviado,
            'destinatario': email_centralizado,
            'archivos': len(archivos_existentes)
        }
    
    def verificar_configuracion(self) -> Dict[str, Any]:
        """
        Verifica que la configuración del email sea correcta.
        
        Returns:
            Dict[str, Any]: Resultado de la verificación
        """
        resultado = {
            'valido': True,
            'remitente': {},
            'smtp': {},
            'destinatarios': {},
            'problemas': []
        }
        
        # Verificar remitente
        if self.remitente.get('email'):
            resultado['remitente'] = {
                'email': self.remitente['email'],
                'nombre': self.remitente.get('nombre', ''),
                'configurado': True
            }
        else:
            resultado['problemas'].append("Email del remitente no configurado")
            resultado['valido'] = False
        
        # Verificar SMTP
        resultado['smtp'] = {
            'servidor': self.smtp_config.get('servidor', ''),
            'puerto': self.smtp_config.get('puerto', ''),
            'ssl': self.smtp_config.get('usar_ssl', True),
            'tls': self.smtp_config.get('usar_tls', False)
        }
        
        # Verificar destinatarios
        secciones_configuradas = []
        for seccion, correos in self.destinatarios.items():
            if isinstance(correos, str):
                correos = [c.strip() for c in correos.split(',') if c.strip()]
            elif isinstance(correos, list):
                correos = [c.strip() for c in correos if c.strip()]
            else:
                correos = []
            
            if correos:
                secciones_configuradas.append(seccion)
                resultado['destinatarios'][seccion] = {
                    'correos': correos,
                    'nombre_encargado': self.encargados_por_seccion.get(seccion, 'No definido')
                }
        
        resultado['secciones_configuradas'] = len(secciones_configuradas)
        
        if not secciones_configuradas:
            resultado['problemas'].append("No hay destinatarios configurados en config.json")
            resultado['valido'] = False
        
        # Verificar contraseña
        try:
            password = self._obtener_password()
            resultado['password'] = {'configurado': True, 'variable': 'EMAIL_PASSWORD'}
        except ValueError:
            resultado['password'] = {'configurado': False}
            resultado['problemas'].append("Contraseña no configurada (variable EMAIL_PASSWORD)")
        
        return resultado


def crear_email_service(config: dict) -> EmailService:
    """
    Crea una instancia del EmailService.
    
    Args:
        config (dict): Configuración del sistema
        
    Returns:
        EmailService: Instancia inicializada del servicio de email
    """
    return EmailService(config)


def verificar_configuracion_email(config: dict) -> Dict[str, Any]:
    """
    Verifica la configuración de email sin crear el servicio.
    
    Args:
        config (dict): Configuración del sistema
        
    Returns:
        Dict[str, Any]: Resultado de la verificación
    """
    try:
        email_config = config.get('email', {})
        
        # Verificar destinatarios desde config.json
        destinatarios = email_config.get('destinatarios', {})
        secciones_con_correo = []
        
        for seccion, correos in destinatarios.items():
            if correos:
                secciones_con_correo.append(seccion)
        
        return {
            'valido': len(secciones_con_correo) > 0,
            'secciones_configuradas': len(secciones_con_correo),
            'secciones': secciones_con_correo,
            'remitente': email_config.get('remitente', {}).get('email', 'No configurado'),
            'smtp': email_config.get('servidor', 'No configurado'),
            'mensaje': ('Configuración verificada. Los correos se leen desde config.json. '
                       'Los nombres de encargados se leen desde config/encargados.json.')
        }
        
    except Exception as e:
        return {
            'valido': False,
            'error': str(e)
        }


if __name__ == "__main__":
    # Ejemplo de uso
    print("EmailService - Módulo de envío de correos")
    print("=" * 50)
    
    # Configurar logging
    logging.basicConfig(level=logging.INFO)
    
    # Ejemplo de configuración
    config_ejemplo = {
        'email': {
            'servidor': 'smtp.serviciodecorreo.es',
            'puerto': 465,
            'usar_ssl': True,
            'remitente': {
                'email': 'ivan.delgado@viveverde.es',
                'nombre': 'Sistema de Pedidos VIVEVERDE'
            },
            'destinatarios': {
                'maf': 'ivan.delgado@viveverde.es',
                'interior': 'ivan.delgado@viveverde.es',
                'mascotas_vivo': 'ivan.delgado@viveverde.es',
                'mascotas_manufacturado': 'ivan.delgado@viveverde.es',
                'deco_interior': 'ivan.delgado@viveverde.es',
                'deco_exterior': 'ivan.delgado@viveverde.es',
                'tierras_aridos': 'ivan.delgado@viveverde.es',
                'fitos': 'ivan.delgado@viveverde.es',
                'semillas': 'ivan.delgado@viveverde.es',
                'utiles_jardin': 'ivan.delgado@viveverde.es',
                'vivero': 'ivan.delgado@viveverde.es'
            },
            'plantillas': {
                'asunto': 'VIVEVERDE: Pedido de compra - Semana {semana} - {seccion}',
                'cuerpo': 'Buenos días {nombre_encargado}. Te adjunto el pedido de compra '
                         'generado para la semana {semana} de la sección {seccion}. '
                         'Atentamente, Sistema de Pedidos automáticos VIVEVERDE.'
            }
        },
        'archivos_entrada': {
            'archivo_encargados': None
        },
        'rutas': {
            'directorio_base': '.'
        }
    }
    
    # Verificar configuración
    print("\nVerificación de configuración:")
    verificacion = verificar_configuracion_email(config_ejemplo)
    for clave, valor in verificacion.items():
        print(f"  {clave}: {valor}")
    
    # Crear servicio
    print("\nCreando EmailService:")
    email_service = crear_email_service(config_ejemplo)
    
    # Verificar configuración completa
    print("\nVerificación completa:")
    verif = email_service.verificar_configuracion()
    for clave, valor in verif.items():
        print(f"  {clave}: {valor}")
    
    # Ejemplo de destinatarios
    print("\nEjemplo de destinatarios para sección 'maf':")
    destinatarios = email_service.obtener_destinatarios_seccion('maf')
    for d in destinatarios:
        print(f"  - Email: {d['email']}, Nombre: {d['nombre']}")
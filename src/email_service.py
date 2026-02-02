#!/usr/bin/env python3
"""
Módulo EmailService - Servicio de envío de emails con adjuntos

Este módulo proporciona funcionalidad completa para enviar emails con archivos
adjuntos utilizando SMTP. Implementa toda la lógica de configuración, templating
de mensajes y manejo de adjuntos.

Características:
- Configuración externa desde JSON (no hardcodeado)
- Soporte para múltiples destinatarios por sección
- Templates de asunto y cuerpo configurables
- Adjunto automático de archivos generados
- Manejo de errores robusto
- Logging detallado

Autor: Sistema de Pedidos Vivero V2
Fecha: 2026-02-03
"""

import smtplib
import ssl
import os
import logging
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
    Servicio de envío de emails para el sistema de pedidos.
    
    Esta clase encapsula toda la lógica necesaria para enviar emails
    con archivos adjuntos, incluyendo configuración SMTP, templates
    de mensajes y gestión de destinatarios por sección.
    
    Attributes:
        config (dict): Configuración del sistema
        smtp_config (dict): Configuración del servidor SMTP
        remitente (dict): Información del remitente
        destinatarios (dict): Mapeo sección -> lista de destinatarios
        plantilla_asunto (str): Template del asunto del email
        plantilla_cuerpo (str): Template del cuerpo del email
    """
    
    def __init__(self, config: dict):
        """
        Inicializa el EmailService con la configuración proporcionada.
        
        Args:
            config (dict): Diccionario con la configuración del sistema
        """
        self.config = config
        self._cargar_configuracion()
        logger.info("EmailService inicializado correctamente")
    
    def _cargar_configuracion(self):
        """
        Carga la configuración del email desde el diccionario config.
        
        Extrae la configuración SMTP, remitente, destinatarios y templates
        desde la configuración global. Todos los datos son externos y
        configurables, sin hardcoding.
        """
        # Configuración del servidor SMTP
        email_config = self.config.get('email', {})
        self.smtp_config = email_config.get('smtp', {
            'servidor': 'smtp.serviciodecorreo.es',
            'puerto': 465,
            'usar_tls': True,
            'usar_ssl': True
        })
        
        # Información del remitente
        self.remitente = email_config.get('remitente', {
            'email': 'ivan.delgado@viveverde.es',
            'nombre': 'Sistema de Pedidos VIVEVERDE'
        })
        
        # Destinatarios por sección (mapeo sección -> lista de destinatarios)
        self.destinatarios = email_config.get('destinatarios', {})
        
        # Templates de mensaje
        self.plantilla_asunto = email_config.get('plantillas', {}).get(
            'asunto',
            'VIVEVERDE: Pedido de compra - Semana {semana} - {seccion}'
        )
        self.plantilla_cuerpo = email_config.get('plantillas', {}).get(
            'cuerpo',
            'Buenos días {nombre_encargado}. \n'
            'Te adjunto el pedido de compra generado para la semana {semana} de la sección {seccion}.\n\n'
            'Atentamente,\n'
            'Sistema de Pedidos automáticos VIVEVERDE.'
        )
        
        logger.debug(f"[DEBUG] SMTP configurado: {self.smtp_config}")
        logger.debug(f"[DEBUG] Remitente: {self.remitente}")
        logger.debug(f"[DEBUG] Destinatarios: {self.destinatarios}")
    
    def _obtener_password(self) -> str:
        """
        Obtiene la contraseña del remitente desde variable de entorno.
        
        Returns:
            str: Contraseña del correo remitente
            
        Raises:
            ValueError: Si la variable de entorno no está configurada
        """
        password = os.environ.get('EMAIL_PASSWORD')
        
        if not password:
            error_msg = (
                "La variable de entorno 'EMAIL_PASSWORD' no está configurada. "
                "Por favor, configúrala antes de enviar emails."
            )
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
            nombre_encargado (str): Nombre del responsable
            
        Returns:
            str: Cuerpo formateado
        """
        return self.plantilla_cuerpo.format(
            semana=semana,
            seccion=seccion,
            nombre_encargado=nombre_encargado
        )
    
    def _crear_mensaje(
        self,
        destinatarios: List[str],
        asunto: str,
        cuerpo: str,
        archivos_adjuntos: List[str]
    ) -> MIMEMultipart:
        """
        Crea el mensaje MIME con adjuntos.
        
        Args:
            destinatarios (List[str]): Lista de direcciones email
            asunto (str): Asunto del mensaje
            cuerpo (str): Cuerpo del mensaje
            archivos_adjuntos (List[str]): Lista de rutas de archivos a adjuntar
            
        Returns:
            MIMEMultipart: Mensaje preparado para enviar
        """
        # Crear mensaje multipart
        msg = MIMEMultipart()
        msg['From'] = f"{self.remitente['nombre']} <{self.remitente['email']}>"
        msg['To'] = ', '.join(destinatarios)
        msg['Subject'] = asunto
        
        # Adjuntar cuerpo en texto plano
        msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
        
        # Adjuntar archivos
        for archivo in archivos_adjuntos:
            if os.path.exists(archivo):
                try:
                    # Determinar tipo MIME basado en extensión
                    nombre_archivo = os.path.basename(archivo)
                    if archivo.endswith('.xlsx') or archivo.endswith('.xls'):
                        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    elif archivo.endswith('.csv'):
                        mime_type = 'text/csv'
                    elif archivo.endswith('.pdf'):
                        mime_type = 'application/pdf'
                    else:
                        mime_type = 'application/octet-stream'
                    
                    # Leer archivo y crear parte MIME
                    with open(archivo, 'rb') as f:
                        parte = MIMEBase(mime_type, 'octet-stream')
                        parte.set_payload(f.read())
                    
                    # Codificar en Base64
                    encoders.encode_base64(parte)
                    
                    # Añadir header del archivo
                    parte.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {nombre_archivo}'
                    )
                    
                    msg.attach(parte)
                    logger.info(f"Archivo adjuntado: {nombre_archivo}")
                    
                except Exception as e:
                    logger.error(f"Error al adjuntar archivo {archivo}: {str(e)}")
            else:
                logger.warning(f"Archivo no encontrado para adjuntar: {archivo}")
        
        return msg
    
    def _enviar_email(self, msg: MIMEMultipart) -> bool:
        """
        Envía el email a través del servidor SMTP.
        
        Args:
            msg (MIMEMultipart): Mensaje preparado
            
        Returns:
            bool: True si el envío fue exitoso, False en caso contrario
        """
        try:
            password = self._obtener_password()
            
            # Crear contexto SSL
            contexto = ssl.create_default_context()
            
            # Conectar al servidor SMTP
            if self.smtp_config.get('usar_ssl', True):
                with smtplib.SMTP_SSL(
                    self.smtp_config['servidor'],
                    self.smtp_config['puerto'],
                    context=contexto
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
                    server.starttls(context=contexto)
                    server.login(self.remitente['email'], password)
                    server.sendmail(
                        self.remitente['email'],
                        msg['To'].split(', '),
                        msg.as_string()
                    )
            
            logger.info(f"Email enviado exitosamente a: {msg['To']}")
            return True
            
        except smtplib.SMTPException as e:
            logger.error(f"Error SMTP al enviar email: {str(e)}")
            return False
        except Exception as e:
            logger.error(f"Error al enviar email: {str(e)}")
            return False
    
    def obtener_destinatarios_seccion(self, seccion: str) -> List[Dict[str, str]]:
        """
        Obtiene la lista de destinatarios para una sección específica.
        
        Args:
            seccion (str): Nombre de la sección
            
        Returns:
            List[Dict]: Lista de diccionarios con 'email' y 'nombre' de cada destinatario
        """
        destinatarios_config = self.destinatarios.get(seccion, [])
        
        # Si es string, convertir a lista
        if isinstance(destinatarios_config, str):
            destinatarios_config = [destinatarios_config]
        
        # Si es lista de strings, convertir a formato con nombre genérico
        resultado = []
        for dest in destinatorios_config:
            if isinstance(dest, str):
                resultado.append({
                    'email': dest,
                    'nombre': 'Encargado'  # Nombre genérico
                })
            elif isinstance(dest, dict):
                resultado.append(dest)
        
        logger.debug(f"[DEBUG] Destinatarios para {seccion}: {resultado}")
        return resultado
    
    def enviar_pedido_por_seccion(
        self,
        semana: int,
        seccion: str,
        archivos: List[str]
    ) -> Dict[str, Any]:
        """
        Envía los archivos de pedido por email al responsable de la sección.
        
        Args:
            semana (int): Número de semana
            seccion (str): Nombre de la sección
            archivos (List[str]): Lista de rutas de archivos a enviar
            
        Returns:
            Dict: Resultado del envío con métricas
        """
        logger.info("=" * 60)
        logger.info(f"ENVIANDO EMAIL PARA SECCIÓN: {seccion.upper()}")
        logger.info("=" * 60)
        
        # Obtener destinatarios
        destinatarios_info = self.obtener_destinatarios_seccion(seccion)
        
        if not destinatarios_info:
            logger.warning(f"No hay destinatarios configurados para la sección: {seccion}")
            return {
                'exito': False,
                'seccion': seccion,
                'error': 'sin_destinatarios',
                'mensaje': f'No hay destinatarios configurados para {seccion}'
            }
        
        # Extraer solo los emails
        destinatarios = [d['email'] for d in destinatarios_info]
        nombres = [d['nombre'] for d in destinatarios_info]
        
        logger.info(f"Destinatarios: {destinatarios}")
        logger.info(f"Archivos a enviar: {archivos}")
        
        # Generar asunto y cuerpo
        asunto = self._generar_asunto(semana, seccion)
        
        # Usar el primer nombre o nombre genérico
        nombre_encargado = nombres[0] if nombres else 'Encargado'
        cuerpo = self._generar_cuerpo(semana, seccion, nombre_encargado)
        
        # Crear mensaje
        msg = self._crear_mensaje(destinatarios, asunto, cuerpo, archivos)
        
        # Enviar email
        exito = self._enviar_email(msg)
        
        # Preparar resultado
        resultado = {
            'exito': exito,
            'seccion': seccion,
            'semana': semana,
            'destinatarios': destinatarios,
            'archivos_enviados': archivos if exito else [],
            'asunto': asunto
        }
        
        if exito:
            logger.info(f"✓ Email enviado exitosamente para {seccion}")
        else:
            logger.error(f"✗ Error al enviar email para {seccion}")
            resultado['error'] = 'error_envio'
        
        return resultado
    
    def enviar_resumen_centralizado(
        self,
        semana: int,
        archivos: Dict[str, List[str]]
    ) -> Dict[str, Any]:
        """
        Envía todos los pedidos de forma centralizada al responsable principal.
        
        Args:
            semana (int): Número de semana
            archivos (Dict): Mapeo sección -> lista de archivos
            
        Returns:
            Dict: Resultado del envío
        """
        logger.info("=" * 60)
        logger.info("ENVIANDO RESUMEN CENTRALIZADO")
        logger.info("=" * 60)
        
        # Obtener destinatario centralizado
        email_central = self.config.get('email', {}).get('email_centralizado')
        
        if not email_central:
            logger.warning("No hay email centralizado configurado")
            return {
                'exito': False,
                'error': 'sin_email_centralizado'
            }
        
        # Recopilar todos los archivos
        todos_archivos = []
        for archivos_seccion in archivos.values():
            todos_archivos.extend(archivos_seccion)
        
        # Generar asunto y cuerpo
        asunto = f"VIVEVERDE: Resumen de Pedidos - Semana {semana}"
        cuerpo = (
            f"Buenos días.\n\n"
            f"Se adjunta el resumen completo de pedidos de compra generados "
            f"para la semana {semana} de todas las secciones.\n\n"
            f"Secciones procesadas: {', '.join(archivos.keys())}\n\n"
            f"Atentamente,\n"
            f"Sistema de Pedidos automáticos VIVEVERDE."
        )
        
        # Crear y enviar mensaje
        msg = self._crear_mensaje([email_central], asunto, cuerpo, todos_archivos)
        exito = self._enviar_email(msg)
        
        return {
            'exito': exito,
            'semana': semana,
            'destinatario': email_central,
            'archivos_enviados': todos_archivos if exito else []
        }
    
    def verificar_configuracion(self) -> Dict[str, Any]:
        """
        Verifica que la configuración del email sea correcta.
        
        Returns:
            Dict: Resultado de la verificación
        """
        resultado = {
            'valido': True,
            'problemas': []
        }
        
        # Verificar configuración SMTP
        if not self.smtp_config.get('servidor'):
            resultado['valido'] = False
            resultado['problemas'].append('Falta servidor SMTP')
        
        if not self.smtp_config.get('puerto'):
            resultado['valido'] = False
            resultado['problemas'].append('Falta puerto SMTP')
        
        # Verificar remitente
        if not self.remitente.get('email'):
            resultado['valido'] = False
            resultado['problemas'].append('Falta email del remitente')
        
        # Verificar destinatarios
        if not self.destinatarios:
            resultado['valido'] = False
            resultado['problemas'].append('No hay destinatarios configurados')
        
        # Verificar variable de entorno
        try:
            self._obtener_password()
        except ValueError:
            resultado['problemas'].append(
                "La variable de entorno 'EMAIL_PASSWORD' no está configurada"
            )
            # No invalidamos por esto, solo advertimos
        
        if resultado['problemas']:
            logger.warning("Problemas de configuración detectados:")
            for problema in resultado['problemas']:
                logger.warning(f"  - {problema}")
        
        return resultado


# Funciones de utilidad para uso directo
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
        Dict: Resultado de la verificación
    """
    service = EmailService(config)
    return service.verificar_configuracion()


# Ejemplo de uso y testing
if __name__ == "__main__":
    import json
    
    print("EmailService - Módulo de envío de emails")
    print("=" * 50)
    
    # Configurar logging
    logging.basicConfig(level=logging.INFO)
    
    # Ejemplo de configuración
    config_ejemplo = {
        'email': {
            'smtp': {
                'servidor': 'smtp.serviciodecorreo.es',
                'puerto': 465,
                'usar_ssl': True
            },
            'remitente': {
                'email': 'ivan.delgado@viveverde.es',
                'nombre': 'Sistema de Pedidos VIVEVERDE'
            },
            'destinatarios': {
                'maf': 'exterior@viveverde.es',
                'interior': 'interior@viveverde.es',
                'deco_interior': [
                    'decoracion@viveverde.es',
                    'sandra.delgado@viveverde.es'
                ]
            },
            'plantillas': {
                'asunto': 'VIVEVERDE: Pedido de compra - Semana {semana} - {seccion}',
                'cuerpo': 'Buenos días {nombre_encargado}. \nTe adjunto el pedido de compra generado para la semana {semana} de la sección {seccion}.\n\nAtentamente,\nSistema de Pedidos automáticos VIVEVERDE.'
            }
        }
    }
    
    # Crear servicio
    service = crear_email_service(config_ejemplo)
    
    # Verificar configuración
    resultado = service.verificar_configuracion()
    print(f"\nVerificación de configuración:")
    print(f"  Válido: {resultado['valido']}")
    if resultado['problemas']:
        print(f"  Problemas: {resultado['problemas']}")
    
    # Obtener destinatarios de ejemplo
    destinatarios = service.obtener_destinatarios_seccion('maf')
    print(f"\nDestinatarios para 'maf': {destinatarios}")
    
    print("\nEmailService listo para usar.")
    print("\nNota: Asegúrate de configurar la variable de entorno EMAIL_PASSWORD")
    print("      antes de enviar emails realmente.")

"""
ORQUESTADOR PRINCIPAL - Control Diario de Reporte IA SSOMA
==========================================================
Ejecuta los agentes en secuencia:
  1. Busqueda de correos de reporte IA SSOMA
  2. Verificacion de formato del adjunto (Word/PDF con datos empresa y fecha)
  3. Verificacion de subidas a Google Drive
  4. Evaluacion de cumplimiento + notificacion a incumplidores
  5. Generacion y envio de reporte HTML

Uso:
  python main.py                       # Ejecutar todo (revisa reporte de ayer)
  python main.py --fecha 2026-02-27    # Revisar reporte de una fecha especifica
  python main.py --no-notificar        # Sin enviar llamados de atencion
  python main.py --solo-buscar         # Solo buscar y listar
"""
import argparse
import json
import os
import sys
import threading
from datetime import datetime, date, timedelta, timezone

from rich.console import Console
from rich.table import Table
from rich.panel import Panel

from config import REPORT_DIR, REPORT_JSON, REPORT_TXT, OBRAS, MODO_PRUEBA, TEST_EMAIL
from auth_gmail import autenticar_gmail, autenticar_drive, obtener_perfil
from agente_busqueda_reportes import buscar_reportes
from agente_verificador_formato import verificar_formato_reporte
from agente_verificador_drive import verificar_subidas_drive
from agente_cumplimiento import evaluar_cumplimiento, enviar_notificaciones, enviar_notificaciones_adicionales
from enviar_reporte_ssoma import enviar_reporte

# Zona horaria Peru (UTC-5)
PERU_TZ = timezone(timedelta(hours=-5))

console = Console()


def main():
    parser = argparse.ArgumentParser(description="Control Diario de Reporte IA SSOMA")
    parser.add_argument(
        "--fecha", type=str, default=None,
        help="Fecha del reporte a revisar (YYYY-MM-DD). Default: ayer"
    )
    parser.add_argument(
        "--no-notificar", action="store_true",
        help="No enviar correos de llamado de atencion"
    )
    parser.add_argument(
        "--solo-buscar", action="store_true",
        help="Solo buscar y listar correos (sin notificar ni reportar)"
    )
    args = parser.parse_args()

    # Determinar fecha objetivo
    if args.fecha:
        try:
            fecha_objetivo = datetime.strptime(args.fecha, "%Y-%m-%d").date()
        except ValueError:
            console.print(f"[bold red]Formato de fecha invalido: {args.fecha}. Usar YYYY-MM-DD[/bold red]")
            sys.exit(1)
    else:
        hoy = datetime.now(PERU_TZ).date()
        if hoy.weekday() == 0:  # Lunes
            # Revisa sabado Y domingo (captura regularizaciones dominicales)
            fecha_objetivo = hoy - timedelta(days=2)  # Desde sabado
        elif hoy.weekday() == 6:  # Domingo
            fecha_objetivo = hoy - timedelta(days=1)  # Sabado
        else:
            fecha_objetivo = hoy - timedelta(days=1)  # Ayer

    modo_txt = "[bold red][MODO PRUEBA][/bold red] " if MODO_PRUEBA else ""
    console.print(Panel.fit(
        f"{modo_txt}[bold green]CONTROL DIARIO DE REPORTE IA SSOMA[/bold green]\n"
        f"Fecha del reporte: [bold]{fecha_objetivo.strftime('%d/%m/%Y')}[/bold]\n"
        f"Ejecucion: {datetime.now(PERU_TZ).strftime('%d/%m/%Y %H:%M')}",
        border_style="green",
    ))

    if MODO_PRUEBA:
        console.print(f"[bold red]  MODO PRUEBA: Todos los correos iran a {TEST_EMAIL}[/bold red]")

    # === AUTENTICACION ===
    console.print("\n[bold yellow]>>> AUTENTICACION[/bold yellow]")
    try:
        service = autenticar_gmail()
        mi_email = obtener_perfil(service)
        console.print(f"  Gmail conectado como: [green]{mi_email}[/green]")

        drive_service = autenticar_drive()
        console.print(f"  Drive conectado: [green]OK[/green]")
    except Exception as e:
        console.print(f"[bold red]Error de autenticacion: {e}[/bold red]")
        sys.exit(1)

    # === AGENTE 1: BUSQUEDA ===
    console.print("\n[bold yellow]>>> AGENTE 1: BUSQUEDA DE CORREOS DE REPORTE IA SSOMA[/bold yellow]")
    reportes = buscar_reportes(service, fecha_objetivo)

    console.print(f"\n  Obras monitoreadas: [bold]{len(OBRAS)}[/bold]")
    console.print(f"  Correos de reporte encontrados: [bold]{len(reportes)}[/bold]")

    if reportes:
        _mostrar_tabla_reportes(reportes)

    # === AGENTE 2: VERIFICACION DE FORMATO ===
    console.print("\n[bold yellow]>>> AGENTE 2: VERIFICACION DE FORMATO (Word/PDF)[/bold yellow]")

    for i, reporte in enumerate(reportes):
        console.print(f"  [{i+1}/{len(reportes)}] {reporte['obra_nombre']}...", end=" ")

        if not reporte["tiene_adjunto_documento"]:
            reporte["datos_formato"] = {
                "formato_valido": False,
                "tipo_archivo": "sin_adjunto",
                "detalle": "Sin adjunto Word/PDF",
            }
            console.print("[yellow]SIN ADJUNTO[/yellow]")
            continue

        try:
            resultado_datos = [None]
            error_datos = [None]

            def _verificar():
                try:
                    resultado_datos[0] = verificar_formato_reporte(
                        service, reporte["id"], reporte["adjuntos"], fecha_objetivo
                    )
                except Exception as ex:
                    error_datos[0] = ex

            hilo = threading.Thread(target=_verificar)
            hilo.daemon = True
            hilo.start()
            hilo.join(timeout=30)

            if hilo.is_alive():
                reporte["datos_formato"] = {"formato_valido": False, "detalle": "Timeout (>30s)"}
                console.print("[yellow]TIMEOUT[/yellow]")
                continue

            if error_datos[0]:
                raise error_datos[0]

            datos = resultado_datos[0]
            reporte["datos_formato"] = datos or {}

            if datos and datos.get("formato_valido"):
                tipo = datos.get("tipo_archivo", "?").upper()
                fecha_doc = datos.get("fecha_documento", "-")
                correcto = "[green]SI[/green]" if datos.get("fecha_correcta") else f"[yellow]NO ({fecha_doc})[/yellow]"
                console.print(
                    f"[green]OK[/green] "
                    f"(Tipo: {tipo}, Fecha: {fecha_doc}, Correcto: {correcto})"
                )
            else:
                console.print(f"[yellow]PARCIAL[/yellow] ({datos.get('detalle', 'Sin datos')})")

        except Exception as e:
            reporte["datos_formato"] = {"formato_valido": False, "detalle": str(e)}
            console.print(f"[red]ERROR[/red] ({e})")

    if args.solo_buscar:
        _guardar_reporte_parcial(reportes, fecha_objetivo, mi_email)
        console.print("\n[green]Busqueda completada. Ejecuta sin --solo-buscar para el flujo completo.[/green]")
        return

    # === AGENTE 3.5: VERIFICACION GOOGLE DRIVE ===
    console.print("\n[bold yellow]>>> AGENTE 3.5: VERIFICACION DE SUBIDAS A GOOGLE DRIVE[/bold yellow]")

    try:
        resultados_drive = verificar_subidas_drive(drive_service, fecha_objetivo)
        total_drive_ok = sum(1 for v in resultados_drive.values() if v.get("subido"))
        console.print(f"\n  Subidas verificadas: [bold]{total_drive_ok}/{len(resultados_drive)}[/bold] obras con archivos en Drive")
    except Exception as e:
        console.print(f"[red]  Error verificando Drive: {e}[/red]")
        resultados_drive = {}

    # === AGENTE 3: CUMPLIMIENTO Y NOTIFICACION ===
    console.print("\n[bold yellow]>>> AGENTE 3: EVALUACION DE CUMPLIMIENTO[/bold yellow]")

    resultado_cumplimiento = evaluar_cumplimiento(reportes, fecha_objetivo, resultados_drive)

    _mostrar_tabla_cumplimiento(resultado_cumplimiento)

    # Enviar notificaciones a obras incumplidoras
    resultado_notificaciones = []
    resultado_notificaciones_obs = []

    if resultado_cumplimiento["no_enviaron"]:
        if args.no_notificar:
            console.print("\n[yellow]  Notificaciones desactivadas (--no-notificar)[/yellow]")
        else:
            console.print("\n[bold yellow]>>> ENVIANDO NOTIFICACIONES DE LLAMADO DE ATENCION[/bold yellow]")
            resultado_notificaciones = enviar_notificaciones(
                service, resultado_cumplimiento["no_enviaron"], fecha_objetivo, mi_email
            )
    else:
        console.print("\n[green]  Todas las obras cumplieron! No hay notificaciones de NO ENVIO.[/green]")

    # Enviar notificaciones adicionales (formato/Drive)
    if not args.no_notificar:
        console.print("\n[bold yellow]>>> ENVIANDO OBSERVACIONES (FORMATO / DRIVE)[/bold yellow]")
        resultado_notificaciones_obs = enviar_notificaciones_adicionales(
            service, resultado_cumplimiento, fecha_objetivo, mi_email
        )
    else:
        console.print("\n[yellow]  Observaciones de formato/Drive desactivadas (--no-notificar)[/yellow]")

    # === REPORTE ===
    console.print("\n[bold yellow]>>> GENERANDO Y ENVIANDO REPORTE[/bold yellow]")

    enviar_reporte(service, mi_email, resultado_cumplimiento, resultado_notificaciones)

    # Guardar reportes locales
    _guardar_reporte_completo(resultado_cumplimiento, resultado_notificaciones, mi_email)

    # Resumen final
    cum = resultado_cumplimiento
    total_noti = len(resultado_notificaciones) + len(resultado_notificaciones_obs)
    console.print(Panel.fit(
        "[bold green]PROCESO COMPLETADO[/bold green]\n"
        f"{'[MODO PRUEBA] ' if MODO_PRUEBA else ''}"
        f"Fecha del reporte: {cum['fecha_objetivo']}\n"
        f"Cumplieron: {len(cum['cumplieron'])}/{cum['total_obras']}\n"
        f"Fecha incorrecta: {len(cum['tareo_incorrecto'])}\n"
        f"No enviaron: {len(cum['no_enviaron'])}\n"
        f"Notificaciones enviadas: {total_noti} (llamados: {len(resultado_notificaciones)}, observaciones: {len(resultado_notificaciones_obs)})\n"
        f"Reporte guardado en: {REPORT_DIR}",
        border_style="green",
    ))


def _mostrar_tabla_reportes(reportes):
    """Muestra tabla de correos de reporte IA SSOMA encontrados."""
    table = Table(title="CORREOS DE REPORTE IA SSOMA ENCONTRADOS", show_lines=True)
    table.add_column("#", style="cyan", width=4)
    table.add_column("Obra", style="bold white", max_width=20)
    table.add_column("Remitente", style="yellow", max_width=35)
    table.add_column("Fecha Envio", style="white", width=16)
    table.add_column("Asunto", max_width=40)
    table.add_column("Adjunto", style="white", width=8)

    for i, reporte in enumerate(reportes, 1):
        adj_status = "[green]SI[/green]" if reporte["tiene_adjunto_documento"] else "[red]NO[/red]"
        table.add_row(
            str(i),
            reporte["obra_nombre"],
            reporte["de_email"],
            reporte["fecha_envio"][:16] if reporte["fecha_envio"] else "-",
            reporte["asunto"][:40],
            adj_status,
        )

    console.print(table)


def _mostrar_tabla_cumplimiento(resultado):
    """Muestra tabla resumen de cumplimiento."""
    table = Table(title="RESUMEN DE CUMPLIMIENTO - REPORTE IA SSOMA", show_lines=True)
    table.add_column("#", style="cyan", width=4)
    table.add_column("Obra", style="bold white", max_width=20)
    table.add_column("Estado", width=22)
    table.add_column("Fecha Reporte", width=14)
    table.add_column("Tipo Archivo", width=12)
    table.add_column("Formato", width=10)
    table.add_column("G.Drive", width=8)

    i = 0

    for item in resultado["cumplieron"]:
        i += 1
        datos = item.get("datos", {})
        fmt_ok = "[green]Ok[/green]" if item.get("formato_ok") else "[red]Corregir[/red]"
        drv_ok = "[green]Ok[/green]" if item.get("drive_ok") else "[red]Falta[/red]"
        tipo = datos.get("tipo_archivo", "-")
        fecha_doc = datos.get("fecha_documento", "-")
        table.add_row(
            str(i),
            item["obra_nombre"],
            f"[green]{item['estado']}[/green]",
            fecha_doc,
            tipo.upper() if tipo != "-" else "-",
            fmt_ok,
            drv_ok,
        )

    for item in resultado["tareo_incorrecto"]:
        i += 1
        datos = item.get("datos", {})
        fmt_ok = "[green]Ok[/green]" if item.get("formato_ok") else "[red]Corregir[/red]"
        drv_ok = "[green]Ok[/green]" if item.get("drive_ok") else "[red]Falta[/red]"
        tipo = datos.get("tipo_archivo", "-")
        fecha_doc = datos.get("fecha_documento", "-")
        table.add_row(
            str(i),
            item["obra_nombre"],
            f"[yellow]{item['estado']}[/yellow]",
            fecha_doc,
            tipo.upper() if tipo != "-" else "-",
            fmt_ok,
            drv_ok,
        )

    for item in resultado["no_enviaron"]:
        i += 1
        drv_ok = "[green]Ok[/green]" if item.get("drive_ok") else "[red]Falta[/red]"
        table.add_row(
            str(i),
            item["obra_nombre"],
            f"[red]{item['estado']}[/red]",
            "-", "-", "-",
            drv_ok,
        )

    console.print(table)


def _guardar_reporte_completo(resultado_cumplimiento, resultado_notificaciones, mi_email):
    """Guarda los resultados en archivos JSON y TXT."""
    os.makedirs(REPORT_DIR, exist_ok=True)

    def _serializar(item):
        """Convierte un item de cumplimiento a formato serializable."""
        copia = dict(item)
        if "reporte" in copia:
            reporte = dict(copia["reporte"])
            reporte.pop("adjuntos", None)
            copia["reporte"] = reporte
        return copia

    data = {
        "fecha_ejecucion": datetime.now(PERU_TZ).isoformat(),
        "usuario": mi_email,
        "fecha_objetivo": resultado_cumplimiento["fecha_objetivo"],
        "total_obras": resultado_cumplimiento["total_obras"],
        "cumplieron": [_serializar(c) for c in resultado_cumplimiento["cumplieron"]],
        "tareo_incorrecto": [_serializar(c) for c in resultado_cumplimiento["tareo_incorrecto"]],
        "no_enviaron": resultado_cumplimiento["no_enviaron"],
        "notificaciones": resultado_notificaciones,
    }

    with open(REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # Reporte TXT
    with open(REPORT_TXT, "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\n")
        f.write("CONTROL DIARIO DE REPORTE IA SSOMA\n")
        f.write(f"Fecha del reporte: {resultado_cumplimiento['fecha_objetivo']}\n")
        f.write(f"Ejecucion: {datetime.now(PERU_TZ).strftime('%d/%m/%Y %H:%M')}\n")
        f.write(f"Usuario: {mi_email}\n")
        f.write("=" * 70 + "\n\n")

        f.write(f"Total obras monitoreadas: {resultado_cumplimiento['total_obras']}\n")
        f.write(f"Cumplieron: {len(resultado_cumplimiento['cumplieron'])}\n")
        f.write(f"Fecha incorrecta: {len(resultado_cumplimiento['tareo_incorrecto'])}\n")
        f.write(f"No enviaron: {len(resultado_cumplimiento['no_enviaron'])}\n\n")

        f.write("--- OBRAS QUE CUMPLIERON ---\n")
        for item in resultado_cumplimiento["cumplieron"]:
            datos = item.get("datos", {})
            f.write(f"  {item['obra_nombre']} - {item['estado']}\n")
            f.write(f"    Tipo archivo: {datos.get('tipo_archivo', '-')}\n")
            f.write(f"    Fecha reporte: {datos.get('fecha_documento', '-')}\n")
            f.write(f"    Formato valido: {'Si' if datos.get('formato_valido') else 'No'}\n\n")

        if resultado_cumplimiento["tareo_incorrecto"]:
            f.write("--- OBRAS CON FECHA INCORRECTA ---\n")
            for item in resultado_cumplimiento["tareo_incorrecto"]:
                f.write(f"  {item['obra_nombre']} - {item.get('detalle', '')}\n\n")

        if resultado_cumplimiento["no_enviaron"]:
            f.write("--- OBRAS QUE NO ENVIARON ---\n")
            for item in resultado_cumplimiento["no_enviaron"]:
                f.write(f"  {item['obra_nombre']}\n")
                f.write(f"    Emails: {', '.join(item['emails_registrados'])}\n\n")

    console.print(f"\n[dim]Reportes guardados:[/dim]")
    console.print(f"  [dim]JSON: {REPORT_JSON}[/dim]")
    console.print(f"  [dim]TXT:  {REPORT_TXT}[/dim]")


def _guardar_reporte_parcial(reportes, fecha_objetivo, mi_email):
    """Guarda reporte parcial (solo busqueda)."""
    os.makedirs(REPORT_DIR, exist_ok=True)

    data = {
        "fecha_ejecucion": datetime.now(PERU_TZ).isoformat(),
        "usuario": mi_email,
        "fecha_objetivo": fecha_objetivo.strftime("%d/%m/%Y"),
        "modo": "solo-buscar",
        "correos_encontrados": len(reportes),
        "reportes": [
            {
                "obra": r["obra_nombre"],
                "de_email": r["de_email"],
                "asunto": r["asunto"],
                "fecha_envio": r["fecha_envio"],
                "tiene_adjunto": r["tiene_adjunto_documento"],
            }
            for r in reportes
        ],
    }

    with open(REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()

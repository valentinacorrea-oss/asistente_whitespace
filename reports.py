import pandas as pd
import json
import os
import re
from datetime import datetime
from pathlib import Path
from io import BytesIO


class ReportGenerador:
    def __init__(self, csv_path="data/scores_empresa_producto.csv"):
        """Inicializa con datos del CSV"""
        self.df = pd.read_csv(csv_path)
        self._procesar_listas()

    def _procesar_listas(self):
        """Convierte strings con listas a listas reales"""
        for col in ["google_comprados", "pp_ganadas", "pp_perdidas", "pp_abiertas"]:
            if col in self.df.columns:
                self.df[col] = self.df[col].apply(self._lista)

    @staticmethod
    def _lista(v):
        """Convierte string de lista a lista"""
        if isinstance(v, list):
            return v
        try:
            return eval(v) if isinstance(v, str) and v.startswith("[") else []
        except:
            return []

    def generar_excel(self, output_path="reportes/whitespace_por_ejecutivo.xlsx"):
        """Genera Excel con hojas por ejecutivo"""
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Hoja de Resumen General
            self._escribir_resumen(writer)

            # Una hoja por ejecutivo
            ejecutivos = self.df[self.df['owner'].notna() & (self.df['owner'] != '(Sin valor)')]['owner'].unique()
            nombres_hoja = {nombre.casefold() for nombre in writer.sheets}
            for ejecutivo in sorted(ejecutivos):
                nombre_hoja = self._nombre_hoja_valido(ejecutivo)
                base = nombre_hoja
                sufijo = 2
                while nombre_hoja.casefold() in nombres_hoja:
                    agregado = f" ({sufijo})"
                    nombre_hoja = f"{base[:31 - len(agregado)]}{agregado}"
                    sufijo += 1
                nombres_hoja.add(nombre_hoja.casefold())
                self._escribir_hoja_ejecutivo(writer, ejecutivo, nombre_hoja)

        return output_path

    def _escribir_resumen(self, writer):
        """Hoja de resumen general de todos los ejecutivos"""
        ejecutivos_list = []

        for ejecutivo in self.df[self.df['owner'].notna() & (self.df['owner'] != '(Sin valor)')]['owner'].unique():
            datos_ejec = self.df[self.df['owner'] == ejecutivo]

            alta = (datos_ejec['prioridad'] == 'Alta').sum()
            media = (datos_ejec['prioridad'] == 'Media').sum()
            baja = (datos_ejec['prioridad'] == 'Baja').sum()
            total = len(datos_ejec)

            score_prom = datos_ejec['score_final'].mean() if len(datos_ejec) > 0 else 0

            # Top producto
            top_prod = datos_ejec[datos_ejec['producto_1'].notna()]['producto_1'].value_counts()
            top_productos = ", ".join(top_prod.index[:3].tolist()) if len(top_prod) > 0 else "N/A"

            ejecutivos_list.append({
                'Ejecutivo': ejecutivo,
                'Email': self._email_de(ejecutivo),
                'Cuentas': total,
                'Alta': alta,
                'Media': media,
                'Baja': baja,
                'Score Promedio': round(score_prom, 1),
                'Productos Top': top_productos
            })

        resumen_df = pd.DataFrame(ejecutivos_list)
        resumen_df.to_excel(writer, sheet_name='Resumen', index=False)

        # Formatear
        ws = writer.sheets['Resumen']
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 20

    def _escribir_hoja_ejecutivo(self, writer, ejecutivo, sheet_name):
        """Hoja detallada por ejecutivo"""
        datos = self.df[self.df['owner'] == ejecutivo].sort_values('score_final', ascending=False)

        # Preparar datos para Excel
        cols_export = [
            'empresa', 'MACRO', 'pais', 'prioridad', 'score_final', 'score_cuenta',
            'producto_1', 'fit_1', 'estado_1', 'razon_1', 'producto_2',
            'google_comprados', 'pp_ganadas', 'es_cliente', 'deal_abierto',
            'pts_cliente', 'pts_penetracion', 'pts_whitespace', 'pts_presupuesto', 'pts_deal_abierto'
        ]

        export_df = datos[cols_export].copy()

        # Convertir listas a strings
        for col in ['google_comprados', 'pp_ganadas']:
            export_df[col] = export_df[col].apply(lambda x: ', '.join(x) if isinstance(x, list) else x)

        # Renombrar columnas
        renombrar = {
            'empresa': 'Empresa',
            'MACRO': 'Macro-industria',
            'pais': 'País',
            'prioridad': 'Prioridad',
            'score_final': 'Score Final',
            'score_cuenta': 'Score Cuenta',
            'producto_1': 'Producto 1',
            'fit_1': 'Fit',
            'estado_1': 'Estado',
            'razon_1': 'Por Qué',
            'producto_2': 'Producto 2',
            'google_comprados': 'Productos Google',
            'pp_ganadas': 'Producto Propio',
            'es_cliente': 'Cliente',
            'deal_abierto': 'Deal Abierto',
            'pts_cliente': 'Pts Cliente',
            'pts_penetracion': 'Pts Industria',
            'pts_whitespace': 'Pts Whitespace',
            'pts_presupuesto': 'Pts Presupuesto',
            'pts_deal_abierto': 'Pts Deal'
        }

        export_df = export_df.rename(columns=renombrar)

        export_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Formatear
        ws = writer.sheets[sheet_name]
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18
        ws.freeze_panes = 'A2'

    @staticmethod
    def _nombre_hoja_valido(nombre):
        """Convierte nombre a válido para Excel (máx 31 caracteres)"""
        valido = re.sub(r'[\\/:*?\[\]]', '', str(nombre))[:31]
        return valido if valido else "Ejecutivo"

    @staticmethod
    def _email_de(nombre):
        """Genera email del ejecutivo"""
        norm = str(nombre).lower().strip()
        norm = re.sub(r'[^a-z\s]', '', norm)
        norm = re.sub(r'\s+', '.', norm)
        return f"{norm}@servinformacion.com"

    def generar_correo_html(self, ejecutivo):
        """Genera HTML del correo para un ejecutivo (previsualización)"""
        datos = self.df[self.df['owner'] == ejecutivo].sort_values('score_final', ascending=False)

        if len(datos) == 0:
            return "<p>Ejecutivo no encontrado</p>"

        # KPIs
        alta = (datos['prioridad'] == 'Alta').sum()
        media = (datos['prioridad'] == 'Media').sum()
        baja = (datos['prioridad'] == 'Baja').sum()
        total = len(datos)
        score_prom = datos['score_final'].mean()

        # Top 3 Alta prioridad
        top3 = datos[datos['prioridad'] == 'Alta'].head(3)

        top3_html = ""
        for _, row in top3.iterrows():
            top3_html += f"""
            <li style="margin-bottom:10px;">
                <b>{self._esc(row['empresa'])}</b> → {self._esc(row.get('producto_1', 'N/A'))}
                <br><span style="color:#555;">Score: {row.get('score_final', 0)}</span>
            </li>
            """

        if not top3_html:
            top3_html = "<li>No tienes cuentas en prioridad Alta; empieza por las de prioridad Media con mayor score.</li>"

        # Tabla de cuentas
        tabla_html = self._generar_tabla_cuentas(datos)

        # HTML final
        html = f"""
        <div style="font-family:Arial,sans-serif; font-size:14px; color:#333; max-width:900px; margin:0 auto; padding:24px;">
            <div style="background:#e74c3c; padding:20px 24px; border-radius:8px 8px 0 0;">
                <span style="color:white; font-size:16px; font-weight:bold;">Oportunidades de Producto Propio en tu Cartera</span>
            </div>
            <div style="border:1px solid #ddd; border-top:none; border-radius:0 0 8px 8px; padding:28px; background:#fff;">
                <p style="margin-top:0;">Hola <b>{self._esc(ejecutivo.split()[0])}</b>,</p>
                <p>Analizamos tu cartera y encontramos <b>{total} cuenta(s)</b> que hoy compran Google pero aún no tienen nuestros productos propios.</p>

                <table style="border-collapse:collapse; width:100%; margin:16px 0;">
                    <tr>
                        <td style="text-align:center; padding:10px; border:1px solid #eee;">
                            <div style="font-size:20px; font-weight:bold; color:#e74c3c;">{total}</div>
                            <div style="font-size:11px; color:#777;">Cuentas</div>
                        </td>
                        <td style="text-align:center; padding:10px; border:1px solid #eee;">
                            <div style="font-size:20px; font-weight:bold; color:#e74c3c;">{alta}</div>
                            <div style="font-size:11px; color:#777;">Prioridad Alta</div>
                        </td>
                        <td style="text-align:center; padding:10px; border:1px solid #eee;">
                            <div style="font-size:20px; font-weight:bold; color:#e74c3c;">{media}</div>
                            <div style="font-size:11px; color:#777;">Prioridad Media</div>
                        </td>
                        <td style="text-align:center; padding:10px; border:1px solid #eee;">
                            <div style="font-size:20px; font-weight:bold; color:#e74c3c;">{round(score_prom, 1)}</div>
                            <div style="font-size:11px; color:#777;">Score Promedio</div>
                        </td>
                    </tr>
                </table>

                <div style="font-size:13px; background:#f5f5f5; border-left:3px solid #e74c3c; padding:10px 14px;">
                    <b>🚀 Empieza por aquí esta semana:</b>
                    <ul style="margin:8px 0 0 0; padding-left:20px;">{top3_html}</ul>
                </div>

                <h3 style="font-size:15px; margin:28px 0 8px;">Tus cuentas prioritarias</h3>
                {tabla_html}

                <hr style="border:none; border-top:1px solid #eee; margin:24px 0;">
                <p style="margin-bottom:0; font-size:13px; color:#555;">
                    <strong>Equipo Sales Excellence</strong><br>
                    <span style="color:#e74c3c; font-weight:bold;">Servinformación</span>
                </p>
            </div>
        </div>
        """

        return html

    def _generar_tabla_cuentas(self, datos):
        """Genera tabla HTML de cuentas"""
        html = """
        <table style="border-collapse:collapse; width:100%; font-size:12px;">
            <thead>
                <tr style="background:#f2f2f2;">
                    <th style="padding:10px; border:1px solid #ddd; text-align:left;">Cuenta</th>
                    <th style="padding:10px; border:1px solid #ddd; text-align:center;">Prioridad</th>
                    <th style="padding:10px; border:1px solid #ddd; text-align:left;">Producto Propio</th>
                    <th style="padding:10px; border:1px solid #ddd; text-align:left;">Compra Google</th>
                </tr>
            </thead>
            <tbody>
        """

        for i, (_, row) in enumerate(datos.head(20).iterrows()):
            bg = "#fff" if i % 2 else "#f9f9f9"
            color_prio = {
                'Alta': '#c0392b',
                'Media': '#e67e22',
                'Baja': '#7f8c8d'
            }.get(row.get('prioridad', ''), '#333')

            html += f"""
            <tr style="background:{bg};">
                <td style="padding:8px 10px; border:1px solid #ddd;">
                    <b>{self._esc(row.get('empresa', ''))}</b><br>
                    <span style="color:#888; font-size:11px;">{self._esc(row.get('MACRO', ''))}</span>
                </td>
                <td style="padding:8px 10px; border:1px solid #ddd; text-align:center; color:{color_prio}; font-weight:bold;">
                    {self._esc(row.get('prioridad', ''))}<br>
                    <span style="color:#888; font-weight:normal; font-size:11px;">{row.get('score_final', 0)}</span>
                </td>
                <td style="padding:8px 10px; border:1px solid #ddd;">
                    {self._esc(row.get('producto_1', '—'))}
                </td>
                <td style="padding:8px 10px; border:1px solid #ddd;">
                    {', '.join(row['google_comprados']) if isinstance(row.get('google_comprados'), list) and row['google_comprados'] else 'Sin dato'}
                </td>
            </tr>
            """

        html += """
            </tbody>
        </table>
        """

        return html

    @staticmethod
    def _esc(t):
        """Escapa caracteres para HTML"""
        return str(t or "").replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def generar_reportes():
    """Función helper para generar reportes"""
    gen = ReportGenerador()
    path = gen.generar_excel()
    return path


if __name__ == "__main__":
    gen = ReportGenerador()
    path = gen.generar_excel()
    print(f" Excel generado: {path}")
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
import urllib.parse
import requests
import io
import plotly.graph_objects as go
import unicodedata
import re

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Supervisión SEIN - Reservas", layout="wide")
st.title("⚡ Dashboard de Supervisión Operativa COES")
st.markdown("Seguimiento Dinámico de Reservas del SEIN (Soporte Dual para Formatos Históricos y Actuales)")

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Setiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def limpiar_texto_extremo(texto):
    """Limpia el texto eliminando espacios, saltos de línea y tildes."""
    if pd.isna(texto) or str(texto).strip().lower() == 'nan': return ""
    t = str(texto).upper()
    t = unicodedata.normalize('NFKD', t).encode('ASCII', 'ignore').decode('utf-8')
    return re.sub(r'[^A-Z0-9]', '', t)

def generar_urls_coes(fecha):
    """Genera las rutas de descarga vinculando cada formato a su hoja de cálculo correspondiente."""
    año = fecha.strftime("%Y")
    mes_num = fecha.strftime("%m")
    dia = fecha.strftime("%d")
    mes_titulo = MESES[fecha.month]
    fecha_str = fecha.strftime("%d%m")
    
    path_nuevo = f"Post Operación/Reportes/IEOD/{año}/{mes_num}_{mes_titulo}/{dia}/AnexoA_{fecha_str}.xlsx"
    path_legacy = f"Post Operación/Reportes/IEOD/{año}/{mes_num}_{mes_titulo}/{dia}/Anexo1_Resumen_{fecha_str}.xlsx"
    
    return [
        (f"https://www.coes.org.pe/portal/browser/download?url={urllib.parse.quote(path_nuevo)}", "RESERVA_FRÍA"),
        (f"https://www.coes.org.pe/portal/browser/download?url={urllib.parse.quote(path_legacy)}", "DESPACHO_EJECUTADO")
    ]

# --- 2. EXTRACCIÓN Y LIMPIEZA DE DATOS (ETL) ---
@st.cache_data(show_spinner=False)
def extraer_datos_reserva_dinamico(fecha):
    urls = generar_urls_coes(fecha)
    headers = {'User-Agent': 'Mozilla/5.0'}
    df_raw = None
    
    for url, hoja_esperada in urls:
        try:
            res = requests.get(url, headers=headers, timeout=20)
            if res.status_code == 200:
                archivo_excel = io.BytesIO(res.content)
                xls = pd.ExcelFile(archivo_excel, engine='openpyxl')
                hojas = xls.sheet_names
                
                hoja_objetivo = hoja_esperada if hoja_esperada in hojas else hojas[0]
                df_raw = pd.read_excel(xls, sheet_name=hoja_objetivo, header=None, nrows=60)
                break
        except Exception:
            continue
            
    if df_raw is None:
        return None, [f"Error de extracción para {fecha.strftime('%d/%m/%Y')}."]

    idx_termica, idx_mant, idx_efi = None, None, None

    # =========================================================================
    # ESCENARIO 1: FORMATO ACTUAL (Búsqueda por Códigos en Fila 5)
    # =========================================================================
    fila_etiquetas_num = df_raw.iloc[4].values
    
    idx_t = np.where(fila_etiquetas_num == 7000)[0]
    if len(idx_t) > 0: idx_termica = idx_t[0]
    
    idx_m = np.where(fila_etiquetas_num == 7002)[0]
    if len(idx_m) > 0: idx_mant = idx_m[0]
    
    idx_e = np.where(fila_etiquetas_num == 1205)[0]
    if len(idx_e) > 0: idx_efi = idx_e[0]

    # =========================================================================
    # ESCENARIO 2: FORMATO HISTÓRICO (Búsqueda semántica)
    # =========================================================================
    if idx_termica is None and idx_mant is None:
        for col_idx in range(df_raw.shape[1]):
            celda_5 = limpiar_texto_extremo(df_raw.iloc[4, col_idx])
            celda_6 = limpiar_texto_extremo(df_raw.iloc[5, col_idx])
            cab = celda_5 + celda_6
            
            if not cab: continue
            
            if "RESERVA" in cab and "TERMOELECTRICA" in cab and "MANTENIMINETO" not in cab:
                if idx_termica is None: idx_termica = col_idx
            elif "MANTENIMINETO" in cab:
                if idx_mant is None: idx_mant = col_idx
            elif "EFICIENTE" in cab:
                if idx_efi is None: idx_efi = col_idx

    mensajes_alerta = []
    if idx_termica is None: mensajes_alerta.append(f"[{fecha.strftime('%d/%m')}] No se halló Reserva Fría.")
    if idx_mant is None: mensajes_alerta.append(f"[{fecha.strftime('%d/%m')}] No se halló Mantenimiento.")
    
    # --- CONSTRUCCIÓN DEL DATAFRAME ---
    df_datos = pd.DataFrame()
    df_datos['Fecha_Hora'] = [datetime.combine(fecha, time(0, 30)) + timedelta(minutes=30*i) for i in range(48)]
    
    def procesar_columna_numerica(df, idx):
        serie = df.iloc[6:54, idx]
        serie_str = serie.astype(str).str.replace(',', '', regex=False).str.strip()
        serie_str = serie_str.replace(['nan', 'NAN', '', 'None'], np.nan)
        return pd.to_numeric(serie_str, errors='coerce').values

    if idx_termica is not None:
        df_datos['Reserva Fría'] = procesar_columna_numerica(df_raw, idx_termica)
    if idx_efi is not None:
        df_datos['Reserva Eficiente'] = procesar_columna_numerica(df_raw, idx_efi)
    if idx_mant is not None:
        df_datos['Mantenimiento'] = procesar_columna_numerica(df_raw, idx_mant)
        
    return df_datos, mensajes_alerta

def procesar_rango_fechas(start_date, end_date, progress_bar, status_text):
    fechas = pd.date_range(start_date, end_date)
    total_dias = len(fechas)
    lista_dfs = []
    alertas_globales = []
    
    for i, f in enumerate(fechas):
        status_text.markdown(f"**⏳ Descargando y procesando IEOD:** {f.strftime('%d/%m/%Y')} *(Día {i+1} de {total_dias})*")
        
        df_dia, alertas = extraer_datos_reserva_dinamico(f)
        if df_dia is not None and not df_dia.empty:
            lista_dfs.append(df_dia)
        if alertas:
            alertas_globales.extend(alertas)
            
        progress_bar.progress((i + 1) / total_dias)
            
    if lista_dfs:
        df_consolidado = pd.concat(lista_dfs, ignore_index=True)
        # Forzamos la existencia de las 3 columnas para que la interfaz nunca se rompa ni desaparezcan los KPIs
        for col in ['Reserva Fría', 'Reserva Eficiente', 'Mantenimiento']:
            if col not in df_consolidado.columns:
                df_consolidado[col] = np.nan
        return df_consolidado, alertas_globales
    return None, alertas_globales

# --- FUNCIONES DE GRÁFICA Y MÉTRICAS ---
def trazar_tendencia_con_estadisticas(fig, df, columna, color_hex):
    # Se añade la traza siempre. Si la data es toda NaN, plotly dibuja el cuadro de leyenda y una gráfica vacía.
    fig.add_trace(go.Scatter(x=df['Fecha_Hora'], y=df[columna], mode='lines', name=columna, line=dict(color=color_hex, width=2), connectgaps=False))
    
    # Solo calculamos promedios y máximos si la columna tiene al menos un dato válido
    if not df[columna].isna().all():
        promedio = df[columna].mean()
        fig.add_hline(y=promedio, line_dash="dash", line_color=color_hex, 
                      annotation_text=f"Promedio: {promedio:.1f}", annotation_position="top left")
        
        max_idx = df[columna].idxmax()
        if pd.notna(max_idx):
            max_val = df.loc[max_idx, columna]
            max_time = df.loc[max_idx, 'Fecha_Hora']
            fig.add_trace(go.Scatter(x=[max_time], y=[max_val], mode='markers+text', 
                                     marker=dict(color=color_hex, size=12, symbol='triangle-up', line=dict(color='black', width=1)), 
                                     text=[f"Máx: {max_val:.1f}"], textposition="top center", showlegend=False))
        
        min_idx = df[columna].idxmin()
        if pd.notna(min_idx):
            min_val = df.loc[min_idx, columna]
            min_time = df.loc[min_idx, 'Fecha_Hora']
            fig.add_trace(go.Scatter(x=[min_time], y=[min_val], mode='markers+text', 
                                     marker=dict(color=color_hex, size=12, symbol='triangle-down', line=dict(color='black', width=1)), 
                                     text=[f"Mín: {min_val:.1f}"], textposition="bottom center", showlegend=False))

def generar_fila_kpis(df, columna, color_alerta_inversa=False):
    st.markdown(f"**🔹 {columna}**")
    
    if df[columna].isna().all():
        promedio, max_val, min_val = np.nan, np.nan, np.nan
        hora_max, hora_min = "-", "-"
    else:
        promedio = df[columna].mean()
        max_idx = df[columna].idxmax()
        max_val = df.loc[max_idx, columna] if pd.notna(max_idx) else np.nan
        hora_max = df.loc[max_idx, "Fecha_Hora"].strftime("%d/%m/%Y %H:%M") if pd.notna(max_idx) else "-"
        
        min_idx = df[columna].idxmin()
        min_val = df.loc[min_idx, columna] if pd.notna(min_idx) else np.nan
        hora_min = df.loc[min_idx, "Fecha_Hora"].strftime("%d/%m/%Y %H:%M") if pd.notna(min_idx) else "-"
    
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Promedio en el Rango", f"{promedio:.1f} MW" if pd.notna(promedio) else "NaN")
    
    color_max = "inverse" if color_alerta_inversa else "normal"
    color_min = "normal" if color_alerta_inversa else "inverse"
    
    kpi2.metric("Pico MÁXIMO", f"{max_val:.1f} MW" if pd.notna(max_val) else "NaN", f"↑ Fecha y Hora: {hora_max}", delta_color=color_max)
    kpi3.metric("Valle MÍNIMO", f"{min_val:.1f} MW" if pd.notna(min_val) else "NaN", f"↓ Fecha y Hora: {hora_min}", delta_color=color_min)
    st.markdown("---")

# --- 3. INTERFAZ Y VISUALIZACIÓN ---
st.sidebar.header("Parámetros de Supervisión")
hoy = datetime.now()
rango_fechas = st.sidebar.date_input("Intervalo de Fechas (IEOD)", value=(datetime(2022, 9, 23), datetime(2022, 9, 25)))

if st.sidebar.button("Procesar Rango COES", key="btn_procesar_rango_historico"):
    if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
        start_date, end_date = rango_fechas
        
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        df, alertas = procesar_rango_fechas(start_date, end_date, progress_bar, status_text)
        
        status_text.empty()
        progress_bar.empty()
        
        if alertas:
            with st.expander("⚠️ Ver Bitácora de Alertas de Extracción"):
                for alerta in alertas:
                    st.warning(alerta)
                    
        if df is not None and not df.empty:
            st.success("✅ Extracción algorítmica y procesamiento completados.")
            
            # --- 3.1 RESUMEN EJECUTIVO (Con Valores Máximos y Mínimos) ---
            st.markdown("### 📑 Resumen Ejecutivo")
            
            texto_resumen = (
                f"**Supervisión de reservas ({start_date.strftime('%d/%m/%Y')} al {end_date.strftime('%d/%m/%Y')})**: "
                "Se ha validado la información de los anexos diarios publicados por el COES. El dashboard a continuación consolida "
                "la evolución temporal y los indicadores máximos/mínimos reales de las reservas del SEIN. "
                "*Nota Analítica: Las interrupciones en los gráficos y los valores 'NaN' representan intervalos donde el COES "
                "no reportó información de reserva en sus formatos oficiales, diferenciándolos estrictamente de un reporte de 0 MW. "
                "Cabe precisar que la métrica de 'Mantenimiento' se refiere específicamente a la Reserva Térmica en Mantenimiento.*\n\n"
                "**Valores Extremos Detectados en el Periodo:**\n"
            )
            
            if not df['Reserva Fría'].isna().all():
                texto_resumen += f"- **Reserva Fría**: Máximo {df['Reserva Fría'].max():.1f} MW | Mínimo {df['Reserva Fría'].min():.1f} MW\n"
            if not df['Reserva Eficiente'].isna().all():
                texto_resumen += f"- **Reserva Eficiente**: Máximo {df['Reserva Eficiente'].max():.1f} MW | Mínimo {df['Reserva Eficiente'].min():.1f} MW\n"
            if not df['Mantenimiento'].isna().all():
                texto_resumen += f"- **Mantenimiento**: Máximo {df['Mantenimiento'].max():.1f} MW | Mínimo {df['Mantenimiento'].min():.1f} MW\n"
                
            st.info(texto_resumen)
            
            # --- ALERTA DE DÍAS SIN RESERVA EFICIENTE ---
            dias_sin_eficiente = []
            for date_str, group in df.groupby(df['Fecha_Hora'].dt.strftime('%d/%m/%Y')):
                if group['Reserva Eficiente'].isna().all():
                    dias_sin_eficiente.append(date_str)
                        
            if len(dias_sin_eficiente) > 0:
                st.error("🔴 **ATENCIÓN: Se detectaron vacíos de información en la matriz extraída.**")
                with st.expander("🚨 Alerta de Información no encontrada: Días sin reporte de Reserva Eficiente", expanded=False):
                    for dia in dias_sin_eficiente:
                        st.warning(f"El COES no reportó información de Reserva Eficiente el día: {dia}")
            
            # --- 3.2 KPIs Críticos Dinámicos ---
            st.markdown("### 📊 Indicadores de Potencia de Reserva")
            
            generar_fila_kpis(df, 'Reserva Fría', color_alerta_inversa=False)
            generar_fila_kpis(df, 'Reserva Eficiente', color_alerta_inversa=False)
            generar_fila_kpis(df, 'Mantenimiento', color_alerta_inversa=True)
            
            # --- 3.3 Gráfica de Tendencias COMBINADA ---
            st.markdown("### 📈 Tendencias Operativas y Desviaciones en el Intervalo")
            
            fig = go.Figure()
            trazar_tendencia_con_estadisticas(fig, df, 'Reserva Fría', '#1f77b4')
            trazar_tendencia_con_estadisticas(fig, df, 'Reserva Eficiente', '#2ca02c')
            trazar_tendencia_con_estadisticas(fig, df, 'Mantenimiento', '#d62728')
            
            fig.update_layout(
                xaxis_title="Fecha y Hora de Operación",
                yaxis_title="Potencia (MW)",
                hovermode="x unified",
                height=600,
                legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1)
            )
            st.plotly_chart(fig, use_container_width=True)

            # --- 3.4 Gráfica de Tendencias EXCLUSIVA PARA RESERVA EFICIENTE ---
            st.markdown("### 📈 Tendencia Operativa: Solo Reserva Eficiente")
            
            fig_efi = go.Figure()
            trazar_tendencia_con_estadisticas(fig_efi, df, 'Reserva Eficiente', '#2ca02c')
            
            fig_efi.update_layout(
                xaxis_title="Fecha y Hora de Operación",
                yaxis_title="Potencia (MW)",
                hovermode="x unified",
                height=500,
                legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1)
            )
            st.plotly_chart(fig_efi, use_container_width=True)
            
            # --- 3.5 Auditoría de Rango ---
            st.markdown("### 🗄️ Trazabilidad de Datos")
            df_mostrar = df.copy()
            df_mostrar['Fecha_Hora'] = df_mostrar['Fecha_Hora'].dt.strftime('%d/%m/%Y %H:%M')
            st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
            
    else:
        st.error("Por favor, seleccione un intervalo válido con Fecha de Inicio y Fecha de Fin.")
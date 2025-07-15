import streamlit as st
from bs4 import BeautifulSoup
import openpyxl
import io

st.title("Convertidor de HTML a Excel")

uploaded_file = st.file_uploader("Sube tu archivo HTML", type="html")

if uploaded_file:
    soup = BeautifulSoup(uploaded_file, 'html.parser')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Datos"
    headers = ['Cliente', 'Riesgo', 'Monto total', 'Moneda', 'Porcentaje',
               'Tipo de inversión', 'Retorno anualizado', 'Cierre de subasta', 'Pago estimado']
    ws.append(headers)

    rows = soup.select('.grid-table-row')
    for row in rows:
        cells = []
        cliente = row.select_one('[data-name="Cliente"]')
        nombre = cliente.select_one('.title').get_text(strip=True) if cliente else ''
        cells.append(nombre)
        riesgo = row.select_one('[data-name="Riesgo"] span')
        cells.append(riesgo.get_text(strip=True) if riesgo else '')
        monto = row.select_one('[data-name="Monto total"] p')
        cells.append(monto.get_text(strip=True) if monto else '')
        moneda = row.select_one('[data-name="Monto total"] .currency')
        cells.append(moneda.get_text(strip=True) if moneda else '')
        porcentaje = row.select_one('.percentage-number')
        cells.append(porcentaje.get_text(strip=True) if porcentaje else '')
        tipo_inversion = row.select_one('[data-name="Tipo de inversión"] span')
        cells.append(tipo_inversion.get_text(strip=True) if tipo_inversion else '')
        retorno = row.select_one('[data-name="Retorno anualizado"]')
        cells.append(retorno.get_text(strip=True) if retorno else '')
        cierre = row.select_one('[data-name="Cierre de subasta"] .title')
        cells.append(cierre.get_text(strip=True) if cierre else '')
        pago = row.select_one('[data-name="Pago estimado"] .title')
        cells.append(pago.get_text(strip=True) if pago else '')
        ws.append(cells)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Descargar Excel",
        data=output,
        file_name="salida.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

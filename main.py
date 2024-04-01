import pandas as pd

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from copy import copy
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.styles import Border, Side
from openpyxl.drawing.xdr import XDRPositiveSize2D, XDRPoint2D
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor

import streamlit as st

def main():
    st.title('Upload e Visualização de Planilha XLSX')

    # Cria um uploader para o usuário subir o arquivo
    uploaded_file = st.file_uploader("Escolha uma planilha XLSX", type="xlsx")
    if uploaded_file is not None:
        # Usando Pandas para ler o arquivo
        df = pd.read_excel(uploaded_file)

        # Mostrando o dataframe na página
        st.write("Dados da Planilha:")
        st.write(df)

if __name__ == "__main__":
    main()
"""Colores y fuentes de XINGON! para la app (los mismos que la web)."""
import streamlit as st

COLORES = {
    "fondo": "#041539", "tapete": "#0c2459", "linea": "#22407f", "texto": "#f5f7fc", "suave": "#b3bfdc",
    "amarillo": "#eaad02", "azul": "#15428c", "verde": "#226c44", "rojo": "#c01f24",
}

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=ADLaM+Display&family=Nunito+Sans:wght@400;600;700;800&display=swap');

.stApp, .stApp *:not([data-testid="stIconMaterial"]):not(.material-symbols-rounded):not(code):not(pre) {
  font-family: 'Nunito Sans', system-ui, sans-serif !important;
}
.stApp h1, .stApp h2, .stApp h3, .stApp h1 *, .stApp h2 *, .stApp h3 *,
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h1 * {
  font-family: 'ADLaM Display', 'Nunito Sans', sans-serif !important;
  font-weight: 400 !important; letter-spacing: .01em;
}
.stApp h1, .stApp h2 { color: #f5f7fc; }
.stApp h3 { color: #eaad02; }

[data-testid="stSidebar"] { background: #0c2459; border-right: 2px solid #22407f; }
[data-testid="stSidebar"] [role="radiogroup"] label { padding: 4px 6px; border-radius: 10px; }
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) { background: #22407f; }

/* tarjetas de métricas con los 4 colores de las cartas */
[data-testid="stMetric"] {
  background: #0c2459; border: 2px solid #22407f; border-radius: 16px; padding: 12px 16px;
  border-top: 6px solid #eaad02;
}
[data-testid="stColumn"]:nth-child(4n+2) [data-testid="stMetric"] { border-top-color: #15428c; }
[data-testid="stColumn"]:nth-child(4n+3) [data-testid="stMetric"] { border-top-color: #226c44; }
[data-testid="stColumn"]:nth-child(4n+4) [data-testid="stMetric"] { border-top-color: #c01f24; }
[data-testid="stMetricLabel"] { color: #b3bfdc; }
[data-testid="stMetricValue"] { font-weight: 800; }

/* botones: texto blanco con contorno */
.stButton button, .stDownloadButton button, .stFormSubmitButton button,
[data-testid="stFileUploader"] button, [data-testid="stBaseButton-secondary"],
[data-testid="stBaseButton-primary"], [data-testid="stBaseButton-secondaryFormSubmit"],
[data-testid="stBaseButton-primaryFormSubmit"] {
  background: #0c2459 !important; color: #f5f7fc !important;
  border: 2px solid #b3bfdc !important; border-radius: 999px !important; font-weight: 700 !important;
}
.stButton button *, .stDownloadButton button *, .stFormSubmitButton button *,
[data-testid="stFileUploader"] button * { color: #f5f7fc !important; }
.stButton button[kind^="primary"], .stDownloadButton button[kind^="primary"],
.stFormSubmitButton button[kind^="primary"], [data-testid^="stBaseButton-primary"] {
  border-color: #eaad02 !important;
}
.stButton button:hover, .stDownloadButton button:hover, .stFormSubmitButton button:hover,
[data-testid="stFileUploader"] button:hover {
  background: #22407f !important; border-color: #eaad02 !important; color: #f5f7fc !important;
}
.stButton button:disabled, .stFormSubmitButton button:disabled {
  opacity: .55; border-color: #22407f !important;
}

/* pestañas, formularios y desplegables */
.stTabs [aria-selected="true"] { color: #eaad02 !important; }
[data-testid="stForm"], [data-testid="stExpander"] details {
  border: 2px solid #22407f; border-radius: 16px; background: rgba(12, 36, 89, .45);
}
[data-testid="stDataFrame"], [data-testid="stDataEditor"] { border: 2px solid #22407f; border-radius: 12px; }
hr { border-color: #22407f !important; }
</style>
"""


def aplicar():
    st.markdown(CSS, unsafe_allow_html=True)

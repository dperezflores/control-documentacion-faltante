from __future__ import annotations


def get_custom_css() -> str:
    return """
<style>
    :root {
        --inst-orange: #FF5E12;
        --inst-orange-2: #FF7D42;
        --inst-charcoal: #362D32;
        --inst-peach: #FFBAA3;
        --inst-navy: #00304F;
        --inst-gray: #D6D6D6;
        --inst-bg: #F7F8FA;
        --inst-white: #FFFFFF;
    }

    .stApp {
        background: var(--inst-bg);
        color: var(--inst-charcoal);
        color-scheme: light;
    }

    .block-container {
        max-width: 1280px;
        padding-top: 1.6rem;
        padding-bottom: 3rem;
    }

    /* Encabezado institucional */
    .institutional-header {
        background: var(--inst-white);
        border: 1px solid #E7E8EB;
        border-top: 5px solid var(--inst-orange);
        border-radius: 14px;
        padding: 22px 26px 20px 26px;
        margin: 0 0 24px 0;
        box-shadow: 0 5px 18px rgba(0, 48, 79, .06);
    }
    .institutional-eyebrow {
        color: var(--inst-orange);
        font-size: .76rem;
        font-weight: 800;
        letter-spacing: .12em;
        text-transform: uppercase;
        margin-bottom: 5px;
    }
    .institutional-title {
        color: var(--inst-navy);
        font-size: 2rem;
        font-weight: 800;
        letter-spacing: -.025em;
        line-height: 1.15;
        margin: 0;
    }
    .institutional-subtitle {
        color: #66727A;
        font-size: .94rem;
        margin-top: 7px;
    }

    h1, h2, h3, h4 {
        color: var(--inst-navy);
        letter-spacing: -.015em;
    }
    h2 {
        margin-top: 1.35rem !important;
    }
    p, label, span {
        color: inherit;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: var(--inst-navy);
        border-right: 0;
    }
    [data-testid="stSidebar"] > div {
        background: var(--inst-navy);
    }
    [data-testid="stSidebar"] * {
        color: #FFFFFF;
    }
    [data-testid="stSidebar"] hr {
        border-color: rgba(255,255,255,.18);
    }
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
        color: rgba(255,255,255,.76);
    }
    [data-testid="stSidebar"] div[role="radiogroup"] {
        gap: 6px;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label {
        border-radius: 9px;
        padding: 8px 10px;
        transition: background .15s ease;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background: rgba(255,255,255,.10);
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {
        background: var(--inst-orange);
    }
    [data-testid="stSidebar"] .stButton > button {
        border: 1px solid #FFBAA3;
        background: #FFFFFF;
        color: #00304F !important;
    }
    [data-testid="stSidebar"] .stButton > button * {
        color: #00304F !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background: #FFF3EE;
        border-color: #FF7D42;
        color: #FF5E12 !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover * {
        color: #FF5E12 !important;
    }

        /* Mantener intactos los controles nativos de la barra lateral.
       Sólo ocultamos el menú principal de Streamlit. */
    #MainMenu {
        display: none !important;
    }

    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"] {
        position: fixed !important;
        top: .65rem !important;
        left: .65rem !important;
        color: #00304F !important;
    }

    /* Controles */
    [data-baseweb="select"] > div,
    .stTextInput input,
    .stDateInput input,
    .stTextArea textarea {
        background: var(--inst-white) !important;
        border-color: #DDE2E6 !important;
        border-radius: 10px !important;
        color: var(--inst-charcoal) !important;
        min-height: 42px;
    }
    [data-baseweb="select"]:focus-within > div,
    .stTextInput:focus-within input,
    .stDateInput:focus-within input,
    .stTextArea:focus-within textarea {
        border-color: var(--inst-orange) !important;
        box-shadow: 0 0 0 2px rgba(255,94,18,.10) !important;
    }


    /* Forzar tema claro institucional aunque el sistema/navegador esté en modo oscuro */
    [data-testid="stMain"] [data-testid="stWidgetLabel"],
    [data-testid="stMain"] [data-testid="stWidgetLabel"] p,
    [data-testid="stMain"] [data-testid="stWidgetLabel"] span,
    [data-testid="stMain"] .stTextInput label,
    [data-testid="stMain"] .stTextInput label p,
    [data-testid="stMain"] .stTextArea label,
    [data-testid="stMain"] .stTextArea label p,
    [data-testid="stMain"] .stMultiSelect label,
    [data-testid="stMain"] .stMultiSelect label p,
    [data-testid="stMain"] .stSelectbox label,
    [data-testid="stMain"] .stSelectbox label p,
    [data-testid="stMain"] .stDateInput label,
    [data-testid="stMain"] .stDateInput label p {
        color: var(--inst-charcoal) !important;
    }

    [data-testid="stMain"] [data-testid="stTooltipIcon"],
    [data-testid="stMain"] [data-testid="stTooltipIcon"] *,
    [data-testid="stMain"] [data-testid="stWidgetLabel"] svg {
        color: var(--inst-charcoal) !important;
        fill: currentColor !important;
    }

    [data-testid="stMain"] .stTextInput input::placeholder,
    [data-testid="stMain"] .stDateInput input::placeholder,
    [data-testid="stMain"] .stTextArea textarea::placeholder,
    [data-testid="stMain"] [data-baseweb="select"] input::placeholder {
        color: #8A929A !important;
        opacity: 1 !important;
    }

    [data-testid="stMain"] [data-testid="stCaptionContainer"],
    [data-testid="stMain"] [data-testid="stCaptionContainer"] p,
    [data-testid="stMain"] [data-testid="stCaptionContainer"] span {
        color: #7A8288 !important;
    }

    [data-testid="stMain"] h1,
    [data-testid="stMain"] h2,
    [data-testid="stMain"] h3,
    [data-testid="stMain"] h4 {
        color: var(--inst-navy) !important;
    }

    /* Botones */
    .stButton > button,
    .stDownloadButton > button {
        border-radius: 9px;
        font-weight: 700;
        white-space: nowrap;
        min-height: 40px;
        transition: all .15s ease;
    }
    .stButton > button:hover,
    .stDownloadButton > button:hover {
        transform: translateY(-1px);
    }
    .stButton > button[kind="primary"],
    .stDownloadButton > button[kind="primary"] {
        background: var(--inst-orange);
        border-color: var(--inst-orange);
        color: #FFFFFF;
        box-shadow: 0 4px 10px rgba(255,94,18,.18);
    }
    .stButton > button[kind="primary"]:hover,
    .stDownloadButton > button[kind="primary"]:hover {
        background: #E94F08;
        border-color: #E94F08;
    }
    .stButton > button:not([kind="primary"]) {
        background: var(--inst-white);
        border-color: #CBD2D8;
        color: var(--inst-navy);
    }
    .stButton > button:not([kind="primary"]):hover {
        border-color: var(--inst-orange);
        color: var(--inst-orange);
        background: #FFF7F3;
    }

    /* Tarjetas y métricas */
    div[data-testid="stMetric"] {
        background: var(--inst-white);
        border: 1px solid #E2E5E8;
        border-top: 4px solid var(--inst-orange);
        border-radius: 12px;
        padding: 15px 17px;
        box-shadow: 0 4px 14px rgba(0,48,79,.05);
    }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: var(--inst-navy);
        font-weight: 800;
    }
    div[data-testid="stExpander"] {
        background: var(--inst-white);
        border: 1px solid #E1E5E8;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 3px 12px rgba(0,48,79,.04);
    }
    div[data-testid="stExpander"] details summary:hover {
        background: #FFF7F3;
    }

    .info-card {
        background: var(--inst-white);
        border: 1px solid #E2E5E8;
        border-left: 5px solid var(--inst-orange-2);
        border-radius: 12px;
        padding: 18px 21px;
        margin: 8px 0 20px 0;
        box-shadow: 0 4px 14px rgba(0,48,79,.05);
    }
    .info-label {
        color: var(--inst-orange);
        font-size: .72rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: .075em;
    }
    .info-value {
        color: var(--inst-charcoal);
        font-size: .98rem;
        font-weight: 560;
        margin: 3px 0 11px 0;
    }
    .subtle-box {
        background: var(--inst-white);
        border: 1px solid #E2E5E8;
        border-radius: 10px;
        padding: 14px 16px;
        margin-bottom: 12px;
    }

    /* Tablas y separadores */
    [data-testid="stDataFrame"] {
        border: 1px solid #E2E5E8;
        border-radius: 10px;
        overflow: hidden;
        background: var(--inst-white);
    }
    hr {
        border: 0;
        border-top: 1px solid #E0E4E7;
        margin: 1.25rem 0;
    }

    /* Mensajes */
    [data-testid="stAlert"] {
        border-radius: 10px;
    }
    [data-testid="stAlert"][data-baseweb="notification"] {
        border-left-width: 4px;
    }

    /* Etiquetas y captions */
    .section-kicker {
        color: #6B747B;
        font-size: .88rem;
        margin-bottom: 1rem;
    }
    .stCaptionContainer {
        color: #7A8288;
    }

    /* Responsive */
    @media (max-width: 900px) {
        .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
        }
        .institutional-header {
            padding: 18px;
        }
        .institutional-title {
            font-size: 1.65rem;
        }
    }
</style>
"""

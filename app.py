st.markdown("""
    <style>
    /* Ocultar barra lateral y header nativo de Streamlit */
    [data-testid="stSidebar"] {
        display: none !important;
    }
    header, .stApp > header {
        display: none !important;
        visibility: hidden !important;
        height: 0px !important;
    }
    
    /* Habilitar scroll vertical nativo y fluido en toda la página */
    html, body, [data-testid="stAppViewContainer"], .main {
        overflow-y: auto !important;
        overflow-x: hidden !important;
        height: 100% !important;
    }

    .main .block-container {
        padding-top: 0rem !important;
        padding-bottom: 4rem !important;
        padding-left: 0rem !important;
        padding-right: 0rem !important;
        max-width: 100% !important;
        background-color: transparent !important;
    }
    
    /* Posicionamiento fijo de la barra de navegación superior */
    .stTabs {
        position: fixed !important;
        top: 10px !important;
        left: 15px !important;
        z-index: 99999 !important;
        width: auto !important;
        background: transparent !important;
    }
    .stTabs [data-baseweb="tab-list"],
    div[data-baseweb="tab-list"] {
        background: #e3f2fd !important;
        background-color: #e3f2fd !important;
        border-radius: 8px !important;
        padding: 3px 6px !important;
        border-bottom: none !important;
        gap: 6px !important;
        backdrop-filter: blur(8px) !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.06) !important;
    }
    div[data-baseweb="tab-highlight"] {
        background-color: transparent !important;
    }
    .stTabs [data-baseweb="tab"] {
        height: 32px !important;
        background-color: transparent !important;
        border-radius: 6px !important;
        padding: 0px 12px !important;
        font-size: 0.8rem !important;
        font-weight: 600 !important;
        color: #1F4E78 !important;
        border: none !important;
        box-shadow: none !important;
    }
    
    /* Primera pestaña: Logo del MSP grande */
    .stTabs [data-baseweb="tab"]:nth-child(1) {
        width: 55px !important;
        min-width: 55px !important;
        background-image: url('https://images.seeklogo.com/logo-png/23/2/ministerio-de-salud-publica-logo-png_seeklogo-231476.png') !important;
        background-size: contain !important;
        background-repeat: no-repeat !important;
        background-position: center !important;
        color: transparent !important;
        font-size: 0px !important;
        padding: 0 !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1F4E78 !important;
        color: white !important;
    }
    .stTabs [aria-selected="true"]:nth-child(1) {
        background-color: rgba(31, 78, 120, 0.15) !important;
        background-image: url('https://images.seeklogo.com/logo-png/23/2/ministerio-de-salud-publica-logo-png_seeklogo-231476.png') !important;
        background-size: contain !important;
        background-repeat: no-repeat !important;
        background-position: center !important;
    }

    /* Fondo de pantalla estático y protegido al fondo */
    .hero-fullscreen {
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        background: linear-gradient(90deg, #e3f2fd 0%, #f0f7ff 50%, #ffffff 50%);
        display: flex;
        align-items: center;
        justify-content: space-between;
        overflow: hidden;
        margin: 0 !important;
        z-index: 0;
        pointer-events: none;
    }
    .hero-left-content {
        width: 50%;
        padding-left: 6%;
        padding-right: 4%;
        z-index: 1;
        display: flex;
        flex-direction: column;
        justify-content: center;
        transform: translateY(-20px);
    }
    .welcome-text {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        font-size: 1.1rem;
        letter-spacing: 3px;
        color: #0d47a1;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 12px;
    }
    .hero-title {
        font-size: 2.8rem;
        font-weight: 800;
        color: #1F4E78;
        line-height: 1.15;
        margin-bottom: 15px;
    }
    .hero-right-image {
        width: 55%;
        height: 100%;
        position: absolute;
        right: 0;
        top: 0;
        background-image: url('https://www.metroecuador.com.ec/resizer/v2/JPWRHUMSIRGCDIHJTUBWDHGKJU.jpg?auth=5e0819a9b2036e60b75f706117b2939a852dfd313c77b0cc4feb0e1e423b24c0&width=1440&height=752');
        background-size: cover;
        background-position: center;
        -webkit-mask-image: linear-gradient(to right, rgba(0,0,0,0) 0%, rgba(0,0,0,1) 25%, rgba(0,0,0,1) 100%);
        mask-image: linear-gradient(to right, rgba(0,0,0,0) 0%, rgba(0,0,0,1) 25%, rgba(0,0,0,1) 100%);
        z-index: 1;
    }
    
    /* Contenedor de contenido adaptable con flujo libre hacia abajo */
    .content-overlay {
        position: relative;
        margin-top: 65px;
        margin-left: 20px;
        margin-right: 20px;
        margin-bottom: 50px;
        padding: 25px 35px;
        z-index: 10;
        background: rgba(255, 255, 255, 0.88);
        backdrop-filter: blur(10px);
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.9);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.05);
        box-sizing: border-box;
    }
    
    .ai-card {
        background: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 10px;
        padding: 16px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.02);
        height: 100%;
    }
    .metric-card {
        background: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 14px 18px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }

    .stButton>button {
        width: 100%;
        background-color: #1F4E78;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.6rem 1rem;
    }
    .stButton>button:hover {
        background-color: #153450;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

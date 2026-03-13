import streamlit as st

# Configuração geral da página
st.set_page_config(
    page_title="GDL IC - GTA",
    page_icon="🚔",
    layout="wide"
)

# Tela inicial
p_home = st.Page("home.py", title="Início", icon="🏠", default=True)

# --- Definição das Páginas ---
# Veículos
p_vistoria = st.Page("paginas/vistoria.py", title="Vistoria de Veículo", icon="🚗")
p_chassi = st.Page("paginas/chassi.py", title="Exame de Chassi", icon="🔍")

# Eletroeletrônicos
p_celular = st.Page("paginas/celular.py", title="Aparelho Celular", icon="📱")
p_maquininhas = st.Page("paginas/maquininhas.py", title="Maquininhas de Cartão", icon="💳")
p_caca_niquel = st.Page("paginas/caca_niquel.py", title="Caça-Níquel", icon="🎰")

# Drogas
p_entorpecentes = st.Page("paginas/entorpecentes.py", title="Entorpecentes", icon="🌿")
p_apetrechos = st.Page("paginas/apetrechos_drogas.py", title="Apetrechos para Drogas", icon="⚖️")

# Armas e Instrumentos
p_armas = st.Page("paginas/armas.py", title="Armas e Munições", icon="🔫")
p_facas = st.Page("paginas/facas.py", title="Facas", icon="🔪")

# Biológico e Pessoal
p_vestuario = st.Page("paginas/vestuario_biologico.py", title="Vestuário / DNA", icon="🩸")

# Diversos
p_ambiental = st.Page("paginas/ambiental.py", title="Crime Ambiental", icon="🌳")
p_outras = st.Page("paginas/outras_pecas.py", title="Outras Peças", icon="📦")


# --- Configuração do Menu de Navegação Lateral Categorizado ---
pg = st.navigation({
    "Principal": [p_home],
    "Veículos": [p_vistoria, p_chassi],
    "Eletroeletrônicos": [p_celular, p_maquininhas, p_caca_niquel],
    "Drogase Afins": [p_entorpecentes, p_apetrechos],
    "Armas e Instrumentos": [p_armas, p_facas],
    "Biologia": [p_vestuario],
    "Diversos": [p_ambiental, p_outras]
})

# Executa a navegaçã
pg.run()

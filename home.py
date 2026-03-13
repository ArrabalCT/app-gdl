import streamlit as st

st.title("🚔 Gerador de Laudos Periciais")
st.markdown("---")

st.subheader("Selecione o tipo de laudo:")

# --- CATEGORIA: VEÍCULOS ---
st.markdown("#### 🚗 Veículos")
c1, c2, c3 = st.columns(3)
with c1:
    if st.button("🚗 Vistoria de Veículo", use_container_width=True): st.switch_page("paginas/vistoria.py")
with c2:
    if st.button("🔍 Exame de Chassi", use_container_width=True): st.switch_page("paginas/chassi.py")

st.write("") # Espaçament

# --- CATEGORIA: ELETROELETRÔNICOS ---
st.markdown("#### 💻 Eletroeletrônicos")
c1, c2, c3 = st.columns(3)
with c1:
    if st.button("📱 Aparelho Celular", use_container_width=True): st.switch_page("paginas/celular.py")
with c2:
    if st.button("💳 Maquininhas de Cartão", use_container_width=True): st.switch_page("paginas/maquininhas.py")
with c3:
    if st.button("🎰 Caça-Níquel", use_container_width=True): st.switch_page("paginas/caca_niquel.py")

st.write("")

# --- CATEGORIA: DROGAS E AFINS ---
st.markdown("#### 🌿 Drogas e Afins")
c1, c2, c3 = st.columns(3)
with c1:
    if st.button("🌿 Entorpecentes", use_container_width=True): st.switch_page("paginas/entorpecentes.py")
with c2:
    if st.button("⚖️ Apetrechos para Drogas", use_container_width=True): st.switch_page("paginas/apetrechos_drogas.py")

st.write("")

# --- CATEGORIA: ARMAS E INSTRUMENTOS ---
st.markdown("#### 🔫 Armas e Instrumentos")
c1, c2, c3 = st.columns(3)
with c1:
    if st.button("🔫 Armas e Munições", use_container_width=True): st.switch_page("paginas/armas.py")
with c2:
    if st.button("🔪 Facas", use_container_width=True): st.switch_page("paginas/facas.py")

st.write("")

# --- CATEGORIA: BIOLOGIA ---
st.markdown("#### 🩸 Biologia / Local")
c1, c2, c3 = st.columns(3)
with c1:
    if st.button("🩸 Vestuário / Sangue / DNA", use_container_width=True): st.switch_page("paginas/vestuario_biologico.py")

st.write("")

# --- CATEGORIA: DIVERSOS ---
st.markdown("#### 📦 Diversos")
c1, c2, c3 = st.columns(3)
with c1:
    if st.button("🌳 Crime Ambiental", use_container_width=True): st.switch_page("paginas/ambiental.py")
with c2:
    if st.button("📦 Outras Peças", use_container_width=True): st.switch_page("paginas/outras_pecas.py")

st.markdown("---")
st.info("💡 Vistoria, Celular prontas. Armas e munições tbm, mas não testei todos os campos . Use a barra lateral para navegar. Dúvidas e Sugestões no Zap")
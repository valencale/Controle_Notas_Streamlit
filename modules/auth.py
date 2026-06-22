import streamlit as st
import functools

USERS = {
    "admin": {"name": "Administrador", "password": "admin", "role": "ADM"},
    "visita": {"name": "Visitante", "password": "123", "role": "Visitante"}
}

def init_auth():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "user_role" not in st.session_state:
        st.session_state.user_role = None
    if "username" not in st.session_state:
        st.session_state.username = None

def login_screen():
    """Mostra o formulário de login e retorna True se autenticado."""
    if st.session_state.authenticated:
        return True

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            "<div style='text-align:center; font-size:50px;'>🌿</div>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<h1 style='text-align: center; color: #10b981; margin-top:0;'>Green Bags</h1>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<h4 style='text-align: center; opacity:0.7;'>Logística e Entregas</h4>",
            unsafe_allow_html=True,
        )

        with st.form("login_form"):
            username = st.text_input("Usuário")
            password = st.text_input("Senha", type="password")
            submit = st.form_submit_button("Entrar", use_container_width=True)

            if submit:
                if username in USERS and USERS[username]["password"] == password:
                    st.session_state.authenticated = True
                    st.session_state.user_role = USERS[username]["role"]
                    st.session_state.username = USERS[username]["name"]
                    st.rerun()
                else:
                    st.error("Usuário ou senha inválidos.")

        # Botão de acesso rápido como Visitante
        st.markdown(
            "<p style='text-align:center; opacity:0.5; font-size:13px; margin-top:12px;'>— ou —</p>",
            unsafe_allow_html=True,
        )
        if st.button("👤 Entrar como Visitante", use_container_width=True):
            st.session_state.authenticated = True
            st.session_state.user_role = "Visitante"
            st.session_state.username = "Visitante"
            st.rerun()

    return False


def logout():
    """Encerra a sessão do usuário e retorna à tela de login."""
    st.session_state.authenticated = False
    st.session_state.user_role = None
    st.session_state.username = None
    # Limpa o chat da GIA ao sair
    for key in ["gemini_chat", "gemini_client", "chat_history"]:
        st.session_state.pop(key, None)
    st.rerun()


def is_admin():
    """Retorna True se o usuário logado for ADM."""
    return st.session_state.get("user_role") == "ADM"


# Alias para manter compatibilidade com o resto do código
is_adm = is_admin


def require_adm(func):
    """Decorator que bloqueia a execução de funções de escrita para não-administradores."""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        if not is_admin():
            st.warning("⚠️ Acesso restrito: somente administradores podem executar esta ação.")
            return None
        return func(*args, **kwargs)
    return wrapper

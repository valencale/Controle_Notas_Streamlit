import streamlit as st
import streamlit.components.v1 as components
from modules.chatbot import send_message_to_gemini, get_chat_session, init_gemini


def render_floating_chat():
    """
    Renderiza o popover do Chatbot IA.
    Após renderizar, injeta JS para marcar o container do popover
    com data-gb-chat='1', permitindo que o app.py o encontre e reposicione.
    """

    with st.popover("🤖"):
        st.markdown("""
        <h3 style='margin-top:0; color: var(--gb-accent, #10b981);'>
            Olá! Eu sou a GIA 🌿
        </h3>
        <p style='font-size: 14px; margin-bottom: 0;'>
            Sua assistente de Gestão Logística.
            Pergunte-me sobre os pedidos em andamento!
        </p>
        """, unsafe_allow_html=True)
        st.divider()

        # Checa se a chave da API está presente
        if not init_gemini():
            st.warning("⚠️ API Key do Gemini não encontrada no secrets.toml.")
            return

        # Inicializa mensagens no session state
        if "chat_messages" not in st.session_state:
            st.session_state.chat_messages = []

        # Renderiza mensagens anteriores
        for msg in st.session_state.chat_messages:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        # Componente de upload (opcional)
        uploaded_file = st.file_uploader(
            "Anexar documento (PDF, Excel)",
            type=["pdf", "xlsx", "xls", "txt"],
            label_visibility="collapsed"
        )

        # Input do chat
        user_input = st.chat_input("Digite sua pergunta...")

        if user_input:
            st.session_state.chat_messages.append({"role": "user", "content": user_input})
            with st.chat_message("user"):
                st.markdown(user_input)

            with st.chat_message("assistant"):
                with st.spinner("Pensando..."):
                    file_bytes = None
                    mime_type = None

                    if uploaded_file:
                        file_bytes = uploaded_file.getvalue()
                        mime_type = uploaded_file.type

                    response_text = send_message_to_gemini(user_input, file_bytes, mime_type)

                    st.markdown(response_text)
                    st.session_state.chat_messages.append(
                        {"role": "assistant", "content": response_text}
                    )

    # ── Marca o container do popover para que o app.py possa encontrá-lo ──
    components.html("""<script>
    (function(){
        var pd = window.parent.document;
        // Encontra TODOS os popovers e marca o último (que é o chat)
        var pops = pd.querySelectorAll('[data-testid="stPopover"]');
        if(pops.length > 0){
            pops[pops.length - 1].setAttribute('data-gb-chat', '1');
        }
    })();
    </script>""", height=0, width=0)

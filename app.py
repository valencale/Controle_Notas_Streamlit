import streamlit as st
import sys
import os
import streamlit.components.v1 as components

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.ui_components import inject_custom_css
from modules.auth import init_auth, login_screen

# ══════════════════════════════════════════════════════════════
# PAGE CONFIG — Deve ser o primeiro comando Streamlit
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Gestão Logística — Armazém",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

if st.query_params.get("logout") == "1":
    st.query_params.clear()
    from modules.auth import logout
    logout()

inject_custom_css()


# ══════════════════════════════════════════════════════════════
# TEMA + TOOLBAR + CHAT — Tudo injetado via JS no documento pai
# ══════════════════════════════════════════════════════════════
def _inject_ui_enhancements(is_logged_in: bool, username: str, role: str):
    """
    Injeta diretamente no window.parent.document:
    1. <style> com CSS do toolbar + popover de chat (dark/light)
    2. Div do toolbar com toggle switch (barrinha + bolinha) e botões de usuário
    3. Tema via localStorage + classList no <html>
    4. Reposicionamento do botão do chat via setInterval
    """
    is_logged_in_str = 'true' if is_logged_in else 'false'
    
    # Prepara variaveis para o JS (evitando f-string em bloco enorme de JS)
    js_vars = f"""
    var gb_is_logged_in = '{is_logged_in_str}';
    var gb_username = '{username}';
    var gb_role = '{role}';
    """
    
    js_script = """<script>""" + js_vars + """
(function(){
    var pd = window.parent.document;
    var ls = window.parent.localStorage;

    /* Helper: is current theme light? */
    function isLightTheme(){ return pd.documentElement.classList.contains('gb-light'); }
    function txtColor(){ return isLightTheme() ? '#064e3b' : '#f0fdf4'; }

    /* ═══ 1. INJECT / UPDATE STYLES (always overwrite) ═══ */
    var s = pd.getElementById('gb-ui-styles');
    if(!s){ s = pd.createElement('style'); s.id='gb-ui-styles'; pd.head.appendChild(s); }
    s.textContent = [
        /* Toolbar container – background so text is always visible */
        '#gb-toolbar{position:fixed;top:6px;right:76px;z-index:999990;display:flex;align-items:center;gap:6px;font-family:Inter,-apple-system,sans-serif;background:rgba(12,26,20,.85);backdrop-filter:blur(8px);padding:4px 12px;border-radius:10px;border:1px solid rgba(16,185,129,.2);}',
        'html.gb-light #gb-toolbar{background:rgba(236,253,245,.9);border-color:rgba(16,185,129,.25);}',
        '#gb-toolbar *{color:#f0fdf4 !important;}',
        'html.gb-light #gb-toolbar *{color:#064e3b !important;}',
        '#gb-toolbar a{color:#fff !important;}',
        'html.gb-light #gb-toolbar a{color:#fff !important;}',

        /* Toggle switch */
        '#gb-toolbar .sw{position:relative;width:44px;height:24px;display:inline-block;}',
        '#gb-toolbar .sw input{opacity:0;width:0;height:0;position:absolute;}',
        '#gb-toolbar .sw .tk{position:absolute;inset:0;background:#1a3328;border-radius:24px;transition:background .3s;border:1px solid rgba(16,185,129,.35);cursor:pointer;}',
        '#gb-toolbar .sw .tk::before{content:"";position:absolute;width:18px;height:18px;left:2px;top:2px;background:#10b981;border-radius:50%;transition:transform .3s cubic-bezier(.4,0,.2,1);box-shadow:0 1px 3px rgba(0,0,0,.3);}',
        '#gb-toolbar .sw input:checked+.tk{background:#a7f3d0;border-color:rgba(5,150,105,.4);}',
        '#gb-toolbar .sw input:checked+.tk::before{transform:translateX(20px);background:#059669;}',
        '#gb-toolbar .ic{font-size:13px;line-height:1;user-select:none;}',

        /* ── Chat: COLLAPSE the stPopover container, position only the button ── */
        '[data-gb-chat="1"]{position:relative !important;width:0 !important;height:0 !important;min-width:0 !important;min-height:0 !important;padding:0 !important;margin:0 !important;overflow:visible !important;display:block !important;}',
        '[data-gb-chat="1"] > button, [data-gb-chat="1"] > div > button{position:fixed !important;top:8px !important;left:60px !important;right:auto !important;z-index:999991 !important;border-radius:50% !important;width:36px !important;height:36px !important;min-height:0 !important;padding:0 !important;background:linear-gradient(135deg,#10b981,#047857) !important;color:#fff !important;border:none !important;box-shadow:0 2px 8px rgba(0,0,0,.25) !important;display:flex !important;align-items:center !important;justify-content:center !important;cursor:pointer !important;}',
        '[data-gb-chat="1"] > button p, [data-gb-chat="1"] > div > button p{font-size:18px !important;margin:0 !important;line-height:1 !important;}',

        /* ── Popover Body: DARK ── */
        'div[data-testid="stPopoverBody"]{background-color:#0f2a1e !important;background:#0f2a1e !important;border:1px solid rgba(16,185,129,.25) !important;border-radius:16px !important;box-shadow:0 8px 32px rgba(0,0,0,.4) !important;width:380px !important;max-width:90vw !important;}',
        'div[data-testid="stPopoverBody"] *:not(svg):not(path):not(button){color:#f0fdf4 !important;}',
        'div[data-testid="stPopoverBody"] div{background-color:transparent !important;background:transparent !important;}',
        'div[data-testid="stPopoverBody"] h3{color:#10b981 !important;}',
        'div[data-testid="stPopoverBody"] [data-testid="stChatMessage"]{background:rgba(16,185,129,.08) !important;border:1px solid rgba(16,185,129,.15) !important;border-radius:12px !important;}',
        'div[data-testid="stPopoverBody"] textarea{background:#1a3328 !important;color:#f0fdf4 !important;border:1px solid rgba(16,185,129,.3) !important;border-radius:12px !important;}',
        'div[data-testid="stPopoverBody"] section[data-testid="stFileUploader"]{background:rgba(16,185,129,.08) !important;border:1px dashed rgba(16,185,129,.4) !important;border-radius:8px !important;}',
        'div[data-testid="stPopoverBody"] button[data-testid="stBaseButton-secondary"]{background:#10b981 !important;color:#fff !important;border:none !important;}',

        /* ── Popover Body: LIGHT ── */
        'html.gb-light div[data-testid="stPopoverBody"]{background-color:#ffffff !important;background:#ffffff !important;border-color:rgba(16,185,129,.3) !important;box-shadow:0 8px 32px rgba(0,0,0,.1) !important;}',
        'html.gb-light div[data-testid="stPopoverBody"] *:not(svg):not(path):not(button){color:#064e3b !important;}',
        'html.gb-light div[data-testid="stPopoverBody"] div{background-color:transparent !important;background:transparent !important;}',
        'html.gb-light div[data-testid="stPopoverBody"] h3{color:#059669 !important;}',
        'html.gb-light div[data-testid="stPopoverBody"] textarea{background:#f8fafc !important;color:#064e3b !important;}',
        'div[data-testid="stPopoverBody"] textarea:focus{border-color:#10b981 !important;}'
    ].join('\\n');

    /* ═══ 2. CREATE TOOLBAR IN PARENT BODY ═══ */
    var tb = pd.getElementById('gb-toolbar');
    if(!tb){
        tb = pd.createElement('div');
        tb.id = 'gb-toolbar';
        pd.body.appendChild(tb);
    } else {
        tb.innerHTML = '';
    }

    // Apenas se o usuário estiver logado
    if (gb_is_logged_in === 'true') {
        var userInfo = pd.createElement('div');
        userInfo.id = 'gb-user-info';
        userInfo.innerHTML = "Logado como: <strong>" + gb_username + "</strong> (" + gb_role + ")";
        userInfo.style.cssText = "font-size:13px;margin-right:12px;opacity:0.85;color:" + txtColor() + ";display:flex;align-items:center;";
        tb.appendChild(userInfo);

        var logoutBtn = pd.createElement('a');
        logoutBtn.href = "?logout=1";
        logoutBtn.target = "_self";
        logoutBtn.textContent = 'Sair';
        logoutBtn.style.cssText = "background:#10b981;color:white;border:none;border-radius:6px;padding:4px 12px;font-size:13px;font-weight:600;cursor:pointer;margin-right:16px;text-decoration:none;display:inline-flex;align-items:center;transition:background 0.3s;";
        logoutBtn.onmouseover = function(){ this.style.background='#059669'; };
        logoutBtn.onmouseout  = function(){ this.style.background='#10b981'; };
        tb.appendChild(logoutBtn);
    }

    var moonSpan = pd.createElement('span');
    moonSpan.className = 'ic';
    moonSpan.id = 'gb-moon';
    moonSpan.textContent = String.fromCodePoint(0x1F319);

    var sunSpan = pd.createElement('span');
    sunSpan.className = 'ic';
    sunSpan.textContent = String.fromCodePoint(0x2600) + String.fromCodePoint(0xFE0F);

    var label = pd.createElement('label');
    label.className = 'sw';
    label.title = 'Alternar tema';

    var checkbox = pd.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = 'gbThemeToggle';

    var track = pd.createElement('span');
    track.className = 'tk';

    label.appendChild(checkbox);
    label.appendChild(track);

    tb.appendChild(moonSpan);
    tb.appendChild(label);
    tb.appendChild(sunSpan);

    checkbox.addEventListener('change', function(){
        var light = this.checked;
        pd.documentElement.classList.toggle('gb-light', light);
        var a = pd.querySelector('.stApp');
        if(a) a.classList.toggle('gb-light', light);
        ls.setItem('gb-theme', light ? 'light' : 'dark');
        var ui = pd.getElementById('gb-user-info');
        if(ui) ui.style.color = light ? '#064e3b' : '#f0fdf4';
    });

    /* ═══ 3. APPLY THEME FROM LOCALSTORAGE ═══ */
    var isLight = ls.getItem('gb-theme') === 'light';
    pd.documentElement.classList.toggle('gb-light', isLight);
    var app = pd.querySelector('.stApp');
    if(app) app.classList.toggle('gb-light', isLight);
    var cb = pd.getElementById('gbThemeToggle');
    if(cb) cb.checked = isLight;

    /* ═══ 4. CHAT BUTTON: find, tag, and position ═══ */
    /* Kill old loop if exists */
    if(window.parent._gbChatLoop){ clearInterval(window.parent._gbChatLoop); window.parent._gbChatLoop=null; }

    function findAndPositionChat(){
        var c = pd.querySelector('[data-gb-chat="1"]');

        /* Se nao encontrou, busca pelo emoji robot */
        if(!c){
            var allBtns = pd.querySelectorAll('button');
            for(var i=0; i<allBtns.length; i++){
                var t = allBtns[i].textContent || '';
                if(t.indexOf(String.fromCodePoint(0x1F916)) !== -1){
                    var container = allBtns[i].closest('[data-testid="stPopover"]');
                    if(!container) container = allBtns[i].parentElement;
                    if(container){
                        container.setAttribute('data-gb-chat','1');
                        c = container;
                        break;
                    }
                }
            }
        }

        if(c){
            /* CSS already sets left:16px on the button via [data-gb-chat="1"] > button.
               Nothing else needed for positioning — CSS handles it. */
        }

        /* ═══ 5. FORCE POPOVER BODY STYLING (JS nuclear override) ═══ */
        var popBodies = pd.querySelectorAll('div[data-testid="stPopoverBody"]');
        for(var bi=0; bi<popBodies.length; bi++){
            var popBody = popBodies[bi];
            var lt = isLightTheme();
            var bgC = lt ? '#ffffff' : '#0f2a1e';
            var fc  = lt ? '#064e3b' : '#f0fdf4';

            /* Force background on the body itself */
            popBody.style.setProperty('background-color', bgC, 'important');
            popBody.style.setProperty('background', bgC, 'important');

            /* Force ALL descendant divs to transparent bg */
            var allDivs = popBody.getElementsByTagName('div');
            for(var d=0; d<allDivs.length; d++){
                allDivs[d].style.setProperty('background-color', 'transparent', 'important');
                allDivs[d].style.setProperty('background', 'transparent', 'important');
            }

            /* Force text color on all text elements */
            var textEls = popBody.querySelectorAll('p,span,label,h1,h2,h3,h4,h5,h6,strong,em,li,small,a,div');
            for(var j=0; j<textEls.length; j++){
                textEls[j].style.setProperty('color', fc, 'important');
            }

            /* Chat messages bg */
            var msgs = popBody.querySelectorAll('[data-testid="stChatMessage"]');
            for(var m=0; m<msgs.length; m++){
                msgs[m].style.setProperty('background', 'rgba(16,185,129,.08)', 'important');
                msgs[m].style.setProperty('border', '1px solid rgba(16,185,129,.15)', 'important');
            }

            /* Textarea */
            var tas = popBody.querySelectorAll('textarea');
            for(var ti=0; ti<tas.length; ti++){
                tas[ti].style.setProperty('background', lt ? '#f8fafc' : '#1a3328', 'important');
                tas[ti].style.setProperty('color', fc, 'important');
            }

            /* H3 accent */
            var h3s = popBody.querySelectorAll('h3');
            for(var k=0; k<h3s.length; k++){
                h3s[k].style.setProperty('color', lt ? '#059669' : '#10b981', 'important');
            }
        }
    }

    /* Run immediately + every 500ms */
    findAndPositionChat();
    window.parent._gbChatLoop = setInterval(findAndPositionChat, 500);

})();
</script>"""
    components.html(js_script, height=0)


def _apply_login_mode(is_login):
    """Adiciona/remove classe que esconde a sidebar na tela de login."""
    action = "add" if is_login else "remove"
    components.html(f"""<script>
    (function(){{
        var app = window.parent.document.querySelector('.stApp');
        if(app) app.classList.{action}('gb-login-mode');
    }})();
    </script>""", height=0)


# ══════════════════════════════════════════════════════════════
# AUTENTICAÇÃO
# ══════════════════════════════════════════════════════════════
init_auth()

is_logged_in = st.session_state.get("authenticated", False)
_user = st.session_state.get("username", "Visitante")
_role = st.session_state.get("user_role", "Visitante")

# Injeta a UI (agora sabe se o user tá logado para mostrar os menus no toolbar)
_inject_ui_enhancements(is_logged_in, _user, _role)

# Se NÃO logado, esconde sidebar e mostra login
if not is_logged_in:
    _apply_login_mode(True)
    login_screen()
    st.stop()
else:
    _apply_login_mode(False)

# ══════════════════════════════════════════════════════════════
# NAVEGAÇÃO MULTIPAGE (Streamlit 1.36+)
# ══════════════════════════════════════════════════════════════
inicio = st.Page("pages/Inicio.py", title="Início", icon="🏠", default=True)
esteira = st.Page("pages/Esteira.py", title="Esteira", icon="🏭")
historico = st.Page("pages/Historico.py", title="Histórico", icon="📦")
ingestao = st.Page("pages/Ingestao_PDF.py", title="Ingestão PDF", icon="📄")
mapa = st.Page("pages/Mapa.py", title="Mapa", icon="🗺️")
checklist = st.Page("pages/Checklist_Separacao.py", title="Checklist Separação", icon="📋")
extracao = st.Page("pages/Extracao_Notas.py", title="Extração Notas", icon="📄")
bi = st.Page("pages/BI_Analytics.py", title="BI Analytics", icon="📊")
expedicao = st.Page("pages/Expedicao.py", title="Expedição", icon="🚚")
hist_expedicoes = st.Page("pages/Historico_Expedicoes.py", title="Histórico Expedições", icon="📆")

pg = st.navigation(
    {
        "Geral": [inicio, esteira, historico],
        "Operações": [ingestao, extracao, mapa, checklist],
        "Expedição & Analytics": [expedicao, hist_expedicoes, bi]
    }
)

# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 10px 0;">
        <div style="font-size: 40px; margin-bottom: -10px;">🌿</div>
        <h1 style="margin: 0; font-size: 22px; font-weight: 700; color: #10b981;">Green Bags</h1>
        <p style="opacity: 0.6; font-size: 13px; margin-top: 4px;">
            Logística e Entregas
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.divider()

    # ── Informação do usuário logado ──
    _user = st.session_state.get("username", "Visitante")
    _role = st.session_state.get("user_role", "Visitante")
    st.caption(f"👤 **{_user}** ({_role})")

    st.divider()

    from modules.auth import is_admin
    if is_admin():
        with st.expander("⚙️ Administração"):
            # -- Reconstruir Índice RAG --
            if st.button("🧠 Reconstruir Índice IA (RAG)", use_container_width=True):
                with st.spinner("Lendo arquivos e recriando o cérebro da GIA..."):
                    try:
                        from modules.rag_engine import rebuild_index
                        if rebuild_index(os.path.dirname(os.path.abspath(__file__))):
                            st.success("Índice RAG reconstruído com sucesso!")
                        else:
                            st.warning("Nenhum texto extraído dos arquivos.")
                    except ImportError:
                        st.error("Pacotes de IA não instalados. Execute: pip install sentence-transformers faiss-cpu")

            # -- Atualizar Cache de Entregas --
            if st.button("🔄 Atualizar Cache Entregas", use_container_width=True,
                         help="Força a leitura do Relatório de Entregas mais recente (XLSB)"):
                import subprocess
                with st.spinner("Extraindo dados do XLSB... (aprox 20s)"):
                    try:
                        subprocess.run(
                            ["powershell", "-ExecutionPolicy", "Bypass", "-File", "scratch/convert_xlsb_xlsx.ps1"],
                            check=True
                        )
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro: {e}")

            # -- DEBUG: Listar modelos Gemini --
            if st.checkbox("🐛 DEBUG: Listar modelos Gemini", key="debug_models"):
                try:
                    from google import genai
                    api_key = st.secrets.get("gemini", {}).get("api_key")
                    client = genai.Client(api_key=api_key)
                    st.write([m.name for m in client.models.list()])
                except Exception as e:
                    st.error(f"Erro: {e}")

# ══════════════════════════════════════════════════════════════
# WIDGET GLOBAL IA (Chat flutuante)
# ══════════════════════════════════════════════════════════════
try:
    from modules.chat_ui import render_floating_chat
    render_floating_chat()
except Exception as e:
    pass

# Executa a página selecionada
pg.run()

with st.sidebar:
    st.divider()
    st.markdown("""
    <div style="padding: 8px 0; opacity: 0.5; font-size: 11px;">
        <strong>Backend:</strong> CONTROLE NOTAS.xlsm<br>
        <strong>Motor:</strong> openpyxl + Streamlit<br>
        <strong>Versão:</strong> 2.0.0
    </div>
    """, unsafe_allow_html=True)

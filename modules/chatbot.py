import streamlit as st
from google import genai
from google.genai import types
import streamlit as st
import pandas as pd
from typing import Optional

def init_gemini() -> bool:
    """Verifica se a API Key do Gemini está configurada."""
    api_key = st.secrets.get("gemini", {}).get("api_key")
    if not api_key:
        return False
    return True

def get_system_context() -> str:
    """Gera um texto de contexto lendo a persona e as expedições ativas."""
    context = ""
    # Tenta carregar a persona do arquivo de configuração
    import os
    persona_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "GIA_Persona.txt")
    try:
        if os.path.exists(persona_path):
            with open(persona_path, "r", encoding="utf-8") as f:
                context += f.read() + "\n\n"
        else:
            context += "Você é a GIA (GreenIA), a assistente virtual de logística.\n\n"
    except Exception:
        context += "Você é a GIA (GreenIA), a assistente virtual de logística.\n\n"
        
    return context

def consultar_pedido_ou_nf(codigo: str) -> str:
    """Busca em TODAS as fontes de dados (Historico, Dados, Relatório de Entregas) o status de um pedido ou NF.
    Argumentos:
        codigo: A numeração exata do Pedido ou da NF a ser buscada.
    """
    try:
        from modules.excel_handler import read_historico, read_principal
        import pandas as pd
        
        all_results = []
        
        def is_match(series: pd.Series, target: str) -> pd.Series:
            # Tenta match exato string
            mask_str = series.astype(str).str.strip().str.upper() == str(target).strip().upper()
            # Tenta match numérico (ignora zeros a esquerda e .0)
            try:
                num_target = float(target)
                mask_num = pd.to_numeric(series, errors='coerce') == num_target
                return mask_str | mask_num
            except:
                return mask_str
        
        # 1. Busca no Historico (aba Historico do CONTROLE NOTAS)
        try:
            df_h = read_historico()
            if not df_h.empty:
                mask_h = pd.Series(False, index=df_h.index)
                if "PEDIDO" in df_h.columns:
                    mask_h = mask_h | is_match(df_h["PEDIDO"], codigo)
                if "NF" in df_h.columns:
                    mask_h = mask_h | is_match(df_h["NF"], codigo)
                found_h = df_h[mask_h]
                if not found_h.empty:
                    found_h = found_h.copy()
                    found_h["_FONTE"] = "Historico"
                    all_results.append(found_h)
        except Exception:
            pass
        
        # 2. Busca na aba Dados (aba principal do CONTROLE NOTAS)
        try:
            df_p = read_principal()
            if not df_p.empty:
                mask_p = pd.Series(False, index=df_p.index)
                if "PEDIDO" in df_p.columns:
                    mask_p = mask_p | is_match(df_p["PEDIDO"], codigo)
                if "NF" in df_p.columns:
                    mask_p = mask_p | is_match(df_p["NF"], codigo)
                found_p = df_p[mask_p]
                if not found_p.empty:
                    found_p = found_p.copy()
                    found_p["_FONTE"] = "Dados"
                    all_results.append(found_p)
        except Exception:
            pass
        
        # 3. Busca no Relatório de Entregas (coluna NOTA_FISCAL, não NF)
        try:
            from modules.delivery_reader import read_deliveries_report
            df_r = read_deliveries_report()
            if not df_r.empty:
                mask_r = pd.Series(False, index=df_r.index)
                if "NOTA_FISCAL" in df_r.columns:
                    mask_r = mask_r | is_match(df_r["NOTA_FISCAL"], codigo)
                if "NF" in df_r.columns:
                    mask_r = mask_r | is_match(df_r["NF"], codigo)
                found_r = df_r[mask_r]
                if not found_r.empty:
                    found_r = found_r.copy()
                    found_r["_FONTE"] = "Relatório de Entregas"
                    all_results.append(found_r)
        except Exception:
            pass
        
        if not all_results:
            return f"Não encontrei nenhum pedido ou NF correspondente a '{codigo}' em nenhuma das bases (Dados, Historico, Relatório de Entregas)."
        
        match = pd.concat(all_results, ignore_index=True)
        
        # Limpa formato numérico para remover .0
        for col in ["PEDIDO", "NF", "NOTA_FISCAL"]:
            if col in match.columns:
                match[col] = match[col].astype(str).str.replace(r'\.0$', '', regex=True)
                
        return match.to_json(orient="records", force_ascii=False)
    except Exception as e:
        return f"Erro ao consultar ferramenta: {str(e)}"

def resumo_expedicoes_ativas() -> str:
    """Traz os veículos, motoristas e entregas ativas do dia atual da esteira de expedição. Use quando perguntarem sobre o que está saindo hoje."""
    try:
        from modules.expedition_engine import read_expeditions
        df = read_expeditions()
        if df.empty:
            return "Nenhuma expedição ativa registrada no momento."
        cols = ["PEDIDO", "DATA_EXPEDICAO", "VEICULO", "MOTORISTA", "STATUS_ENTREGA", "NF", "CLIENTE", "DESTINO"]
        df_view = df[[c for c in cols if c in df.columns]].copy()
        
        # Limpa formato numérico para remover .0
        for col in ["PEDIDO", "NF"]:
            if col in df_view.columns:
                df_view[col] = df_view[col].astype(str).str.replace(r'\.0$', '', regex=True)
                
        return df_view.to_json(orient="records", force_ascii=False)
    except Exception as e:
        return f"Erro ao buscar resumo: {str(e)}"

def analisar_historico_cliente(nome_cliente: str) -> str:
    """Busca entregas de um cliente em TODAS as fontes: aba Historico, aba Dados e Relatório de Entregas.
    Use quando NÃO houver filtro de data. Para filtrar por período, use consultar_entregas_por_periodo.
    Argumentos:
        nome_cliente: Nome parcial ou completo do cliente.
    """
    try:
        from modules.excel_handler import read_historico, read_principal
        import pandas as pd
        
        totais = {}  # {fonte: contagem}
        all_results = []
        
        # 1. Aba Historico do CONTROLE NOTAS
        try:
            df_h = read_historico()
            if not df_h.empty and "CLIENTE" in df_h.columns:
                mask_h = df_h["CLIENTE"].str.contains(nome_cliente, case=False, na=False)
                found_h = df_h[mask_h]
                if not found_h.empty:
                    found_h = found_h.copy()
                    found_h["_FONTE"] = "Historico"
                    totais["Historico"] = len(found_h)
                    all_results.append(found_h)
        except Exception:
            pass
        
        # 2. Aba Dados (principal) do CONTROLE NOTAS
        try:
            df_p = read_principal()
            if not df_p.empty and "CLIENTE" in df_p.columns:
                mask_p = df_p["CLIENTE"].str.contains(nome_cliente, case=False, na=False)
                found_p = df_p[mask_p]
                if not found_p.empty:
                    found_p = found_p.copy()
                    found_p["_FONTE"] = "Dados"
                    totais["Dados"] = len(found_p)
                    all_results.append(found_p)
        except Exception:
            pass
        
        # 3. Relatório de Entregas
        try:
            from modules.delivery_reader import read_deliveries_report
            df_r = read_deliveries_report()
            if not df_r.empty and "CLIENTE" in df_r.columns:
                mask_r = df_r["CLIENTE"].str.contains(nome_cliente, case=False, na=False)
                found_r = df_r[mask_r]
                if not found_r.empty:
                    found_r = found_r.copy()
                    found_r["_FONTE"] = "Relatório de Entregas"
                    totais["Relatório de Entregas"] = len(found_r)
                    all_results.append(found_r)
        except Exception:
            pass
        
        if not all_results:
            return f"Nenhuma entrega encontrada para o cliente '{nome_cliente}' em nenhuma base de dados."
        
        total_geral = sum(totais.values())
        resumo_fontes = ", ".join([f"{fonte}: {qtd}" for fonte, qtd in totais.items()])
        
        if total_geral > 10:
            return (
                f"Encontrei {total_geral} registros para '{nome_cliente}' nas seguintes fontes: {resumo_fontes}. "
                f"Informe ao usuário exatamente assim: 'Encontrei {total_geral} registros de {nome_cliente} "
                f"({resumo_fontes}). Deseja ver todas, as últimas 10, ou filtrar por uma fonte específica? "
                f"Lembrando que quanto maior o volume, maior meu consumo.'"
            )
        
        match = pd.concat(all_results, ignore_index=True)
        
        # Limpa formato numérico para remover .0
        for col in ["PEDIDO", "NF", "NOTA_FISCAL"]:
            if col in match.columns:
                match[col] = match[col].astype(str).str.replace(r'\.0$', '', regex=True)
                
        return match.to_json(orient="records", force_ascii=False)
    except Exception as e:
        return f"Erro ao consultar cliente: {str(e)}"

def consultar_entregas_por_periodo(nome_cliente: str, data_inicio: str, data_fim: str) -> str:
    """Busca entregas de um cliente filtradas por período de datas em TODAS as fontes.
    SEMPRE use esta ferramenta quando o usuário perguntar sobre entregas em uma data, semana, mês ou período específico.
    Argumentos:
        nome_cliente: Nome parcial ou completo do cliente (ex: 'BRF', 'AMBEV').
        data_inicio: Data inicial no formato DD/MM/AAAA (ex: '22/06/2026').
        data_fim: Data final no formato DD/MM/AAAA (ex: '23/06/2026'). Use a mesma data de inicio se for apenas um dia.
    """
    try:
        from modules.excel_handler import read_historico, read_principal
        from datetime import datetime
        import pandas as pd
        
        # Parsear as datas
        try:
            dt_inicio = pd.to_datetime(data_inicio, format='%d/%m/%Y')
            dt_fim = pd.to_datetime(data_fim, format='%d/%m/%Y')
        except Exception:
            return f"Formato de data inválido. Use DD/MM/AAAA. Recebido: início='{data_inicio}', fim='{data_fim}'."
        
        # Incluir o dia final inteiro (até 23:59:59)
        dt_fim = dt_fim + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        
        all_results = []
        totais = {}
        
        # Colunas de data conhecidas em cada fonte
        DATE_COLS = ["DATA", "DATA_ENTREGA", "DATA_EXPEDICAO", "DT_ENTREGA", "DATA_NF"]
        
        def _find_date_col(df):
            """Encontra a coluna de data disponível no DataFrame."""
            for col in DATE_COLS:
                if col in df.columns:
                    return col
            # Tenta qualquer coluna que tenha 'DATA' no nome
            for col in df.columns:
                if "DATA" in col.upper():
                    return col
            return None
        
        def _filter_by_client_and_date(df, fonte_name):
            """Filtra DataFrame por cliente E data."""
            if df.empty:
                return
            
            # Filtro de cliente
            client_col = None
            for col in ["CLIENTE", "REMETENTE"]:
                if col in df.columns:
                    client_col = col
                    break
            if client_col is None:
                return
            
            mask_cliente = df[client_col].astype(str).str.contains(nome_cliente, case=False, na=False)
            
            # Filtro de data
            date_col = _find_date_col(df)
            if date_col is None:
                # Sem coluna de data, retorna só filtro de cliente
                found = df[mask_cliente]
                if not found.empty:
                    found = found.copy()
                    found["_FONTE"] = fonte_name
                    found["_OBS"] = "Sem coluna de data para filtrar período"
                    totais[fonte_name] = len(found)
                    all_results.append(found)
                return
            
            # Converter coluna de data
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            mask_data = (df[date_col] >= dt_inicio) & (df[date_col] <= dt_fim)
            
            found = df[mask_cliente & mask_data]
            if not found.empty:
                found = found.copy()
                found["_FONTE"] = fonte_name
                totais[fonte_name] = len(found)
                all_results.append(found)
        
        # 1. Aba Historico
        try:
            df_h = read_historico()
            _filter_by_client_and_date(df_h, "Historico")
        except Exception:
            pass
        
        # 2. Aba Dados (principal)
        try:
            df_p = read_principal()
            _filter_by_client_and_date(df_p, "Dados")
        except Exception:
            pass
        
        # 3. Relatório de Entregas
        try:
            from modules.delivery_reader import read_deliveries_report
            df_r = read_deliveries_report()
            # No relatório de entregas, o cliente pode estar em REMETENTE ou CLIENTE
            _filter_by_client_and_date(df_r, "Relatório de Entregas")
        except Exception:
            pass
        
        if not all_results:
            return (
                f"Nenhuma entrega encontrada para '{nome_cliente}' "
                f"no período de {data_inicio} a {data_fim} em nenhuma das bases de dados."
            )
        
        match = pd.concat(all_results, ignore_index=True)
        total = len(match)
        resumo_fontes = ", ".join([f"{fonte}: {qtd}" for fonte, qtd in totais.items()])
        
        # Limpa formato numérico
        for col in ["PEDIDO", "NF", "NOTA_FISCAL"]:
            if col in match.columns:
                match[col] = match[col].astype(str).str.replace(r'\.0$', '', regex=True)
        
        # Para resultados grandes, retornar resumo + dados
        header = (
            f"Total: {total} registros de '{nome_cliente}' entre {data_inicio} e {data_fim}. "
            f"Fontes: {resumo_fontes}.\n\n"
        )
        
        return header + match.to_json(orient="records", force_ascii=False)
    except Exception as e:
        return f"Erro ao consultar entregas por período: {str(e)}"

def get_gemini_client():
    if "gemini_client" not in st.session_state:
        api_key = st.secrets.get("gemini", {}).get("api_key")
        # Deixamos o SDK gerenciar a versão de API padrão recomendada (que é v1beta para tools)
        st.session_state["gemini_client"] = genai.Client(api_key=api_key)
    return st.session_state["gemini_client"]



def buscar_em_documentos(pergunta: str) -> str:
    """Busca em PDFs e no histórico genérico informações semânticas. Use para perguntas abertas sobre entregas ou documentos anexados.
    Argumentos:
        pergunta: A pergunta ou tema a ser buscado.
    """
    try:
        from modules.rag_engine import search_rag
        resultados = search_rag(pergunta, k=5)
        return "\n\n".join(resultados)
    except Exception as e:
        return f"Erro ao buscar no índice RAG: {e}"

def get_chat_session():
    """Retorna a sessão atual de chat, inicializando-a se necessário usando o novo SDK."""
    if "gemini_chat" not in st.session_state:
        client = get_gemini_client()
        
        sys_instruction = get_system_context()
        sys_instruction += "\n\nREGRA DE OURO PARA USO DE DADOS:\nVocê é um Agente Autônomo com acesso a ferramentas de banco de dados e documentos (RAG).\nSEMPRE utilize as ferramentas (tools) fornecidas para consultar pedidos, NFs, clientes, expedições ou documentos ANTES de responder.\nNÃO TENTE INVENTAR DADOS. NUNCA DIGA QUE VOCÊ NÃO TEM ACESSO AO SISTEMA EXTERNO. VOCÊ TEM AS FERRAMENTAS PARA ISSO.\n"
        sys_instruction += "\n- Para pesquisar NFs ou Pedidos antigos e de hoje, use a ferramenta 'consultar_pedido_ou_nf'."
        sys_instruction += "\n- Para saber quem está viajando hoje, veículos ou entregas atuais, use a ferramenta 'resumo_expedicoes_ativas'."
        sys_instruction += "\n- Para pesquisar o histórico GERAL de um cliente (sem filtro de data), use 'analisar_historico_cliente'."
        sys_instruction += "\n- IMPORTANTE: Quando o usuário perguntar sobre entregas em uma DATA, SEMANA, MÊS ou PERÍODO específico (ex: 'essa semana', 'dia 22', 'em junho'), use OBRIGATORIAMENTE a ferramenta 'consultar_entregas_por_periodo'. Converta datas relativas (hoje, essa semana, este mês) para o formato DD/MM/AAAA."
        sys_instruction += "\n- Para perguntas gerais, PDFs ou dados amplos, use a ferramenta 'buscar_em_documentos'."
        sys_instruction += "\n1. VOCÊ É ESTRITAMENTE PROIBIDA DE INVENTAR OU ALUCINAR INFORMAÇÕES.\n"
        
        # Inicia o chat usando o modelo escolhido
        chat = client.chats.create(
            model='gemma-4-31b-it',  # Usando o modelo funcional
            config=types.GenerateContentConfig(
                system_instruction=sys_instruction,
                temperature=0.2,
                tools=[consultar_pedido_ou_nf, resumo_expedicoes_ativas, analisar_historico_cliente, consultar_entregas_por_periodo, buscar_em_documentos]
            )
        )

        st.session_state["gemini_chat"] = chat
        
    return st.session_state["gemini_chat"]

def send_message_to_gemini(user_message: str, file_bytes: Optional[bytes] = None, mime_type: str = None) -> str:
    """Envia uma mensagem (e opcionalmente um arquivo) para o Gemini e retorna a resposta."""
    if not init_gemini():
        return "Erro: Chave de API do Gemini não configurada no secrets.toml."
        
    chat = get_chat_session()
    
    import time
    import os
    from datetime import datetime
    
    # Configuração de tentativas (Retries) para lidar com erros 500
    max_retries = 3
    delay = 2  # Segundos de espera entre as tentativas
    
    # Caminho do log de ocorrências
    log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "ocorrencias_gia.txt")
    
    def _log_error(error_msg: str, attempt: int, is_final: bool):
        """Registra o erro no arquivo de ocorrências."""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            status = "FALHA_FINAL" if is_final else f"TENTATIVA_{attempt+1}/{max_retries}"
            user_msg_preview = user_message[:100].replace("\n", " ") if user_message else "(vazio)"
            had_file = "SIM" if file_bytes else "NAO"
            
            log_entry = (
                f"[{timestamp}] [{status}] "
                f"Erro: {error_msg} | "
                f"Arquivo: {had_file} | "
                f"Mensagem: {user_msg_preview}\n"
            )
            
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(log_entry)
        except Exception:
            pass  # Não falhar por causa do log
    
    for attempt in range(max_retries):
        try:
            if file_bytes and mime_type:
                # Upload do arquivo usando o novo SDK
                client = get_gemini_client()
                
                import tempfile
                
                ext = ".pdf" if mime_type == "application/pdf" else ".txt"
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                    tmp.write(file_bytes)
                    tmp_path = tmp.name
                    
                try:
                    uploaded_file = client.files.upload(file=tmp_path, config={'mime_type': mime_type})
                    response = chat.send_message([user_message, uploaded_file])
                finally:
                    os.unlink(tmp_path)
            else:
                response = chat.send_message(user_message)
                
            return response.text
            
        except Exception as e:
            error_str = str(e)
            
            # Se for erro do servidor (500, Internal error) e ainda houver tentativas
            is_server_error = any(code in error_str for code in ["500", "503", "Internal error", "UNAVAILABLE", "RESOURCE_EXHAUSTED"])
            
            if is_server_error and attempt < max_retries - 1:
                _log_error(error_str, attempt, is_final=False)
                time.sleep(delay)
                delay *= 2  # Backoff exponencial
                continue
            
            # Erro final — registrar e retornar mensagem profissional
            _log_error(error_str, attempt, is_final=True)
            
            if is_server_error:
                return (
                    "⚠️ O serviço da GIA está temporariamente indisponível. "
                    "Tentamos processar sua solicitação múltiplas vezes sem sucesso. "
                    "Por favor, tente novamente em alguns minutos."
                )
            else:
                return (
                    "⚠️ Ocorreu um erro ao processar sua mensagem. "
                    "Nossa equipe foi notificada e o incidente foi registrado. "
                    "Por favor, tente novamente ou reformule sua pergunta."
                )


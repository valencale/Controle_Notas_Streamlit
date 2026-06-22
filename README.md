# 🌿 Green Bags - Sistema de Gestão Logística e Entregas

Este projeto é uma aplicação web de gestão logística focada em otimizar o controle de notas fiscais, pedidos, expedição de veículos e roteirização das entregas. Criado utilizando o framework **Streamlit** (Python), ele consolida processos que antes eram manuais em uma interface amigável, inteligente e de fácil acesso.

## 🎯 Finalidade do Projeto

A finalidade principal do sistema é atuar como uma **Torre de Controle Logístico**. Ele permite que a equipe operacional, administradores e motoristas acompanhem em tempo real o status de cada pedido (desde a separação até a entrega final), além de realizar o fechamento financeiro, gerenciar ocorrências (GIA) e automatizar o cadastro de dados a partir de documentos em PDF.

O projeto foi construído para atender as demandas logísticas de empresas parceiras (como ONFINITY e GREEN BAGS), cruzando dados de motoristas, veículos e notas fiscais.

---

## ✨ Principais Features

### 1. 📊 Painel de Operações (Dashboard Central)
Visualização em tempo real das operações do dia. Inclui KPIs (Indicadores-Chave de Desempenho) como total de entregas, status dos pedidos (Em Separação, Parcial, Em Rota, Entregue) e atalhos rápidos de navegação.

### 2. 🤖 Extração Inteligente de PDFs
Módulo automatizado que lê os PDFs de "Mapa de Separação por Pedido", identificando automaticamente:
- Endereços completos e CEPs.
- Números de Pedido e Notas Fiscais.
- Quantidade de Volumes e Pesos.
Isso elimina a necessidade de digitação manual, inserindo os dados extraídos diretamente na base do sistema.

### 3. 🗺️ Mapa Interativo de Roteirização
Geração de mapas interativos (utilizando *Folium* e geocodificação via *Nominatim/ViaCep*) com todos os pontos de entrega do dia. Permite que os analistas de logística visualizem a densidade das entregas e planejem a distribuição entre os motoristas.

### 4. 🚚 Planejamento de Expedição
Controle completo de veículos e frotas. Permite despachar pedidos específicos para placas de veículos pré-cadastradas (ex: QJJ-9302, RLK-0E24), designar motoristas e acompanhar a ordem de entrega da carga.

### 5. 💰 Módulo Financeiro e de Repasses
Cálculos automatizados de fechamento logístico:
- Custo por motorista e quilometragem rodada.
- Resumo de Entregas (Baseado no histórico "Viagens WhatsApp").
- Visão de receitas brutas e faturamento líquido.

### 6. 💬 Assistente Virtual (GIA)
Integração com Inteligência Artificial para atuar como uma "assistente de tráfego" (GIA). O chatbot tem contexto sobre os pedidos do dia, podendo responder rapidamente a dúvidas como *"Onde está o pedido da cliente Maria?"* ou *"Quantas entregas faltam para o veículo placa XYZ?"*.

### 7. 🔐 Controle de Acesso (RBAC)
Sistema de autenticação com separação de perfis:
- **Administrador**: Acesso total, incluindo permissão para editar dados, aprovar faturamentos e modificar configurações.
- **Visitante**: Acesso apenas para leitura e visualização de painéis.

### 8. 📂 Arquitetura Baseada em Arquivos (Excel & CSV)
O banco de dados do sistema funciona operando e sincronizando diretamente com planilhas Excel (como `CONTROLE NOTAS.xlsm` e `expedicoes.xlsx`), facilitando a extração de relatórios manuais se necessário e mantendo a familiaridade do usuário com o ecossistema Office.

---

## 🚀 Como Executar Localmente

1. **Clone o repositório:**
   ```bash
   git clone https://github.com/SEU-USUARIO/Controle_Notas_Streamlit.git
   ```

2. **Crie e ative o ambiente virtual:**
   ```bash
   python -m venv armazem_venv
   armazem_venv\Scripts\activate
   ```

3. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configuração de Secrets (Senha Admin e IA):**
   - Na raiz do projeto, crie a pasta `.streamlit`.
   - Dentro dela, crie o arquivo `secrets.toml` copiando a estrutura do `secrets.example.toml` e insira suas chaves e senhas.

5. **Inicie o sistema:**
   ```bash
   streamlit run app.py
   ```
   *(Ou utilize os atalhos `START_APP.bat` / `INICIAR_LOGISTICA.bat` inclusos no projeto)*.

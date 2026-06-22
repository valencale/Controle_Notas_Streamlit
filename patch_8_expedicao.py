import re

filepath = "pages/8_🚚_Expedicao.py"
with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Add import
if "from modules.auth import is_adm" not in content:
    content = content.replace("import streamlit as st", "import streamlit as st\nfrom modules.auth import is_adm")

# 2. Add disabled=not is_adm() to st.button
# We'll find all st.button( ... )
# Some might already have disabled=...
# We can use regex to find st.button
def patch_button(match):
    m = match.group(0)
    if "disabled=" in m:
        if "is_adm" in m:
            return m
        # replace disabled=X with disabled=(X) or not is_adm()
        return re.sub(r"disabled=([^,\)]+)", r"disabled=(\1) or not is_adm()", m)
    else:
        # replace closing paren with , disabled=not is_adm())
        # wait, the match is just `st.button(` we need to find the matching closing paren.
        pass
    return m

# Since regex for balanced parens is hard, let's just do it by tokenizing or simpler logic.
# Wait, let's just use Python's ast? No, ast unparsing loses comments and formatting.

# Let's do string replacement for the specific lines we found.
replacements = [
    ('if st.button("▶ Atribuir Selecionados", type="primary", use_container_width=True, disabled=not final_vehicle):',
     'if st.button("▶ Atribuir Selecionados", type="primary", use_container_width=True, disabled=(not final_vehicle) or not is_adm()):'),
     
    ('if st.button("💾 Associar NF", use_container_width=True):',
     'if st.button("💾 Associar NF", use_container_width=True, disabled=not is_adm()):'),

    ('if st.button("💾 Salvar Plano", type="primary", use_container_width=True):',
     'if st.button("💾 Salvar Plano", type="primary", use_container_width=True, disabled=not is_adm()):'),

    ('if st.button("📥 Exportar Checklist", use_container_width=True):',
     'if st.button("📥 Exportar Checklist", use_container_width=True, disabled=not is_adm()):'),

    ('if st.button("🗑️ Limpar Plano", use_container_width=True):',
     'if st.button("🗑️ Limpar Plano", use_container_width=True, disabled=not is_adm()):'),

    ('if st.button("✅ Sim, limpar", key="confirm_clear_yes"):',
     'if st.button("✅ Sim, limpar", key="confirm_clear_yes", disabled=not is_adm()):'),

    ('if st.button("💾 Salvar", key="save_drivers"):',
     'if st.button("💾 Salvar", key="save_drivers", disabled=not is_adm()):'),

    ('if st.button("✏️ Aplicar Alterações da Tabela", type="secondary"):',
     'if st.button("✏️ Aplicar Alterações da Tabela", type="secondary", disabled=not is_adm()):'),

    ('if st.button("💾 Salvar Todas as Viagens", type="primary", use_container_width=True, key="exp_save_batch"):',
     'if st.button("💾 Salvar Todas as Viagens", type="primary", use_container_width=True, key="exp_save_batch", disabled=not is_adm()):'),

    ('if st.button("🗑️ Excluir Viagem", type="secondary", key="exp_btn_delete_viagem", use_container_width=True):',
     'if st.button("🗑️ Excluir Viagem", type="secondary", key="exp_btn_delete_viagem", use_container_width=True, disabled=not is_adm()):'),

    ('if st.button("✅ Sim, excluir", type="primary", key="exp_confirm_del_yes", use_container_width=True):',
     'if st.button("✅ Sim, excluir", type="primary", key="exp_confirm_del_yes", use_container_width=True, disabled=not is_adm()):'),

    ('if st.button("✏️ Editar Viagem Selecionada", key="exp_btn_edit_viagem", use_container_width=True):',
     'if st.button("✏️ Editar Viagem Selecionada", key="exp_btn_edit_viagem", use_container_width=True, disabled=not is_adm()):'),

    ('save_edit = st.form_submit_button("💾 Salvar Alterações", type="primary", use_container_width=True)',
     'save_edit = st.form_submit_button("💾 Salvar Alterações", type="primary", use_container_width=True, disabled=not is_adm())'),

    ('if st.button(\n                        f"🏁 Marcar pedidos do {placa} como ENTREGUE",\n                        key=f"deliver_{placa}",\n                        use_container_width=True,\n                    ):',
     'if st.button(\n                        f"🏁 Marcar pedidos do {placa} como ENTREGUE",\n                        key=f"deliver_{placa}",\n                        use_container_width=True,\n                        disabled=not is_adm(),\n                    ):'),
     
    ('if st.button(\n                        "🔍 Verificar no Relatório",\n                        type="secondary",\n                        use_container_width=True,\n                        disabled=valid_count == 0,\n                        key="reverse_verify_batch",\n                    ):',
     'if st.button(\n                        "🔍 Verificar no Relatório",\n                        type="secondary",\n                        use_container_width=True,\n                        disabled=(valid_count == 0) or not is_adm(),\n                        key="reverse_verify_batch",\n                    ):'),

    ('if st.button(\n                        "✅ Confirmar Todas",\n                        type="primary",\n                        use_container_width=True,\n                        disabled=valid_count == 0,\n                        key="reverse_confirm_batch",\n                        help="Confirma TODAS as DANFEs — encontradas e não encontradas",\n                    ):',
     'if st.button(\n                        "✅ Confirmar Todas",\n                        type="primary",\n                        use_container_width=True,\n                        disabled=(valid_count == 0) or not is_adm(),\n                        key="reverse_confirm_batch",\n                        help="Confirma TODAS as DANFEs — encontradas e não encontradas",\n                    ):'),

    ('if st.button(\n                            f"✅ Confirmar Apenas Encontradas ({found_count})",\n                            type="primary",\n                            use_container_width=True,\n                            key="reverse_confirm_found_only",\n                        ):',
     'if st.button(\n                            f"✅ Confirmar Apenas Encontradas ({found_count})",\n                            type="primary",\n                            use_container_width=True,\n                            key="reverse_confirm_found_only",\n                            disabled=not is_adm(),\n                        ):'),

    ('if st.button("🗑️ Limpar resultados", key="reverse_clear_results"):',
     'if st.button("🗑️ Limpar resultados", key="reverse_clear_results", disabled=not is_adm()):'),

    ('submitted = st.form_submit_button(\n                    "🔍 Verificar e Confirmar",\n                    type="primary",\n                    use_container_width=True,\n                )',
     'submitted = st.form_submit_button(\n                    "🔍 Verificar e Confirmar",\n                    type="primary",\n                    use_container_width=True,\n                    disabled=not is_adm()\n                )'),

     ('nf_files = st.file_uploader(\n                "Upload da NF (DANFE)",\n                type=["pdf"],\n                key="nf_upload",\n                accept_multiple_files=True,\n            )',
      'nf_files = st.file_uploader(\n                "Upload da NF (DANFE)",\n                type=["pdf"],\n                key="nf_upload",\n                accept_multiple_files=True,\n                disabled=not is_adm(),\n            )'),

      ('uploaded = st.file_uploader(\n            "Upload do chat exportado (.txt)",\n            type=["txt"],\n            key="exp_chat_upload",\n        )',
       'uploaded = st.file_uploader(\n            "Upload do chat exportado (.txt)",\n            type=["txt"],\n            key="exp_chat_upload",\n            disabled=not is_adm(),\n        )'),

       ('danfe_files = st.file_uploader(\n            "Selecione os PDFs de DANFE",\n            type=["pdf"],\n            accept_multiple_files=True,\n            key="reverse_danfe_upload",\n            label_visibility="collapsed",\n        )',
        'danfe_files = st.file_uploader(\n            "Selecione os PDFs de DANFE",\n            type=["pdf"],\n            accept_multiple_files=True,\n            key="reverse_danfe_upload",\n            label_visibility="collapsed",\n            disabled=not is_adm(),\n        )'),
]

for old, new in replacements:
    if old in content:
        content = content.replace(old, new)
    else:
        print("COULD NOT FIND:", old[:50])

with open(filepath, "w", encoding="utf-8") as f:
    f.write(content)

print("Patch applied to", filepath)

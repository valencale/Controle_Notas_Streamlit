import os
import re

def patch_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Add import if not present
    if "from modules.auth import is_adm" not in content:
        # Find where to insert it. Usually after "import streamlit as st"
        content = re.sub(
            r"(import streamlit as st)",
            r"\1\nfrom modules.auth import is_adm",
            content,
            count=1
        )

    # 2. Patch st.button
    # Regex explanation:
    # st\.button\([^)]+\) -> need to carefully inject disabled=not is_adm() if not present
    # Since buttons can span multiple lines, let's just do a simpler search:
    # Replace `st.button(` with something else? No, button arguments are comma separated.
    # We can do this: look for `st.button(` and if it doesn't have `disabled=`, add it.
    
    # A robust way: find `if st.button(` or similar. Actually, some buttons already have `disabled=`.
    # Let's just find `st.button(` up to the closing `)` and insert `disabled=not is_adm()`.
    
    def replace_button(match):
        full_match = match.group(0)
        inner = match.group(1)
        if "disabled=" in inner:
            # If it already has disabled, we might want to combine it, e.g. disabled=not is_adm() or (original)
            # This is hard. For now, let's just replace `disabled=(.*?)` with `disabled=(not is_adm()) or (\1)`
            # actually let's see what's there
            if "not is_adm()" in inner:
                return full_match # Already patched
            # Let's do a simple regex: disabled=([^,]+)
            def replace_disabled(m):
                orig = m.group(1).strip()
                return f"disabled=(not is_adm()) or ({orig})"
            new_inner = re.sub(r"disabled=([^,)]+)", replace_disabled, inner)
            if new_inner == inner: # failed to replace or no match inside? 
                return full_match
            return full_match.replace(inner, new_inner)
        else:
            # Add disabled=not is_adm()
            # Find the end of the argument list and insert it.
            # `inner` does not have the closing paren, so we can just append `, disabled=not is_adm()`
            # wait, `full_match` includes the closing paren if we match correctly?
            # Let's match: st\.button\((.*?)\)  -- this fails if there are nested parens.
            pass
        return full_match

    # Let's use an easier method: Python script to parse AST? Too complex to rewrite.
    # Let's just use string replacement on known lines.
    pass

if __name__ == "__main__":
    pass

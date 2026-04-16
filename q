[33mcommit 44cc6fab3deadecd98f75da563fc4c16e93446aa[m[33m ([m[1;36mHEAD[m[33m -> [m[1;32mmain[m[33m)[m
Author: BlueBalou <estebanpueblos@googlemail.com>
Date:   Thu Apr 16 18:45:13 2026 +0200

    password protection

[1mdiff --git a/streamlit_app.py b/streamlit_app.py[m
[1mindex e1f7bb9..dc4ba9d 100644[m
[1m--- a/streamlit_app.py[m
[1m+++ b/streamlit_app.py[m
[36m@@ -112,9 +112,57 @@[m [mdef _init_session_state() -> None:[m
 _init_session_state()[m
 [m
 # ---------------------------------------------------------------------------[m
[31m-# Page layout[m
[32m+[m[32m# Password gate[m
 # ---------------------------------------------------------------------------[m
 [m
[32m+[m[32mdef _check_password() -> bool:[m
[32m+[m[32m    if st.session_state.get("authenticated"):[m
[32m+[m[32m        return True[m
[32m+[m
[32m+[m[32m    st.markdown([m
[32m+[m[32m        """[m
[32m+[m[32m        <style>[m
[32m+[m[32m        .login-box {[m
[32m+[m[32m            max-width: 340px;[m
[32m+[m[32m            margin: 8rem auto 0 auto;[m
[32m+[m[32m            padding: 2rem 2rem 1.5rem 2rem;[m
[32m+[m[32m            border-radius: 8px;[m
[32m+[m[32m            background: #1a1a1a;[m
[32m+[m[32m            box-shadow: 0 4px 24px rgba(0,0,0,0.5);[m
[32m+[m[32m            text-align: center;[m
[32m+[m[32m        }[m
[32m+[m[32m        .login-title {[m
[32m+[m[32m            font-size: 1.2rem;[m
[32m+[m[32m            font-weight: 600;[m
[32m+[m[32m            color: #e0e0e0;[m
[32m+[m[32m            margin-bottom: 1.5rem;[m
[32m+[m[32m            letter-spacing: 0.05em;[m
[32m+[m[32m        }[m
[32m+[m[32m        </style>[m
[32m+[m[32m        <div class="login-box">[m
[32m+[m[32m            <div class="login-title">🔒 Wochenplan Scheduler</div>[m
[32m+[m[32m        </div>[m
[32m+[m[32m        """,[m
[32m+[m[32m        unsafe_allow_html=True,[m
[32m+[m[32m    )[m
[32m+[m
[32m+[m[32m    col1, col2, col3 = st.columns([1, 2, 1])[m
[32m+[m[32m    with col2:[m
[32m+[m[32m        pw = st.text_input("Passwort", type="password", label_visibility="collapsed",[m
[32m+[m[32m                           placeholder="Passwort eingeben")[m
[32m+[m[32m        if st.button("Anmelden", use_container_width=True, type="primary"):[m
[32m+[m[32m            if pw == st.secrets.get("password", ""):[m
[32m+[m[32m                st.session_state["authenticated"] = True[m
[32m+[m[32m                st.rerun()[m
[32m+[m[32m            else:[m
[32m+[m[32m                st.error("Falsches Passwort.")[m
[32m+[m[32m    return False[m
[32m+[m
[32m+[m[32mif not _check_password():[m
[32m+[m[32m    st.stop()[m
[32m+[m
[32m+[m
[32m+[m
 st.set_page_config([m
     page_title="Wochenplan Scheduler",[m
     page_icon="📋",[m

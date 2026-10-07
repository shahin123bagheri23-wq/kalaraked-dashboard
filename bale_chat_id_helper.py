"""کمک برای گرفتن chat_id از بله و نسبت دادن به مخاطبین"""
import streamlit as st
from email_manager import load_contacts, save_contacts
from bale_sender import collect_chat_ids


def render_chat_id_fetcher(key_suffix="up"):
    st.markdown("#### 📥 دریافت خودکار chat_id از بله")
    st.caption("لیست کسانی که به ربات پیام دادن رو می‌گیره و به مخاطبین نسبت می‌ده")

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("🔄 دریافت لیست از بله", key=f"fetch_chats_{key_suffix}",
                     type="primary", use_container_width=True):
            try:
                chats = collect_chat_ids()
                st.session_state[f"bale_chats_{key_suffix}"] = chats
                if not chats:
                    st.warning("هیچ چتی پیدا نشد. مطمئن شو به ربات پیام دادن.")
            except Exception as e:
                st.error(f"❌ {e}")
    with col_b:
        if st.button("🗑 پاک کردن لیست", key=f"clear_chats_{key_suffix}",
                     use_container_width=True):
            st.session_state.pop(f"bale_chats_{key_suffix}", None)
            st.rerun()

    chats = st.session_state.get(f"bale_chats_{key_suffix}")
    if not chats:
        st.info("💡 روی دکمه «دریافت لیست از بله» بزن.")
        return

    st.success(f"✅ {len(chats)} چت پیدا شد")

    contacts = load_contacts()

    for chat in chats:
        cid = chat["chat_id"]
        name = chat["name"] or "—"
        username = chat["username"]

        with st.container(border=True):
            col_a, col_b = st.columns([2, 2])
            with col_a:
                st.markdown(f"**👤 {name}**")
                if username:
                    st.caption(f"@{username}")
            with col_b:
                st.markdown(f"**chat_id:** `{cid}`")

            already = [c for c in contacts if c.get("bale_chat_id") == cid]
            if already:
                existing_name = already[0].get("name") or already[0].get("email")
                st.caption(f"✅ قبلاً به «{existing_name}» نسبت داده شده")
                continue

            options = ["(انتخاب کن)"] + [
                f"{c.get('name') or c.get('email')} — {c.get('email') or '—'}"
                for c in contacts
            ]
            selected = st.selectbox(
                "به کدام مخاطب نسبت داده بشه؟",
                options,
                key=f"assign_{cid}_{key_suffix}",
            )

            if st.button("✅ ثبت", key=f"save_{cid}_{key_suffix}", type="primary"):
                if selected == "(انتخاب کن)":
                    st.warning("یک مخاطب انتخاب کن")
                else:
                    idx = options.index(selected) - 1
                    contacts[idx]["bale_chat_id"] = cid
                    save_contacts(contacts)
                    st.success(f"✅ به «{contacts[idx].get('name')}» نسبت داده شد")
                    st.rerun()
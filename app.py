from datetime import date
import traceback

import pandas as pd
import streamlit as st

from agent import generate_report
from sheets_loader import discover_master_sheet_structure


def _format_exception(error: Exception) -> str:
    message = str(error).strip()
    if not message:
        message = repr(error)
    if not message:
        message = "Unknown error (no details returned by library)."
    return f"{type(error).__name__}: {message}"

st.set_page_config(page_title="Inventory Order Agent", page_icon="🏭", layout="wide")

with st.sidebar:
    st.image("https://img.icons8.com/color/96/factory.png", width=80)
    st.title("Inventory Agent")
    st.markdown("---")
    st.markdown("### 🔗 Google Sheet")
    st.caption("Use a single Google Sheet URL with: 'software stock' + monthly tabs (e.g., June-2024)")

    master_sheet_url = st.text_input(
        "Master Google Sheet URL",
        value=st.session_state.get("master_sheet_url", ""),
        placeholder="https://docs.google.com/spreadsheets/d/...",
        key="master_sheet_url_input",
    )
    lt_url = st.text_input(
        "Lead Time Sheet URL (optional)",
        value=st.session_state.get("lt_url", ""),
        placeholder="https://docs.google.com/spreadsheets/d/...",
        key="lt_url_input",
    )

    if master_sheet_url:
        st.session_state["master_sheet_url"] = master_sheet_url
    if lt_url:
        st.session_state["lt_url"] = lt_url

    st.markdown("---")
    st.markdown("### ⏱️ Default Lead Times (days)")
    mach_lead = st.number_input("Machining Lead Time", min_value=1, value=7, key="mach_lead_input")
    mfg_lead = st.number_input("Manufacturing Lead Time", min_value=1, value=30, key="mfg_lead_input")

    st.markdown("---")
    st.markdown("### 🔌 Connection Status")
    st.markdown("🟢 Master Sheet" if master_sheet_url else "🔴 Master Sheet — URL missing")
    st.markdown("🟡 Lead Time" if lt_url else "🔵 Lead Time — using defaults")

    if master_sheet_url:
        st.markdown("---")
        st.markdown("### 📋 Sheet Preview")
        if st.button("🔍 Preview Sheet Structure", key="preview_btn"):
            try:
                structure = discover_master_sheet_structure(master_sheet_url)
                st.markdown(
                    "🟢 software stock tab found"
                    if structure["has_software_stock"]
                    else "🔴 software stock tab missing"
                )

                monthly_tabs = structure["monthly_tabs"]
                if monthly_tabs:
                    st.caption(f"Detected monthly sales tabs: {len(monthly_tabs)}")
                    st.dataframe(pd.DataFrame(monthly_tabs), use_container_width=True, hide_index=True)
                else:
                    st.warning("No monthly sales tabs found. Expected names like June-2024.")
            except Exception as preview_error:
                st.warning(f"Could not preview sheet structure: {_format_exception(preview_error)}")
                with st.expander("Show technical details"):
                    st.code(traceback.format_exc())

    st.markdown("---")
    st.caption("Built for V-Pulley Inventory Management")

st.title("🏭 Inventory Order Agent")
st.markdown(f"**Today:** {date.today().strftime('%A, %d %B %Y')}")

if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": (
                "👋 Hello! I'm your Inventory Order Agent.\n\n"
                "**To get started:**\n"
                "1. Paste one **Master Google Sheet URL** in the sidebar\n"
                "2. Ensure stock tab is **software stock** with **Stocklevel** and **Reorderlevel**\n"
                "3. Ensure historical sales tabs are named like **June-2024**\n"
                "4. Click **Generate Report**\n"
            ),
        }
    ]

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

col1, col2, _ = st.columns([1, 1, 2])
with col1:
    gen_btn = st.button("📊 Generate Report", type="primary", use_container_width=True, key="generate_btn")
with col2:
    clr_btn = st.button("🗑️ Clear Chat", use_container_width=True, key="clear_btn")

if clr_btn:
    st.session_state.messages = st.session_state.messages[:1]
    st.rerun()


def run_report():
    if not master_sheet_url:
        return "⚠️ Please paste the **Master Google Sheet URL** in the sidebar first."
    try:
        with st.spinner("📡 Reading stock + monthly sales tabs from Google Sheets..."):
            buf, mach_df, mfg_df, today = generate_report(
                master_sheet_url,
                lt_url if lt_url else None,
                mach_lead,
                mfg_lead,
            )

        st.session_state["report_buf"] = buf
        st.session_state["report_fname"] = f"Inventory_Order_Report_{today.strftime('%Y-%m-%d')}.xlsx"
        st.session_state["mach_df"] = mach_df
        st.session_state["mfg_df"] = mfg_df

        total_mach = int(mach_df["Suggested Order Qty"].sum())
        total_mfg = int(mfg_df["Suggested Order Qty"].sum())

        return (
            f"✅ **Report generated for {today.strftime('%d %B %Y')}**\n\n"
            f"| | Items | Total Units |\n"
            f"|---|---|---|\n"
            f"| 🔧 Machining Orders | {len(mach_df)} | {total_mach} |\n"
            f"| 🏭 Manufacturing Orders (RC) | {len(mfg_df)} | {total_mfg} |"
        )
    except Exception as e:
        return f"❌ Error: {_format_exception(e)}"


def answer_question(question):
    if "mach_df" not in st.session_state:
        return "Please generate a report first, then I can answer questions about it."

    mach_df = st.session_state["mach_df"]
    mfg_df = st.session_state["mfg_df"]
    q = question.lower()

    if any(w in q for w in ["how many", "count", "total items"]):
        if "machining" in q:
            return f"There are **{len(mach_df)} machining orders** in today's report."
        if any(w in q for w in ["manufacturing", "rc", "rough casting"]):
            return f"There are **{len(mfg_df)} manufacturing (RC) orders** in today's report."
        return f"Today's report has **{len(mach_df)} machining orders** and **{len(mfg_df)} manufacturing orders**."

    if "shortage" in q and any(w in q for w in ["highest", "most", "critical", "top"]):
        all_df = pd.concat([
            mach_df[["Product Name", "Shortage", "Suggested Order Qty"]],
            mfg_df[["RC Product Name", "Shortage", "Suggested Order Qty"]].rename(columns={"RC Product Name": "Product Name"}),
        ]).sort_values("Shortage", ascending=False).head(5)

        rows = "\n".join([
            f"- **{r['Product Name']}** — Shortage: {int(r['Shortage'])}, Order: {int(r['Suggested Order Qty'])}"
            for _, r in all_df.iterrows()
        ])
        return f"**Top 5 critical items by shortage:**\n\n{rows}"

    if any(w in q for w in ["total units", "total order", "how many units"]):
        return (
            f"**Total units to order today:**\n\n"
            f"- 🔧 Machining: **{int(mach_df['Suggested Order Qty'].sum())} units**\n"
            f"- 🏭 Manufacturing: **{int(mfg_df['Suggested Order Qty'].sum())} units**"
        )

    return (
        "I can answer questions like:\n"
        "- *How many machining orders?*\n"
        "- *How many manufacturing orders?*\n"
        "- *Show top critical items*\n"
        "- *Total units to order?*"
    )


if gen_btn:
    response = run_report()
    st.session_state.messages.append({"role": "user", "content": "Generate report"})
    st.session_state.messages.append({"role": "assistant", "content": response})
    st.rerun()

if "report_buf" in st.session_state:
    st.download_button(
        label="📥 Download Report",
        data=st.session_state["report_buf"],
        file_name=st.session_state["report_fname"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="download_btn",
    )

if prompt := st.chat_input("Ask a question or type 'generate report'..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    if any(w in prompt.lower() for w in ["generate", "report", "create report", "run"]):
        response = run_report()
    else:
        response = answer_question(prompt)

    st.session_state.messages.append({"role": "assistant", "content": response})
    with st.chat_message("assistant"):
        st.markdown(response)
    st.rerun()

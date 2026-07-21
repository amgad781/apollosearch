"""
Streamlit UI for Apollo exhibitor contact enrichment.

Deploy this on Streamlit Community Cloud or Hugging Face Spaces (see
DEPLOY.md). Nobody's API key is stored anywhere except in that user's
browser session while the tab is open.
"""

import io
import time

import openpyxl
from openpyxl.styles import Font
import streamlit as st

import apollo_core as core

st.set_page_config(page_title="Exhibitor Contact Enrichment", page_icon="🔎")
st.title("🔎 Exhibitor Purchasing-Contact Finder")
st.caption("Uses your own Apollo.io API key. Nothing is stored server-side beyond this session.")

# ---------------------------------------------------------------- session ---
if "cache" not in st.session_state:
    st.session_state.cache = {}   # {company_name: match_company()-style dict}
if "credits_used" not in st.session_state:
    st.session_state.credits_used = 0

# ------------------------------------------------------------------ inputs --
st.subheader("1. Inputs")

key_file = st.file_uploader("Apollo API key (.txt file containing just the key)", type=["txt"])
xlsx_file = st.file_uploader("Exhibitor list (.xlsx with a 'Company name' column)", type=["xlsx"])

api_key = None
if key_file is not None:
    api_key = key_file.read().decode("utf-8").strip()
    st.success("API key loaded for this session (not saved anywhere).")

titles_text = st.text_area(
    "Target job titles (one per line)",
    value="\n".join(core.TARGET_TITLES_DEFAULT),
    height=150,
)
target_titles = [t.strip() for t in titles_text.splitlines() if t.strip()]

st.divider()

# --------------------------------------------------------------- debug -----
with st.expander("🔧 Debug: test a single company"):
    test_name = st.text_input("Company name to test", value="")
    if st.button("Test this one company", disabled=not (api_key and test_name)):
        try:
            result = core.match_company(test_name, api_key, target_titles)
            st.json(result)
        except Exception as e:
            st.error(f"Raw error: {repr(e)}")
            if hasattr(e, "response") and e.response is not None:
                st.code(f"Status: {e.response.status_code}\nBody: {e.response.text[:1000]}")

st.divider()

# ------------------------------------------------------------- phase 1 -----
st.subheader("2. Match companies (free — no credits spent)")

col_a, col_b = st.columns([1, 1])
with col_a:
    run_matching = st.button("Run matching", disabled=not (api_key and xlsx_file))
with col_b:
    if st.button("Clear cache / start over"):
        st.session_state.cache = {}
        st.session_state.credits_used = 0
        st.rerun()

if run_matching:
    names = core.read_company_names(xlsx_file)
    progress = st.progress(0.0, text="Starting...")
    log = st.empty()

    consecutive_errors = 0
    for i, name in enumerate(names, 1):
        if name in st.session_state.cache and st.session_state.cache[name].get("matched") is not None:
            continue
        try:
            result = core.match_company(name, api_key, target_titles)
            consecutive_errors = 0
        except core.ApolloError as e:
            st.error(f"Apollo auth/plan error: {e}")
            st.stop()
        except Exception as e:
            result = {"matched": False, "org_id": None, "people": [], "error": str(e)}
            consecutive_errors += 1
            log.warning(f"Error on '{name}': {e}")
            if consecutive_errors >= 5:
                st.error(
                    f"Stopped after 5 consecutive identical-looking failures. "
                    f"Last error: {e}\n\n"
                    f"This usually means every call is failing the same way "
                    f"(wrong endpoint, plan doesn't include this API, or bad "
                    f"request format) rather than genuine per-company misses. "
                    f"Fix this before continuing — re-running now will just "
                    f"burn through the rest of the list with the same error."
                )
                st.session_state.cache[name] = result
                break
        st.session_state.cache[name] = result
        progress.progress(i / len(names), text=f"{i}/{len(names)}: {name}")
        time.sleep(core.REQUEST_DELAY_SECONDS)
    log.empty()

    matched = sum(1 for c in st.session_state.cache.values() if c.get("matched"))
    with_people = sum(1 for c in st.session_state.cache.values() if c.get("people"))
    st.success(
        f"Done. {matched}/{len(names)} companies matched in Apollo. "
        f"{with_people} have at least one candidate contact. "
        f"No credits spent."
    )

if st.session_state.cache:
    total_people = sum(len(c.get("people", [])) for c in st.session_state.cache.values())
    st.info(f"Currently cached: {len(st.session_state.cache)} companies, "
            f"{total_people} candidate contacts found so far.")

st.divider()

# ------------------------------------------------------------- phase 2 -----
st.subheader("3. Reveal emails (spends Apollo credits)")

max_credits = st.number_input("Max credits to spend this run", min_value=1, max_value=10000, value=100)

if st.button("Reveal emails", disabled=not (api_key and st.session_state.cache)):
    revealed = 0
    progress = st.progress(0.0, text="Starting...")

    pending = [
        (name, p)
        for name, info in st.session_state.cache.items()
        for p in info.get("people", [])
        if not p["email_revealed"]
    ]
    total_to_try = min(len(pending), max_credits)

    for idx, (name, person) in enumerate(pending):
        if revealed >= max_credits:
            break
        try:
            email = core.reveal_person(person["id"], api_key)
        except core.ApolloError as e:
            st.error(str(e))
            break
        except Exception as e:
            st.warning(f"Skipped {person.get('name')}: {e}")
            continue
        person["email"] = email
        person["email_revealed"] = True
        revealed += 1
        st.session_state.credits_used += 1
        progress.progress(min(revealed / max(total_to_try, 1), 1.0),
                           text=f"Revealed {revealed}/{total_to_try}")
        time.sleep(core.REQUEST_DELAY_SECONDS)

    st.success(f"Revealed {revealed} email(s) this run. "
               f"Total revealed this session: {st.session_state.credits_used}.")

st.divider()

# ------------------------------------------------------------- export ------
st.subheader("4. Download results")

if st.session_state.cache:
    rows = core.build_report_rows(st.session_state.cache)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contacts"
    headers = ["Company", "Org Domain", "Contact Name", "Title", "Email", "LinkedIn", "Match Status"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.download_button(
        "Download enriched_contacts.xlsx",
        data=buf,
        file_name="enriched_contacts.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.caption("Run step 2 first to see a download button here.")

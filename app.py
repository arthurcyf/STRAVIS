import time
import streamlit as st
from script_core import run_automation

ALL_ENTITIES = [
    "D341_HSO_HGM",
    "D342_HSO_HGMD",
    "CC41_HSO_HMIN",
    "C741_HSO_HMSH",
    "AN41_HSO_HMSP",
    "D941_HSO_HMSZ",
    "J34V_HSO_HOME",
    "EM41_HSO_HSEU",
    "A441_HSO_HSOT",
    "WB41_HSOU",
    "D841_HSO_HSOK",
    "CY41_HSO_INNOVIA",
    "WM41_HSO_MID LAB INC",
    "GG41_HSO_FRITZ RUCK",
    "GH41_HSO_EOS",
]

DEFAULT_SELECTED = [
    "D341_HSO_HGM",
    "D342_HSO_HGMD",
    "CC41_HSO_HMIN",
    "C741_HSO_HMSH",
    "AN41_HSO_HMSP",
    "D941_HSO_HMSZ",
    "J34V_HSO_HOME",
    "EM41_HSO_HSEU",
    "A441_HSO_HSOT",
    "WB41_HSOU",
    "D841_HSO_HSOK",
]

st.set_page_config(page_title="STRAVIS Automation", page_icon="🛠️", layout="centered")
st.title("STRAVIS Automation Runner")
st.caption("Fill the inputs, click **Run automation** and immediately switch focus to STRAVIS.")
st.info("ℹ️ All files will be saved into your system's 'Downloads' folder.")

with st.form("params"):
    target_period = st.text_input(
        "Target period (YYYY.MM)",
        value="2025.03",
        help="Used in Ctrl+F search inside STRAVIS",
    )

    st.markdown("**Select the entities to INCLUDE (others will be deselected):**")
    selected_entities = []
    cols = st.columns(3)
    for i, ent in enumerate(ALL_ENTITIES):
        with cols[i % 3]:
            checked = st.checkbox(ent, value=(ent in DEFAULT_SELECTED), key=f"ent_{i}")
            if checked:
                selected_entities.append(ent)

    submitted = st.form_submit_button("Run automation", use_container_width=True)

if submitted:
    if len(selected_entities) == 0:
        st.error("Please select at least one entity.")
        st.stop()

    # Deselect everything that is NOT selected
    to_deselect = [e for e in ALL_ENTITIES if e not in selected_entities]

    # iterations should equal number of selected entities
    iterations = len(selected_entities)

    st.warning(
        "After you press **OK**, you have ~3 seconds to bring STRAVIS to the foreground. "
        "Don't touch mouse/keyboard while it runs. Move mouse to top-left corner to abort."
    )
    st.button("OK", key="ok", disabled=True)
    with st.spinner("Starting in 3 seconds… switch to STRAVIS now"):
        time.sleep(3)

    try:
        # Hardcode Shift+Down rows to 20 per your request
        run_automation(target_period, to_deselect, select_n=20, iterations=iterations)
        st.success("Automation finished without raising errors.")
    except Exception as e:
        st.error(f"Automation failed: {e}")
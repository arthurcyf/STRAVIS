# gradio_app.py
import time
import threading
import platform
import gradio as gr
from gradio.themes.utils import colors, fonts  # 👈 add this

# Must be Windows (STRAVIS + UIA)
if platform.system() != "Windows":
    raise SystemExit("This app must run on Windows (UI automation uses Windows APIs).")

# Lazy import so the module import doesn't explode if script_core has Windows-only imports
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

# Prevent concurrent runs (UI automation must be single-threaded)
_lock = threading.Lock()

def submit(period: str, selected: list[str]):
    if not period or len(period) != 7 or period[4] != ".":
        return "❌ Period must look like YYYY.MM (e.g., 2025.03)."
    if not selected:
        return "❌ Please select at least one entity."

    if not _lock.acquire(blocking=False):
        return "⚠️ A run is already in progress. Please wait."

    try:
        # User heads-up & handoff time
        yield "⏳ Starting in 3 seconds… please bring STRAVIS to the foreground and do not touch the mouse/keyboard."
        time.sleep(3)

        # Try NEW signature first: (target_period, selected_entities, all_entities, iterations)
        try:
            run_automation(
                target_period=period,
                selected_entities=selected,
                all_entities=ALL_ENTITIES,
                iterations=len(selected),
            )
            yield f"✅ Finished. Exported {len(selected)} entities (new API)."
            return
        except TypeError:
            # Fallback to OLD signature: (target_period, to_deselect, select_n, iterations)
            to_deselect = [e for e in ALL_ENTITIES if e not in selected]
            # Set select_n=1 so it doesn't grab batches
            run_automation(period, to_deselect, select_n=1, iterations=len(selected))
            yield f"✅ Finished. Exported {len(selected)} entities (legacy API)."
            return

    except Exception as e:
        yield f"❌ Automation failed: {e}"
    finally:
        _lock.release()

blue_theme = gr.themes.Soft(
    primary_hue=colors.blue,
    secondary_hue=colors.blue,
    neutral_hue=colors.gray,
    font=fonts.GoogleFont("Inter"),             # <-- UI text font
    font_mono=fonts.GoogleFont("JetBrains Mono")# <-- code/mono font
)

with gr.Blocks(theme=blue_theme) as demo:  # 👈 apply theme here
    gr.Markdown("# STRAVIS Automation\nLog into STRAVIS, Fill in the period, choose entities, then click **Run automation**.")
    period = gr.Textbox(label="Target period (YYYY.MM)", value="2025.03")
    entities = gr.CheckboxGroup(choices=ALL_ENTITIES, value=DEFAULT_SELECTED, label="Entities to INCLUDE")
    run_btn = gr.Button("Run automation", variant="primary")
    out = gr.Textbox(label="Status", lines=8)

    # stream=True lets us yield status messages
    run_btn.click(fn=submit, inputs=[period, entities], outputs=out, api_name="run", scroll_to_output=True)

# LAN access on 8501; set share=True for a temporary public link (be cautious!)
demo.queue().launch(server_name="0.0.0.0", server_port=8501)

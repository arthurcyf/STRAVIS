# ───────────────────────────── script_core.py ─────────────────────────────
"""
Core automation entrypoint refactored from your script.py.
Call run_automation(target_period, to_deselect) from any UI (e.g., Streamlit).

Notes:
- Run as Administrator for best UIA reliability.
- Keep STRAVIS on the primary monitor and do not touch keyboard/mouse while it runs.
- Move the mouse to the top-left corner to trigger PyAutoGUI failsafe if you need to abort.
"""
import time, re, ctypes, collections
import pyautogui
import uiautomation as ui

# Virtual-Key codes
VK_SHIFT    = 0x10
VK_DOWN     = 0x28
KEYEVENTF_KEYDOWN = 0x0000
KEYEVENTF_KEYUP   = 0x0002

pyautogui.FAILSAFE = True

# ------------- helpers copied from your script (unchanged unless parameterized) -------------

def find_control(root, **kwargs):
    Name = kwargs.pop('Name', None)
    AutomationId = kwargs.pop('AutomationId', None)
    searchDepth = kwargs.pop('searchDepth', 10)
    timeout = kwargs.pop('timeout', 5.0)
    retry_interval = kwargs.pop('retry_interval', 0.5)

    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            ctrl = root.Control(Name=Name, AutomationId=AutomationId, searchDepth=searchDepth)
            if ctrl and ctrl.Exists(0, 0):
                return ctrl
        except Exception:
            pass
        time.sleep(retry_interval)
    return None


def shift_select_down(n=20, delay=0.05):
    ctypes.windll.user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYDOWN, 0)
    time.sleep(0.02)
    for _ in range(n):
        ctypes.windll.user32.keybd_event(VK_DOWN, 0, KEYEVENTF_KEYDOWN, 0)
        ctypes.windll.user32.keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP,   0)
        time.sleep(delay)
    ctypes.windll.user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)


def deselect_entity(code, search_delay=1.0, clear_delay=0.2):
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(clear_delay)
    pyautogui.write(code, interval=0.02)
    time.sleep(search_delay)

    ui.SendKeys('{DOWN}')
    time.sleep(0.1)
    ui.SendKeys('{SPACE}')

    pyautogui.hotkey('ctrl', 'f')
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(clear_delay)
    pyautogui.press('backspace')
    time.sleep(0.1)

def press_open():
    for _ in range(3):
        ui.SendKeys('{TAB}')
        time.sleep(0.1)
    ui.SendKeys('{ENTER}')

def find_with_retry(factory_fn, timeout=8, interval=0.2):
    end = time.time() + timeout
    last_err = None
    while time.time() < end:
        try:
            ctrl = factory_fn()
            if ctrl.Exists(1, 0.1):
                return ctrl
        except Exception as e:
            last_err = e
        time.sleep(interval)
    raise RuntimeError(f"Find_with_retry timeout. Last error: {last_err}")


def switch_ribbon_tab(window, tab_name, ribbon_name='The Ribbon', tabs_name='Ribbon Tabs'):
    ribbon = find_with_retry(lambda: window.PaneControl(Name=ribbon_name, searchDepth=6))
    tabs   = find_with_retry(lambda: ribbon.TabControl(Name=tabs_name, searchDepth=4))
    tab    = find_with_retry(lambda: tabs.TabItemControl(Name=tab_name, searchDepth=3))
    try:
        tab.Click()
    except Exception:
        tab.GetInvokePattern().Invoke()


def click_button(root, *, Name=None, AutomationId=None, searchDepth=10, timeout=5.0):
    btn = find_control(root, Name=Name, AutomationId=AutomationId, searchDepth=searchDepth, timeout=timeout)
    if not btn:
        desc = AutomationId or Name
        raise RuntimeError(f"Could not find button '{desc}' under {root}")
    btn.Click()


def safe_snapshot(root):
    snap = []
    try:
        children = list(root.GetChildren())
    except Exception:
        return snap

    for c in children:
        try:
            if not c.Exists(0, 0):
                continue
        except Exception:
            continue
        try:
            rid = tuple(c.GetRuntimeId())
        except Exception:
            rid = None
        try:
            ctype = getattr(c, 'ControlTypeName', None)
        except Exception:
            ctype = None
        try:
            name = c.Name
        except Exception:
            name = ""
        snap.append((rid, ctype, name))
    return snap


def wait_for_change(root, snapshot_fn=None, timeout=10.0, interval=0.5):
    if snapshot_fn is None:
        snapshot_fn = safe_snapshot
    before = snapshot_fn(root)
    end = time.time() + timeout
    while time.time() < end:
        time.sleep(interval)
        after = snapshot_fn(root)
        if after != before:
            return True
    return False


def wait_for_base_input(stravis, name='Base List/Data Input', timeout=12):
    deadline = time.time() + timeout
    last_err = None
    while time.time() < deadline:
        try:
            win = ui.WindowControl(Name=name)
            if win.Exists(1, 0.1):
                return win
            pane = stravis.PaneControl(Name=name, searchDepth=30)
            if pane.Exists(1, 0.1):
                return pane
        except Exception as e:
            last_err = e
        time.sleep(0.2)
    raise RuntimeError(f"Timed out waiting for '{name}' (last error: {last_err})")


def press_e(root_for_waits=None):
    for _ in range(15):
        ui.SendKeys('{DOWN}')
        time.sleep(0.1)
    for _ in range(3):
        ui.SendKeys('{RIGHT}')
        time.sleep(0.1)
    for _ in range(2):
        pyautogui.hotkey('ctrl', 'up')

    pyautogui.hotkey('alt', 'down')
    if root_for_waits is not None:
        wait_for_change(root_for_waits, timeout=5, interval=0.2)

    ui.SendKeys('{DOWN}')
    ui.SendKeys('{ENTER}')
    if root_for_waits is not None:
        wait_for_change(root_for_waits, timeout=5, interval=0.2)


def wait_until_tab_active(window, tab_name='Operation', ribbon_name='The Ribbon', tabs_name='Ribbon Tabs', timeout=10, interval=0.2):
    deadline = time.time() + timeout
    last_exc = None
    while time.time() < deadline:
        try:
            ribbon = window.PaneControl(Name=ribbon_name, searchDepth=10)
            tabs = ribbon.TabControl(Name=tabs_name, searchDepth=10)
            tab = tabs.TabItemControl(Name=tab_name, searchDepth=5)
            if tab.Exists(0, 0):
                try:
                    if tab.GetSelectionItemPattern().IsSelected:
                        return tab
                except Exception:
                    pass
                try:
                    if getattr(tab, 'HasKeyboardFocus', False):
                        return tab
                except Exception:
                    pass
        except Exception as e:
            last_exc = e
        time.sleep(interval)
    raise RuntimeError(f"Tab '{tab_name}' not active within {timeout}s (last error: {last_exc})")


def click_save_as_excel(stravis, timeout=8):
    wait_until_tab_active(stravis, 'Operation')
    ribbon = stravis.PaneControl(Name='The Ribbon', searchDepth=10)
    lower  = ribbon.PaneControl(Name='Lower Ribbon', searchDepth=6)
    op     = lower.PaneControl(Name='Operation', searchDepth=6)
    filetb = op.ToolBarControl(Name='File', searchDepth=6)

    btn = filetb.ButtonControl(Name='Save As Excel', searchDepth=3)
    if not btn.Exists(5, 0.2):
        btn = ribbon.ButtonControl(Name='Save As Excel', searchDepth=30)
    if not btn.Exists(5, 0.2):
        raise RuntimeError("Could not find 'Save As Excel' button")

    btn.Click()
    wait_for_change(ui.GetRootControl(), timeout=8, interval=0.3)


def click_save_as_tree_item(target='Downloads', timeout=10):
    deadline = time.time() + timeout
    save_win = None
    while time.time() < deadline:
        w = ui.WindowControl(Name='Save As')
        if w.Exists(0, 0):
            save_win = w
            break
        time.sleep(0.2)
    if not save_win:
        raise RuntimeError("Save As window not found")

    search_root = save_win
    try:
        side = save_win.PaneControl(Name='sidePanel1', searchDepth=10)
        if side.Exists(0, 0):
            search_root = side
    except Exception:
        pass

    try:
        data_panel = search_root.GroupControl(Name='Data Panel', searchDepth=10)
        if data_panel.Exists(0, 0):
            search_root = data_panel
    except Exception:
        pass

    q = collections.deque([search_root])
    target_ctrl = None
    while q:
        node = q.popleft()
        try:
            if getattr(node, 'ControlTypeName', None) == 'DataItemControl':
                try:
                    vp = node.GetValuePattern()
                    val = vp.Value
                except Exception:
                    val = None
                name_ok = False
                try:
                    name_ok = (node.Name == target)
                except Exception:
                    pass
                if val == target or name_ok:
                    target_ctrl = node
                    break
        except Exception:
            pass
        try:
            q.extend(node.GetChildren())
        except Exception:
            pass

    if not target_ctrl:
        raise RuntimeError(f"Could not find DataItem with Value '{target}' in Save As")

    try:
        target_ctrl.GetInvokePattern().Invoke()
    except Exception:
        target_ctrl.Click()

    try:
        wait_for_change(save_win, timeout=5, interval=0.2)
    except Exception:
        pass


def click_operation_close(stravis, timeout=8):
    wait_until_tab_active(stravis, 'Operation')
    ribbon = stravis.PaneControl(Name='The Ribbon', searchDepth=10)
    lower  = ribbon.PaneControl(Name='Lower Ribbon', searchDepth=8)
    op     = lower.PaneControl(Name='Operation', searchDepth=8)

    btn = op.ButtonControl(Name='Close', searchDepth=20)
    if not btn.Exists(5, 0.2):
        btn = lower.ButtonControl(Name='Close', searchDepth=30) if not btn.Exists(0,0) else btn
    if not btn.Exists(5, 0.2):
        btn = ribbon.ButtonControl(Name='Close', searchDepth=40)
    if not btn.Exists(5, 0.2):
        raise RuntimeError("Could not find 'Close' button in the ribbon")

    try:
        btn.GetInvokePattern().Invoke()
    except Exception:
        btn.Click()

    wait_for_change(stravis, timeout=8, interval=0.3)


def wait_dialog_gone(name='Save As', timeout=10, interval=0.2):
    end = time.time() + timeout
    while time.time() < end:
        try:
            if not ui.WindowControl(Name=name).Exists(0, 0):
                return True
        except Exception:
            pass
        time.sleep(interval)
    raise RuntimeError(f"Dialog '{name}' did not close in time")


# ------------- MAIN PARAMETERIZED ENTRYPOINT -------------

def run_automation(target_period: str, to_deselect: list[str], select_n: int = 20, iterations: int = 11):
    """Run the STRAVIS flow using the given period string (e.g., '2025.03')
    and a list of entity codes to deselect.
    """
    if not re.match(r"^\d{4}\.\d{2}$", target_period):
        raise ValueError("target_period must look like 'YYYY.MM', e.g. '2025.03'")

    ui.SetGlobalSearchTimeout(3.0)

    # 1) Attach to STRAVIS
    stravis = ui.WindowControl(Name='STRAVIS')
    if not stravis.Exists(10, 0.2):
        raise RuntimeError('STRAVIS window not found')
    stravis.SetFocus()

    # 2) Double-click Data Collection (Node1)
    dc_node = find_control(stravis, Name='Node1', timeout=8)
    if not dc_node:
        raise RuntimeError("Could not find 'Node1' (Data Collection)")
    dc_node.DoubleClick()

    # 3) Double-click Base List/Data Input (Node2)
    time.sleep(1)
    ui.SendKeys('{DOWN}')
    time.sleep(0.1)
    ui.SendKeys('{ENTER}')
    print("Clicked Base List/Data Input")

    if not wait_for_change(stravis, timeout=10, interval=0.5):
        raise RuntimeError('UI didn’t change after opening Base List/Data Input')

    base_input = wait_for_base_input(stravis, 'Base List/Data Input', timeout=12)
    if not base_input.Exists(5, 0.2):
        raise RuntimeError('Base List/Data Input exists check failed unexpectedly')
    base_input.SetFocus()

    # 5) Click the period entry matching AY…(YTD)
    pattern = re.compile(r'^AY.*\(YTD\)$')
    queue = collections.deque([base_input])
    period_ctrl = None
    while queue:
        node = queue.popleft()
        try:
            name = node.Name or ''
        except Exception:
            name = ''
        if pattern.match(name):
            period_ctrl = node
            break
        try:
            queue.extend(node.GetChildren())
        except Exception:
            pass

    if not period_ctrl:
        raise RuntimeError('No period entry matching AY…(YTD) found')

    period_ctrl.Click()

    # 6) Activate Clear via Down+Space
    ui.SendKeys('{DOWN}')
    time.sleep(0.1)
    ui.SendKeys('{SPACE}')

    # 7) Ctrl+F and type the requested period
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.2)
    pyautogui.write(target_period, interval=0.02)
    time.sleep(1)

    # 8) Select found item
    ui.SendKeys('{DOWN}')
    time.sleep(0.1)
    ui.SendKeys('{SPACE}')

    # 9) Ensure Operation tab, locate org pane, click Open
    wait_dialog_gone('Save As')
    switch_ribbon_tab(stravis, 'Operation')

    org_pane = base_input.PaneControl(AutomationId='pnlCndOrganization', searchDepth=30)
    if not org_pane.Exists(8, 0.2):
        raise RuntimeError("Organization pane not found (AutomationId='pnlCndOrganization')")
    open_btn = org_pane.ButtonControl(Name='Open', searchDepth=8)
    if not open_btn.Exists(5, 0.2):
        raise RuntimeError('Open button not found in Organization pane')
    open_btn.Click()

    # 10) Select all, then deselect the provided list
    ui.SendKeys('{DOWN}')
    time.sleep(0.1)
    for _ in range(20):
        ui.SendKeys('{UP}')
        time.sleep(0.05)
    shift_select_down(n=20, delay=0.05)
    ui.SendKeys('{SPACE}')

    for code in to_deselect:
        deselect_entity(code)

    # 11) Run Display
    switch_ribbon_tab(stravis, 'Operation')
    click_button(base_input, Name='Display', AutomationId='btnDisp', searchDepth=30, timeout=8)
    wait_for_change(base_input, timeout=15, interval=0.5)

    # 12) Iterate items and save-as flow (unchanged from your logic)
    time.sleep(1)
    for _ in range(2):
        ui.SendKeys('{DOWN}')
        time.sleep(0.1)

    for i in range(iterations):
        press_open()
        wait_until_tab_active(stravis, 'Operation')
        time.sleep(20)  # consider replacing with a waiter if you want
        press_e(root_for_waits=stravis)

        switch_ribbon_tab(stravis, 'Operation')
        wait_until_tab_active(stravis, 'Operation')
        click_save_as_excel(stravis)
        click_save_as_tree_item('Downloads')
        time.sleep(1)
        for _ in range(4):
            ui.SendKeys('{TAB}')
            time.sleep(0.1)
        ui.SendKeys('{ENTER}')

        switch_ribbon_tab(stravis, 'Operation')
        click_operation_close(stravis)
        for _ in range(4):
            ui.SendKeys('{TAB}')
            time.sleep(0.1)

        for _ in range(3):
            ui.SendKeys('{DOWN}')
    
    print("Download Complete")
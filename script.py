import uiautomation as ui
import pyautogui
import time
import ctypes
import collections, re

# Virtual-Key codes
VK_SHIFT    = 0x10
VK_DOWN     = 0x28
KEYEVENTF_KEYDOWN = 0x0000
KEYEVENTF_KEYUP   = 0x0002

def find_control(root, **kwargs):
    """
    Attempts to find a control matching the given properties under `root`.
    Supports Name, AutomationId, and searchDepth, with a retry loop.
    """
    Name = kwargs.pop('Name', None)
    AutomationId = kwargs.pop('AutomationId', None)
    searchDepth = kwargs.pop('searchDepth', 10)
    timeout = kwargs.pop('timeout', 5.0)
    retry_interval = kwargs.pop('retry_interval', 0.5)

    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            ctrl = root.Control(
                Name=Name,
                AutomationId=AutomationId,
                searchDepth=searchDepth
            )
            if ctrl and ctrl.Exists(0, 0):
                return ctrl
        except Exception:
            pass
        time.sleep(retry_interval)
    return None

def shift_select_down(n=20, delay=0.05):
    """
    Hold Shift, press Down n times, then release Shift.
    """
    ctypes.windll.user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYDOWN, 0)
    time.sleep(0.05)
    for _ in range(n):
        ctypes.windll.user32.keybd_event(VK_DOWN, 0, KEYEVENTF_KEYDOWN, 0)
        ctypes.windll.user32.keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP,   0)
        time.sleep(delay)
    ctypes.windll.user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)

def deselect_entity(code, search_delay=1.0, clear_delay=0.2):
    """
    Uses Ctrl+F to find `code`, presses Down+Space to toggle it off,
    then clears the search box.
    """
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(clear_delay)
    pyautogui.write(code, interval=0.02)
    print(f'Opened Find and entered "{code}"')
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

def switch_ribbon_tab(window, tab_name,
                      ribbon_name='The Ribbon',
                      tabs_name='Ribbon Tabs'):
    # smaller depths reduce walking and flakiness
    ribbon = find_with_retry(lambda: window.PaneControl(Name=ribbon_name, searchDepth=6))
    tabs   = find_with_retry(lambda: ribbon.TabControl(Name=tabs_name, searchDepth=4))
    tab    = find_with_retry(lambda: tabs.TabItemControl(Name=tab_name, searchDepth=3))
    try:
        tab.Click()
    except Exception:
        # Click fallback
        tab.GetInvokePattern().Invoke()
    print(f"Switched to the {tab_name} tab")

def click_button(root, *, Name=None, AutomationId=None, searchDepth=10, timeout=5.0):
    """
    Finds a ButtonControl under `root` by Name or AutomationId and clicks it.
    """
    btn = find_control(root,
                       Name=Name,
                       AutomationId=AutomationId,
                       searchDepth=searchDepth,
                       timeout=timeout)
    if not btn:
        desc = AutomationId or Name
        raise RuntimeError(f"Could not find button '{desc}' under {root}")
    btn.Click()

def safe_snapshot(root):
    """
    Take a defensive snapshot of immediate children: (RuntimeId, ControlTypeName, Name).
    """
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
    """
    Poll until the UI under `root` changes, or timeout.
    """
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
    """
    Resolve the 'Base List/Data Input' surface as either a top-level WindowControl
    or a PaneControl inside STRAVIS. Returns the resolved control.
    """
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
    """
    Navigate language dropdown without a long sleep.
    - root_for_waits: a container whose children change when menus open/close (e.g., the new detail window or STRAVIS)
    """
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

def wait_until_tab_active(window, tab_name='Operation',
                          ribbon_name='The Ribbon',
                          tabs_name='Ribbon Tabs',
                          timeout=10, interval=0.2):
    """
    Wait until a specific ribbon tab is selected (or at least focused).
    Returns the TabItemControl on success; raises on timeout.
    """
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
    # Make sure the Operation tab is actually active first
    wait_until_tab_active(stravis, 'Operation')

    # Walk the exact ancestor chain from your UIA dump
    ribbon = stravis.PaneControl(Name='The Ribbon', searchDepth=10)
    lower  = ribbon.PaneControl(Name='Lower Ribbon', searchDepth=6)
    op     = lower.PaneControl(Name='Operation', searchDepth=6)
    filetb = op.ToolBarControl(Name='File', searchDepth=6)

    # Find the button by Name and click
    btn = filetb.ButtonControl(Name='Save As Excel', searchDepth=3)
    if not btn.Exists(5, 0.2):
        # Fallback: search anywhere under the ribbon if layout shifts
        btn = ribbon.ButtonControl(Name='Save As Excel', searchDepth=30)
    if not btn.Exists(5, 0.2):
        raise RuntimeError("Could not find 'Save As Excel' button")

    btn.Click()
    print("Clicked 'Save As Excel'")

    # Optional: wait for the Save dialog (or any top-level window change)
    wait_for_change(ui.GetRootControl(), timeout=8, interval=0.3)

def click_save_as_tree_item(target='Downloads', timeout=3):
    """
    In the 'Save As' dialog, find the left tree item by ValuePattern.Value (e.g. 'Downloads')
    and click/invoke it.
    """
    # 1) Wait for the Save As window
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

    # 2) Prefer searching under sidePanel1 → Data Panel if present
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

    # 3) BFS: find a DataItem whose ValuePattern.Value == target (or Name == target as fallback)
    q = collections.deque([search_root])
    target_ctrl = None
    while q:
        node = q.popleft()
        try:
            # Check if it's a DataItem with ValuePattern
            if getattr(node, 'ControlTypeName', None) == 'DataItemControl':
                try:
                    vp = node.GetValuePattern()  # raises if not available
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

        # Enqueue children
        try:
            q.extend(node.GetChildren())
        except Exception:
            pass

    if not target_ctrl:
        raise RuntimeError(f"Could not find DataItem with Value '{target}' in Save As")

    # 4) Invoke if possible; otherwise click
    try:
        target_ctrl.GetInvokePattern().Invoke()
    except Exception:
        target_ctrl.Click()

    print(f"Selected '{target}' in Save As")
    # Optional: wait for the right pane to refresh
    try:
        wait_for_change(save_win, timeout=5, interval=0.2)
    except Exception:
        pass

def click_operation_close(stravis, timeout=8):
    """Click the 'Close' ribbon button under Operation, then wait for a UI change."""
    # Ensure we're actually on the Operation tab
    wait_until_tab_active(stravis, 'Operation')

    ribbon = stravis.PaneControl(Name='The Ribbon', searchDepth=10)
    lower  = ribbon.PaneControl(Name='Lower Ribbon', searchDepth=8)
    op     = lower.PaneControl(Name='Operation', searchDepth=8)

    # The toolbar is unnamed in your dump, so just find the button by Name under Operation
    btn = op.ButtonControl(Name='Close', searchDepth=20)
    if not btn.Exists(5, 0.2):
        # Fallbacks in case layout shifts
        btn = lower.ButtonControl(Name='Close', searchDepth=30) if not btn.Exists(0,0) else btn
    if not btn.Exists(5, 0.2):
        btn = ribbon.ButtonControl(Name='Close', searchDepth=40)

    if not btn.Exists(5, 0.2):
        raise RuntimeError("Could not find 'Close' button in the ribbon")

    # Prefer Invoke if available (more reliable than Click on ribbon controls)
    try:
        btn.GetInvokePattern().Invoke()
    except Exception:
        btn.Click()

    print("Clicked 'Close'")

    # Give the app a moment to update – use structural change instead of a hard sleep
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

def main():
    ui.SetGlobalSearchTimeout(3.0)

    # 1) Attach to the STRAVIS main window
    stravis = ui.WindowControl(Name='STRAVIS')
    if not stravis.Exists(10, 0.2):
        raise RuntimeError("STRAVIS window not found")
    stravis.SetFocus()

    # 2) Double-click the “Data Collection” node (Node1)
    dc_node = find_control(stravis, Name='Node1', timeout=8)
    if not dc_node:
        raise RuntimeError("Could not find 'Node1' (Data Collection)")
    dc_node.DoubleClick()

    # 3) Double-click the “Base List/Data Input” item (Node2)
    # bli_item = find_control(stravis, Name='Node2', timeout=8)
    # if not bli_item:
    #     raise RuntimeError("Could not find 'Node2' (Base List/Data Input)")
    # bli_item.DoubleClick()
    time.sleep(1)
    ui.SendKeys('{DOWN}')
    time.sleep(0.1)
    ui.SendKeys('{ENTER}')
    print("Clicked Base List/Data Input")


    # Wait for the STRAVIS window's children to change (view opening)
    if not wait_for_change(stravis, timeout=10, interval=0.5):
        raise RuntimeError("UI didn’t change after opening Base List/Data Input")

    # Resolve and focus the Base List/Data Input host (window or pane)
    base_input = wait_for_base_input(stravis, 'Base List/Data Input', timeout=12)
    if not base_input.Exists(5, 0.2):
        raise RuntimeError("Base List/Data Input exists check failed unexpectedly")
    base_input.SetFocus()

    # 5) Click the period entry
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
        raise RuntimeError("No period entry matching AY…(YTD) found")

    period_ctrl.Click()
    print(f"Clicked period entry '{period_ctrl.Name}'")

    # 6) Activate Clear via Down+Space
    ui.SendKeys('{DOWN}')
    time.sleep(0.1)
    ui.SendKeys('{SPACE}')
    print("Sent Down-Arrow + Space to activate Clear")

    # 7) Fire Ctrl+F and type search
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.2)
    pyautogui.write('2025.03', interval=0.02)
    print('Opened Find and entered "2025.03"')
    time.sleep(1)

    # 8) Select found item
    ui.SendKeys('{DOWN}')
    time.sleep(0.1)
    ui.SendKeys('{SPACE}')

    # 9) Switch to the Operation tab
    # wait_dialog_gone('Save As')   # ensures the tree is stable again
    switch_ribbon_tab(stravis, 'Operation')

    # 10) Click the Open button in organization pane
    org_pane = base_input.PaneControl(AutomationId='pnlCndOrganization', searchDepth=30)
    if not org_pane.Exists(8, 0.2):
        raise RuntimeError("Organization pane not found (AutomationId='pnlCndOrganization')")
    open_btn = org_pane.ButtonControl(Name='Open', searchDepth=8)
    if not open_btn.Exists(5, 0.2):
        raise RuntimeError("Open button not found in Organization pane")
    open_btn.Click()
    print("Clicked the Open button.")

    # 11) Select all the entities
    ui.SendKeys('{DOWN}')
    time.sleep(0.1)
    shift_select_down(n=20, delay=0.05)
    print("Performed Shift + Down x20 via Win32 events.")
    ui.SendKeys('{SPACE}')

    # 12) Deselect the 4 items
    to_deselect = [
        'CY41_HSO_INNOVIA',
        'WM41_HSO_MID LAB INC',
        'GG41_HSO_FRITZ RUCK',
        'GH41_HSO_EOS'
    ]
    for code in to_deselect:
        deselect_entity(code)

    # 13) Click Operation to exit
    switch_ribbon_tab(stravis, 'Operation')

    # 14) Click 'display' and wait for the pane to refresh
    click_button(base_input, Name='Display', AutomationId='btnDisp', searchDepth=30, timeout=8)
    print("Clicked 'Display'")
    wait_for_change(base_input, timeout=15, interval=0.5)

    # 15) Move down to the first entity and open
    time.sleep(1)
    for _ in range(2):
        ui.SendKeys('{DOWN}')
        time.sleep(0.1)

    for i in range(11):
        press_open()

        # Ensure Operation tab is active before next step
        wait_until_tab_active(stravis, 'Operation')

        # (This 20s sleep was in your version; kept as-is. Want me to swap it for a waiter?)
        time.sleep(20)

        # Change the language
        press_e(root_for_waits=stravis)

        # 16) press save
        switch_ribbon_tab(stravis, 'Operation')
        wait_until_tab_active(stravis, 'Operation')
        click_save_as_excel(stravis)

        click_save_as_excel(stravis)          # your earlier helper
        click_save_as_tree_item('Downloads')  # selects the Downloads folder
        time.sleep(1)
        for i in range(4):
            ui.SendKeys('{TAB}')
            time.sleep(0.1)
        ui.SendKeys('{ENTER}')

        # 17) press close
        switch_ribbon_tab(stravis, 'Operation')
        click_operation_close(stravis)

        for i in range(4):
            ui.SendKeys('{TAB}')
            time.sleep(0.1)
        ui.SendKeys('{DOWN}')
    
    print("Downloading Complete.")

    main()
if __name__ == '__main__':


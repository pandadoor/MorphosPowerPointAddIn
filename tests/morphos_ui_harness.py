import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import uuid

import pythoncom
import win32com.client
from pywinauto import Desktop
from pywinauto.keyboard import send_keys


POWERPOINT_WINDOW_REGEX = r".*PowerPoint$"
MORPHOS_PROG_ID = "MorphosPowerPointAddIn"


def log(message):
    print(message, flush=True)


def wait_for(description, predicate, timeout=20.0, interval=0.25):
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            value = predicate()
            if value:
                return value
        except Exception as error:  # pragma: no cover - UI polling
            last_error = error

        time.sleep(interval)

    if last_error is not None:
        raise RuntimeError(f"{description} timed out. Last error: {last_error}")

    raise RuntimeError(f"{description} timed out.")


def get_powerpoint_application():
    try:
        return win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        application = win32com.client.Dispatch("PowerPoint.Application")
        try:
            application.Visible = True
        except Exception:
            pass
        return application


def close_presentations(application):
    for index in range(application.Presentations.Count, 0, -1):
        presentation = application.Presentations.Item(index)
        try:
            presentation.Close()
        except Exception:
            pass


def try_get_single_open_presentation(application):
    try:
        if application.Presentations.Count != 1:
            return None
    except Exception:
        return None

    try:
        return application.ActivePresentation
    except Exception:
        try:
            return application.Presentations.Item(1)
        except Exception:
            return None


def open_temp_copy(application, presentation_path):
    _ = application
    application = get_powerpoint_application()
    source_path = os.path.abspath(presentation_path)
    if not os.path.exists(source_path):
        raise FileNotFoundError(source_path)

    temp_path = os.path.join(
        tempfile.gettempdir(),
        f"morphos-ui-{uuid.uuid4().hex}{os.path.splitext(source_path)[1]}",
    )
    shutil.copy2(source_path, temp_path)

    close_presentations(application)
    application.Presentations.Open(temp_path, False, False, True)
    presentation = wait_for(
        "presentation activation",
        lambda: try_get_single_open_presentation(application),
        timeout=30.0,
    )
    try:
        if presentation.Windows.Count > 0:
            presentation.Windows.Item(1).Activate()
    except Exception:
        pass

    return application, presentation, temp_path


def ensure_morphos_connected(application):
    _ = application
    command = (
        "$pp = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); "
        "$found = $false; "
        "for ($i = 1; $i -le $pp.COMAddIns.Count; $i++) { "
        "  $addin = $pp.COMAddIns.Item($i); "
        f"  if ($addin -and $addin.ProgId -eq '{MORPHOS_PROG_ID}') {{ "
        "    $found = $true; "
        "    if (-not $addin.Connect) { $addin.Connect = $true; Start-Sleep -Milliseconds 500 } "
        "    if ($addin.Connect) { exit 0 } "
        "  } "
        "} "
        "if (-not $found) { throw 'Morphos COM add-in was not registered in PowerPoint.' } "
        "throw 'Morphos COM add-in did not connect.'"
    )
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", command],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        stderr = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(stderr or "Morphos COM add-in did not connect.")

    time.sleep(1.5)


def get_powerpoint_window(presentation_name=None):
    def _find_window():
        windows = Desktop(backend="uia").windows(title_re=POWERPOINT_WINDOW_REGEX)
        if presentation_name:
            for window in windows:
                if presentation_name.lower() in window.window_text().lower():
                    return window

        powerpoint_windows = [window for window in windows if ".ppt" in window.window_text().lower()]
        return powerpoint_windows[0] if powerpoint_windows else (windows[0] if windows else None)

    return wait_for("PowerPoint window", _find_window, timeout=30.0)


def try_find_top_level_window(title):
    matches = Desktop(backend="uia").windows(title=title)
    return matches[0] if matches else None


def try_find_dialog_window(container, title):
    try:
        matches = container.descendants(title=title, control_type="Window")
        if matches:
            return matches[0]
    except Exception:
        pass

    return try_find_top_level_window(title)


def wait_for_dialog_window(container, title, timeout=15.0):
    return wait_for(
        f"window {title}",
        lambda: try_find_dialog_window(container, title),
        timeout=timeout,
    )


def find_descendant(container, title=None, control_type=None):
    def _find():
        matches = container.descendants(title=title, control_type=control_type)
        return matches[0] if matches else None

    return wait_for(
        f"{control_type or 'control'} {title or ''}".strip(),
        _find,
        timeout=15.0,
    )


def click_element(element):
    wrapper = element.wrapper_object() if hasattr(element, "wrapper_object") else element

    try:
        if getattr(wrapper.element_info, "control_type", "") in ("TabItem", "TreeItem") and hasattr(wrapper, "select"):
            wrapper.select()
            return
    except Exception:
        pass

    try:
        wrapper.invoke()
        return
    except Exception:
        pass

    try:
        wrapper.click_input()
        return
    except Exception:
        pass

    wrapper.click()


def set_toggle_button_state(button, should_be_on):
    wrapper = button.wrapper_object() if hasattr(button, "wrapper_object") else button

    try:
        current_state = wrapper.get_toggle_state()
    except Exception:
        current_state = None

    if current_state is not None:
        desired_state = 1 if should_be_on else 0
        if current_state == desired_state:
            return

        try:
            wrapper.toggle()
            return
        except Exception:
            pass

    click_element(wrapper)


def ensure_morphos_pane(window, reset_visibility=False):
    window.set_focus()

    morphos_tab = find_descendant(window, title="Morphos", control_type="TabItem")
    click_element(morphos_tab)
    time.sleep(0.5)

    open_inspector_button = find_descendant(window, title="Open Inspector", control_type="Button")

    if reset_visibility:
        pane = try_find_morphos_pane(window)
        if pane is not None:
            set_toggle_button_state(open_inspector_button, False)
            try:
                wait_for("Morphos pane hidden", lambda: try_find_morphos_pane(window) is None, timeout=3.0)
            except RuntimeError:
                pass

    if try_find_morphos_pane(window) is None:
        set_toggle_button_state(open_inspector_button, True)

    pane = wait_for("Morphos pane", lambda: try_find_morphos_pane(window), timeout=15.0)
    wait_for(
        "Morphos pane content",
        lambda: pane.descendants(title="Refresh", control_type="Button"),
        timeout=20.0,
    )
    return wait_for("Morphos pane", lambda: try_find_morphos_pane(window), timeout=15.0)


def try_find_morphos_pane(window):
    matches = window.descendants(title="Morphos", control_type="Pane")
    for match in matches:
        try:
            if match.descendants(title="Refresh", control_type="Button"):
                return match
        except Exception:
            pass

    return matches[0] if matches else None


def extract_dialog_message(dialog):
    if dialog is None:
        return "Morphos displayed an error dialog."

    lines = []
    seen = set()
    for element in dialog.descendants(control_type="Text"):
        text = (element.window_text() or "").strip()
        if not text or text == dialog.window_text() or text in seen:
            continue

        seen.add(text)
        lines.append(text)

    return " ".join(lines) if lines else "Morphos displayed an error dialog."


def dismiss_dialog(dialog, button_title="OK"):
    if dialog is None:
        return

    try:
        click_element(find_descendant(dialog, title=button_title, control_type="Button"))
    except Exception:
        pass


def wait_for_dialog_resolution(container, dialog_title, timeout=15.0):
    outcome, dialog = wait_for(
        f"{dialog_title} submission",
        lambda: (
            ("error", try_find_dialog_window(container, "Morphos"))
            if try_find_dialog_window(container, "Morphos") is not None
            else ("closed", None)
            if try_find_dialog_window(container, dialog_title) is None
            else None
        ),
        timeout=timeout,
    )

    if outcome == "error":
        message = extract_dialog_message(dialog)
        dismiss_dialog(dialog)
        raise RuntimeError(message)


def ensure_pane_tab_selected(pane, tab_title):
    marker_map = {
        "Fonts": ("Font inventory", "Text"),
        "Colors": ("Color inventory", "Text"),
        "Home": ("At a glance", "Text"),
    }

    marker_title, marker_type = marker_map.get(tab_title, (tab_title, "Text"))
    tab = find_descendant(pane, title=tab_title, control_type="TabItem")

    last_error = None
    for _ in range(3):
        click_element(tab)
        time.sleep(0.5)
        try:
            wait_for_descendant(pane, title=marker_title, control_type=marker_type, timeout=5.0)
            return
        except RuntimeError as error:
            last_error = error

    if last_error is not None:
        raise last_error


def get_tree_items(pane):
    return pane.descendants(control_type="TreeItem")


def wait_for_tree_items(pane, minimum_count=1, timeout=20.0):
    return wait_for(
        "tree items",
        lambda: get_tree_items(pane) if len(get_tree_items(pane)) >= minimum_count else None,
        timeout=timeout,
    )


def get_button(container, title):
    return find_descendant(container, title=title, control_type="Button")


def wait_for_descendant(container, title=None, control_type=None, timeout=15.0):
    return wait_for(
        f"{control_type or 'control'} {title or ''}".strip(),
        lambda: container.descendants(title=title, control_type=control_type),
        timeout=timeout,
    )[0]


def click_until_descendant(container, button, title=None, control_type=None, attempts=3, timeout=5.0):
    last_error = None
    for _ in range(attempts):
        click_element(button)
        time.sleep(0.4)
        try:
            return wait_for_descendant(container, title=title, control_type=control_type, timeout=timeout)
        except RuntimeError as error:
            last_error = error

    if last_error is not None:
        raise last_error

    raise RuntimeError(f"Could not open {title or control_type or 'target'} after clicking the button.")


def ensure_tree_item_expanded(tree_item):
    buttons = tree_item.descendants(control_type="Button")
    if not buttons:
        return

    toggle_button = buttons[0]
    try:
        if toggle_button.get_toggle_state() == 1:
            return
    except Exception:
        pass

    try:
        toggle_button.toggle()
    except Exception:
        click_element(toggle_button)


def try_get_state(application):
    state = {
        "presentations": 0,
        "active_full_name": "",
        "windows": 0,
        "saved": None,
    }

    try:
        state["presentations"] = application.Presentations.Count
    except Exception:
        return state

    if state["presentations"] <= 0:
        return state

    try:
        presentation = application.Presentations.Item(1)
        state["active_full_name"] = presentation.FullName or ""
        state["windows"] = presentation.Windows.Count
        state["saved"] = int(presentation.Saved)
    except Exception:
        pass

    return state


def assert_presentation_window_alive(application, expected_path):
    state = try_get_state(application)
    if state["presentations"] != 1:
        raise RuntimeError(f"Expected one open presentation after UI flow, found {state['presentations']}.")

    if state["windows"] < 1:
        raise RuntimeError(
            "PowerPoint lost its presentation window after the replace action. "
            f"State: {state}"
        )

    normalized_expected = os.path.normcase(os.path.abspath(expected_path))
    normalized_actual = os.path.normcase(os.path.abspath(state["active_full_name"]))
    if normalized_actual != normalized_expected:
        raise RuntimeError(
            "PowerPoint reattached to an unexpected presentation after the replace action. "
            f"Expected '{expected_path}', actual '{state['active_full_name']}'."
        )


def run_autoscan(application, presentation_path):
    application, _, temp_path = open_temp_copy(application, presentation_path)
    window = get_powerpoint_window(os.path.basename(temp_path))
    pane = ensure_morphos_pane(window, reset_visibility=True)
    ensure_pane_tab_selected(pane, "Fonts")
    tree_items = wait_for_tree_items(pane, minimum_count=1, timeout=20.0)
    log(f"Auto-scan loaded {len(tree_items)} visible tree items for {os.path.basename(temp_path)}.")
    return temp_path


def run_open_font_dialog(application, presentation_path):
    application, _, temp_path = open_temp_copy(application, presentation_path)
    window = get_powerpoint_window(os.path.basename(temp_path))
    pane = ensure_morphos_pane(window, reset_visibility=True)
    ensure_pane_tab_selected(pane, "Fonts")
    tree_items = wait_for_tree_items(pane, minimum_count=1, timeout=20.0)
    click_element(tree_items[0])
    time.sleep(0.4)

    replace_button = get_button(pane, "Replace font")
    click_until_descendant(window, replace_button, title="Selected font", control_type="Text", attempts=3, timeout=5.0)
    wait_for_dialog_window(window, "Replace Font", timeout=10.0)
    log(f"Replace Font dialog opened for {os.path.basename(temp_path)}.")
    return temp_path


def run_cycle_tabs(application, presentation_path):
    application, _, temp_path = open_temp_copy(application, presentation_path)
    window = get_powerpoint_window(os.path.basename(temp_path))
    pane = ensure_morphos_pane(window, reset_visibility=True)

    for tab_title in ("Home", "Fonts", "Colors", "Home"):
        ensure_pane_tab_selected(pane, tab_title)
        time.sleep(0.5)

    log(f"Tab navigation succeeded for {os.path.basename(temp_path)}.")
    return temp_path


def run_replace_font(application, presentation_path):
    application, _, temp_path = open_temp_copy(application, presentation_path)
    window = get_powerpoint_window(os.path.basename(temp_path))
    pane = ensure_morphos_pane(window, reset_visibility=True)
    ensure_pane_tab_selected(pane, "Fonts")
    tree_items = wait_for_tree_items(pane, minimum_count=1, timeout=20.0)
    click_element(tree_items[0])
    time.sleep(0.4)

    replace_button = get_button(pane, "Replace font")
    click_until_descendant(window, replace_button, title="Selected font", control_type="Text", attempts=3, timeout=5.0)
    wait_for_dialog_window(window, "Replace Font", timeout=10.0)
    click_element(wait_for_descendant(window, title="Replace", control_type="Button", timeout=15.0))
    wait_for_dialog_resolution(window, "Replace Font", timeout=15.0)
    time.sleep(2.0)

    assert_presentation_window_alive(application, temp_path)
    log(f"Font replace user flow kept PowerPoint attached to {os.path.basename(temp_path)}.")
    return temp_path


def run_open_color_dialog(application, presentation_path):
    application, _, temp_path = open_temp_copy(application, presentation_path)
    window = get_powerpoint_window(os.path.basename(temp_path))
    pane = ensure_morphos_pane(window, reset_visibility=True)
    ensure_pane_tab_selected(pane, "Colors")
    tree_items = wait_for_tree_items(pane, minimum_count=1, timeout=20.0)
    color_group = next(
        (item for item in tree_items if item.window_text().endswith("ColorGroupNodeViewModel")),
        None,
    )
    if color_group is None:
        raise RuntimeError("Morphos did not surface a selectable color group in the pane.")

    ensure_tree_item_expanded(color_group)
    time.sleep(0.5)

    color_item = wait_for(
        "selectable color node",
        lambda: next(
            (
                item
                for item in color_group.descendants(control_type="TreeItem")
                if item.window_text().endswith("ColorNodeViewModel")
            ),
            None,
        ),
        timeout=15.0,
    )
    if color_item is None:
        raise RuntimeError("Morphos did not surface a selectable color node in the pane.")

    click_element(color_item)
    time.sleep(0.4)

    replace_button = get_button(pane, "Replace color")
    click_until_descendant(window, replace_button, title="Pick custom color", control_type="Button", attempts=3, timeout=5.0)
    wait_for_dialog_window(window, "Replace Color", timeout=10.0)
    log(f"Replace Color dialog opened for {os.path.basename(temp_path)}.")
    return temp_path


def run_replace_color(application, presentation_path):
    application, _, temp_path = open_temp_copy(application, presentation_path)
    window = get_powerpoint_window(os.path.basename(temp_path))
    pane = ensure_morphos_pane(window, reset_visibility=True)
    ensure_pane_tab_selected(pane, "Colors")
    tree_items = wait_for_tree_items(pane, minimum_count=1, timeout=20.0)
    color_group = next(
        (item for item in tree_items if item.window_text().endswith("ColorGroupNodeViewModel")),
        None,
    )
    if color_group is None:
        raise RuntimeError("Morphos did not surface a selectable color group in the pane.")

    ensure_tree_item_expanded(color_group)
    time.sleep(0.5)

    color_item = wait_for(
        "selectable color node",
        lambda: next(
            (
                item
                for item in color_group.descendants(control_type="TreeItem")
                if item.window_text().endswith("ColorNodeViewModel")
            ),
            None,
        ),
        timeout=15.0,
    )
    if color_item is None:
        raise RuntimeError("Morphos did not surface a selectable color node in the pane.")

    click_element(color_item)
    time.sleep(0.4)

    replace_button = get_button(pane, "Replace color")
    click_until_descendant(window, replace_button, title="Pick custom color", control_type="Button", attempts=3, timeout=5.0)

    combo = wait_for_descendant(window, control_type="ComboBox", timeout=15.0)
    combo = combo.wrapper_object() if hasattr(combo, "wrapper_object") else combo
    combo.select(1)
    time.sleep(0.4)

    wait_for_dialog_window(window, "Replace Color", timeout=10.0)
    click_element(wait_for_descendant(window, title="Replace", control_type="Button", timeout=15.0))
    wait_for_dialog_resolution(window, "Replace Color", timeout=15.0)
    time.sleep(2.0)

    assert_presentation_window_alive(application, temp_path)
    log(f"Color replace user flow kept PowerPoint attached to {os.path.basename(temp_path)}.")
    return temp_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--presentation", required=True)
    parser.add_argument(
        "--mode",
        default="all",
        choices=["autoscan", "cycle-tabs", "open-font-dialog", "replace-font", "open-color-dialog", "replace-color", "all"],
    )
    args = parser.parse_args()

    pythoncom.CoInitialize()
    try:
        application = get_powerpoint_application()
        try:
            application.Visible = True
        except Exception:
            pass
        ensure_morphos_connected(application)

        if args.mode in ("autoscan", "all"):
            run_autoscan(application, args.presentation)

        if args.mode in ("cycle-tabs", "all"):
            run_cycle_tabs(application, args.presentation)

        if args.mode in ("open-font-dialog", "all"):
            run_open_font_dialog(application, args.presentation)

        if args.mode in ("replace-font", "all"):
            run_replace_font(application, args.presentation)

        if args.mode in ("open-color-dialog", "all"):
            run_open_color_dialog(application, args.presentation)

        if args.mode in ("replace-color", "all"):
            run_replace_color(application, args.presentation)
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        traceback.print_exc()
        print(str(error), file=sys.stderr, flush=True)
        sys.exit(1)

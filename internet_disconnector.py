import hashlib
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
import base64

from collections import OrderedDict
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

psutil = None

CONFIG_PATH = Path(__file__).with_suffix(".config.json")
FALLBACK_CONFIG_PATH = Path.home() / ".internet_disconnector.config.json"
BLOCK_RULE_PREFIX = "InternetDisconnector Block"
ALLOW_RULE_PREFIX = "InternetDisconnector Allow"
BLOCK_ALL_RULE_NAME = "InternetDisconnector Block All"
STATUS_BLOCKED = "Blocked"

CONFIG_DEFAULTS = {
    "ui_mode": "old",
    "hide_console": True,
    "auto_refresh": False,
    "refresh_interval": 5,
    "auto_reconnect_seconds": 60,
    "saved_reconnect_timers": {},
    "disabled_reconnect_timers": [],
    "preset_autosave": False,
    "only_active_default": False,
    "show_block_status": True,
    "show_info_notifications": True,
    "ask_for_admin_on_startup": True,
    "warn_when_not_admin": True,
    "confirm_block_all_except": True,
    "remember_window_geometry": True,
    "max_cached_icons": 192,
    "hotkeys_enabled": True,
    "hotkey_focus_search": "CTRL+F",
    "hotkey_refresh": "VK_F5",
    "hotkey_block_selected": "CTRL+B",
    "hotkey_unblock_selected": "CTRL+U",
    "hotkey_unblock_all": "CTRL+SHIFT+U",
    "hotkey_toggle_selected": "CTRL+VK_RETURN",
    "about_github_url": "base64:aHR0cHM6Ly9naXRodWIuY29tL3NpcmFwYXRwaXBvL0Etc2VsZWN0YWJsZS1kaXNjb25uZWN0",
    "window_geometry": "460x360",
    "settings_window_geometry": "520x620",
}

MODIFIER_STATE_MASKS = {
    "Ctrl": 0x0004,
    "Alt": 0x0008,
    "Shift": 0x0001,
}


def _find_writable_config_path() -> Path:
    for candidate in (CONFIG_PATH, FALLBACK_CONFIG_PATH):
        try:
            candidate.parent.mkdir(parents=True, exist_ok=True)
            if candidate.exists():
                with open(candidate, "a", encoding="utf-8"):
                    pass
            else:
                candidate.write_text("{}", encoding="utf-8")
                candidate.unlink()
            return candidate
        except Exception:
            continue
    return CONFIG_PATH


CONFIG_PATH = _find_writable_config_path()


def get_preset_dir() -> Path:
    preset_dir = CONFIG_PATH.parent / "presets"
    preset_dir.mkdir(parents=True, exist_ok=True)
    return preset_dir


def try_import_psutil() -> bool:
    global psutil
    try:
        import psutil as _ps

        psutil = _ps
        return True
    except ImportError:
        psutil = None
        return False


def _write_config_file(target: Path, config_data: dict) -> bool:
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(config_data, indent=2), encoding="utf-8")
        return True
    except Exception:
        return False


def _normalize_loaded_config(raw: dict | None) -> dict:
    merged = dict(raw or {})
    for key, value in CONFIG_DEFAULTS.items():
        if key not in merged:
            if isinstance(value, dict):
                merged[key] = value.copy()
            elif isinstance(value, list):
                merged[key] = list(value)
            else:
                merged[key] = value
    return merged


def load_config() -> dict:
    global CONFIG_PATH

    if not CONFIG_PATH.exists():
        defaults = dict(CONFIG_DEFAULTS)
        if not _write_config_file(CONFIG_PATH, defaults) and CONFIG_PATH != FALLBACK_CONFIG_PATH:
            if _write_config_file(FALLBACK_CONFIG_PATH, defaults):
                CONFIG_PATH = FALLBACK_CONFIG_PATH
        return defaults

    try:
        raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        if not isinstance(raw, dict):
            raise ValueError("Config file must contain a JSON object.")
        normalized = _normalize_loaded_config(raw)
        if normalized != raw:
            save_config(normalized)
        return normalized
    except Exception:
        root = tk.Tk()
        root.withdraw()
        choice = messagebox.askyesnocancel(
            "Config error",
            "The configuration file is unreadable.\n\n"
            "Yes = reset it to defaults\n"
            "No = continue with defaults only\n"
            "Cancel = close the application",
        )
        root.destroy()

        if choice is None:
            sys.exit(0)

        defaults = dict(CONFIG_DEFAULTS)
        if choice:
            save_config(defaults)
        return defaults


def save_config(config_data: dict) -> None:
    global CONFIG_PATH

    normalized = _normalize_loaded_config(config_data)
    if _write_config_file(CONFIG_PATH, normalized):
        return
    if CONFIG_PATH != FALLBACK_CONFIG_PATH and _write_config_file(FALLBACK_CONFIG_PATH, normalized):
        CONFIG_PATH = FALLBACK_CONFIG_PATH


def install_package(package: str) -> tuple[bool, str]:
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--user", package],
            capture_output=True,
            text=True,
            timeout=600,
            shell=False,
        )
        output = "\n".join(part for part in (result.stdout.strip(), result.stderr.strip()) if part)
        return result.returncode == 0, output or "No output."
    except Exception as error:
        return False, str(error)


def ensure_psutil() -> bool:
    if try_import_psutil():
        return True

    root = tk.Tk()
    root.withdraw()
    install_now = messagebox.askyesno(
        "psutil missing",
        "psutil is required to list running applications.\n\n"
        "Install it now?",
    )
    root.destroy()

    if not install_now:
        return False

    success, output = install_package("psutil")
    if not success:
        messagebox.showerror("Install failed", "Could not install psutil.\n\n" + output)
        return False

    if try_import_psutil():
        return True

    messagebox.showerror("Import failed", "psutil was installed but could not be imported.")
    return False


def is_admin() -> bool:
    try:
        return os.getuid() == 0  # type: ignore[attr-defined]
    except AttributeError:
        import ctypes

        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except Exception:
            return False


def relaunch_as_admin() -> bool:
    try:
        import ctypes

        result = ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.executable,
            f'"{os.path.abspath(__file__)}"',
            None,
            1,
        )
        return int(result) > 32
    except Exception:
        return False


def _normalize_path(value: str) -> str:
    return os.path.normcase(os.path.abspath(value)).replace("/", "\\")


def _build_rule_suffix(exe_path: str) -> str:
    normalized = _normalize_path(exe_path)
    digest = hashlib.sha1(normalized.encode("utf-8")).hexdigest()[:10]
    return f"{os.path.basename(normalized)} [{digest}]"


def build_firewall_rule_name(exe_path: str) -> str:
    return f"{BLOCK_RULE_PREFIX} - {_build_rule_suffix(exe_path)}"


def build_allow_rule_name(exe_path: str) -> str:
    return f"{ALLOW_RULE_PREFIX} - {_build_rule_suffix(exe_path)}"


def _legacy_block_rule_name(exe_path: str) -> str:
    return f"BlockApp - {os.path.basename(exe_path)}"


def _legacy_allow_rule_name(exe_path: str) -> str:
    return f"AllowApp - {os.path.basename(exe_path)}"


def _candidate_block_rule_names(exe_path: str) -> list[str]:
    return [build_firewall_rule_name(exe_path), _legacy_block_rule_name(exe_path)]


def _candidate_allow_rule_names(exe_path: str) -> list[str]:
    return [build_allow_rule_name(exe_path), _legacy_allow_rule_name(exe_path)]


def run_netsh(args: list[str], timeout: int = 60) -> tuple[int, str, str]:
    try:
        result = subprocess.run(
            ["netsh", "advfirewall", "firewall", *args],
            capture_output=True,
            text=True,
            timeout=timeout,
            shell=False,
        )
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as error:
        return 1, "", str(error)


def run_powershell(script: str, timeout: int = 60) -> tuple[int, str, str]:
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", script],
            capture_output=True,
            text=True,
            timeout=timeout,
            shell=False,
        )
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as error:
        return 1, "", str(error)


def _combine_output(stdout: str, stderr: str) -> str:
    return "\n".join(part for part in (stdout.strip(), stderr.strip()) if part)


def _output_mentions_path(output: str, exe_path: str) -> bool:
    return _normalize_path(exe_path).lower() in output.replace("/", "\\").lower()


def firewall_rule_exists(rule_name: str, exe_path: str | None = None) -> bool:
    return_code, stdout, stderr = run_netsh(["show", "rule", f"name={rule_name}"])
    if return_code != 0:
        return False

    combined = _combine_output(stdout, stderr)
    if not combined or "No rules match" in combined:
        return False
    if exe_path is None:
        return True
    return _output_mentions_path(combined, exe_path)


def get_blocked_programs_snapshot() -> set[str] | None:
    script = r"""
$programs = @(
    Get-NetFirewallRule -PolicyStore ActiveStore -ErrorAction SilentlyContinue |
    Where-Object {
        $_.DisplayName -like 'InternetDisconnector Block*' -or
        $_.DisplayName -like 'BlockApp -*'
    } |
    Get-NetFirewallApplicationFilter -ErrorAction SilentlyContinue |
    Where-Object { $_.Program -and $_.Program -ne 'Any' } |
    Select-Object -ExpandProperty Program -Unique
)
ConvertTo-Json -Compress -InputObject $programs
"""
    return_code, stdout, stderr = run_powershell(script, timeout=15)
    if return_code != 0:
        return None

    output = stdout.strip()
    if not output:
        return set()

    try:
        parsed = json.loads(output)
    except json.JSONDecodeError:
        return None

    if isinstance(parsed, str):
        programs = [parsed]
    elif isinstance(parsed, list):
        programs = [value for value in parsed if isinstance(value, str)]
    else:
        return set()

    return {_normalize_path(path) for path in programs if path and path != "Any"}


def list_blocked_programs() -> list[str] | None:
    blocked_programs = get_blocked_programs_snapshot()
    if blocked_programs is None:
        return None
    return sorted(blocked_programs, key=lambda path: (os.path.basename(path).lower(), path.lower()))


def is_blocked(exe_path: str) -> bool:
    return any(firewall_rule_exists(rule_name, exe_path) for rule_name in _candidate_block_rule_names(exe_path))


def _delete_program_rules(rule_names: list[str], exe_path: str) -> tuple[bool, str]:
    deleted_any = False
    messages: list[str] = []

    for rule_name in dict.fromkeys(rule_names):
        return_code, stdout, stderr = run_netsh(["delete", "rule", f"name={rule_name}", f"program={exe_path}"])
        combined = _combine_output(stdout, stderr)
        if return_code == 0:
            deleted_any = True
            if combined:
                messages.append(combined)
        elif combined and "No rules match" not in combined:
            messages.append(combined)

    if deleted_any:
        return True, "Removed matching firewall rules."
    if messages:
        return False, "\n".join(messages)
    return True, "No matching firewall rules were found."


def block_app(exe_path: str) -> tuple[bool, str]:
    if is_blocked(exe_path):
        return True, f"{os.path.basename(exe_path)} is already blocked."

    return_code, stdout, stderr = run_netsh(
        [
            "add",
            "rule",
            f"name={build_firewall_rule_name(exe_path)}",
            "dir=out",
            "action=block",
            f"program={exe_path}",
            "enable=yes",
        ]
    )
    if return_code == 0:
        return True, stdout or "Blocked successfully."
    return False, _combine_output(stdout, stderr) or f"netsh exited with {return_code}"


def unblock_app(exe_path: str) -> tuple[bool, str]:
    return _delete_program_rules(_candidate_block_rule_names(exe_path), exe_path)


def allow_app(exe_path: str) -> tuple[bool, str]:
    return_code, stdout, stderr = run_netsh(
        [
            "add",
            "rule",
            f"name={build_allow_rule_name(exe_path)}",
            "dir=out",
            "action=allow",
            f"program={exe_path}",
            "enable=yes",
        ]
    )
    if return_code == 0:
        return True, stdout or "Allowed successfully."
    return False, _combine_output(stdout, stderr) or f"netsh exited with {return_code}"


def remove_allow_rule(exe_path: str) -> tuple[bool, str]:
    return _delete_program_rules(_candidate_allow_rule_names(exe_path), exe_path)


def block_all() -> tuple[bool, str]:
    return_code, stdout, stderr = run_netsh(
        [
            "add",
            "rule",
            f"name={BLOCK_ALL_RULE_NAME}",
            "dir=out",
            "action=block",
            "enable=yes",
        ]
    )
    if return_code == 0:
        return True, stdout or "Blocked all outbound traffic."
    return False, _combine_output(stdout, stderr) or f"netsh exited with {return_code}"


def unblock_all() -> tuple[bool, str]:
    return_code, stdout, stderr = run_netsh(["delete", "rule", f"name={BLOCK_ALL_RULE_NAME}"])
    if return_code == 0:
        return True, stdout or "Stopped blocking all outbound traffic."

    combined = _combine_output(stdout, stderr)
    if "No rules match" in combined:
        return True, "Block-all rule was not present."
    return False, combined or f"netsh exited with {return_code}"


def find_running_apps(only_active: bool = False) -> list[dict]:
    if psutil is None:
        raise RuntimeError("psutil is required.")

    active_pids: set[int] | None = None
    if only_active:
        active_pids = set()
        try:
            for conn in psutil.net_connections(kind="inet"):
                if not conn.pid:
                    continue
                if conn.raddr or conn.status in {
                    "ESTABLISHED",
                    "SYN_SENT",
                    "SYN_RECV",
                    "FIN_WAIT1",
                    "FIN_WAIT2",
                    "CLOSE_WAIT",
                    "LAST_ACK",
                    "CLOSING",
                    "TIME_WAIT",
                }:
                    active_pids.add(conn.pid)
        except Exception:
            active_pids = None

    apps: dict[str, dict] = {}
    for proc in psutil.process_iter(["pid", "name", "exe"]):
        try:
            exe_path = proc.info.get("exe") or ""
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            continue

        if not exe_path:
            continue

        if only_active and active_pids is not None and proc.info.get("pid") not in active_pids:
            continue

        normalized = _normalize_path(exe_path)
        if normalized in apps:
            continue

        apps[normalized] = {
            "name": proc.info.get("name") or os.path.basename(exe_path),
            "exe": exe_path,
            "pid": proc.info.get("pid"),
            "blocked": None,
        }

    return sorted(apps.values(), key=lambda item: (item["name"].lower(), item["exe"].lower()))


def load_app_statuses(items: list[dict], show_status: bool) -> list[dict]:
    if not show_status:
        return items

    blocked_programs = get_blocked_programs_snapshot()
    if blocked_programs is not None:
        for item in items:
            item["blocked"] = _normalize_path(item["exe"]) in blocked_programs
        return items

    for item in items:
        try:
            item["blocked"] = is_blocked(item["exe"])
        except Exception:
            item["blocked"] = None
    return items


def is_valid_geometry(value: str) -> bool:
    parts = value.lower().split("x")
    if len(parts) != 2:
        return False
    try:
        width = int(parts[0])
        height = int(parts[1])
    except ValueError:
        return False
    return width > 0 and height > 0


def clamp_geometry(value: str, fallback: str, min_width: int, min_height: int) -> str:
    target = value if is_valid_geometry(value) else fallback
    width_text, height_text = target.lower().split("x")
    width = max(int(width_text), min_width)
    height = max(int(height_text), min_height)
    return f"{width}x{height}"


def _normalize_hotkey_key_name(keysym: str) -> str | None:
    if not keysym:
        return None

    ignored = {
        "Control_L",
        "Control_R",
        "Shift_L",
        "Shift_R",
        "Alt_L",
        "Alt_R",
        "Meta_L",
        "Meta_R",
        "Super_L",
        "Super_R",
    }
    if keysym in ignored:
        return None

    special_names = {
        "Return": "Enter",
        "Escape": "Esc",
        "space": "Space",
        "BackSpace": "Backspace",
        "Delete": "Delete",
        "Tab": "Tab",
        "Home": "Home",
        "End": "End",
        "Prior": "Page Up",
        "Next": "Page Down",
        "Insert": "Insert",
        "Up": "Up",
        "Down": "Down",
        "Left": "Left",
        "Right": "Right",
    }
    if keysym in special_names:
        return special_names[keysym]

    if len(keysym) == 1:
        return keysym.upper()

    if keysym.startswith("F") and keysym[1:].isdigit():
        return keysym.upper()

    return keysym.replace("_", " ").title()


VK_TO_TK_KEY = {
    "VK_RETURN": "Return",
    "VK_ENTER": "Return",
    "VK_ESCAPE": "Escape",
    "VK_ESC": "Escape",
    "VK_SPACE": "space",
    "VK_TAB": "Tab",
    "VK_BACK": "BackSpace",
    "VK_BACKSPACE": "BackSpace",
    "VK_DELETE": "Delete",
    "VK_INSERT": "Insert",
    "VK_HOME": "Home",
    "VK_END": "End",
    "VK_LEFT": "Left",
    "VK_RIGHT": "Right",
    "VK_UP": "Up",
    "VK_DOWN": "Down",
    "VK_PRIOR": "Prior",
    "VK_PAGEUP": "Prior",
    "VK_NEXT": "Next",
    "VK_PAGEDOWN": "Next",
    "VK_SNAPSHOT": "Print",
    "VK_PRINT": "Print",
    "VK_HELP": "Help",
}


def hotkey_config_to_sequence(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    if value.startswith("<") and value.endswith(">"):
        return value

    compact = value.replace(" ", "")
    if "+" not in compact and len(compact) == 1 and compact.isprintable():
        return compact.lower()

    modifiers: list[str] = []
    key_token = ""
    for raw_token in [part for part in compact.split("+") if part]:
        token = raw_token.upper()
        if token in {"CTRL", "CONTROL", "VK_CONTROL", "VK_LCONTROL", "VK_RCONTROL"}:
            if "Control" not in modifiers:
                modifiers.append("Control")
            continue
        if token in {"SHIFT", "VK_SHIFT", "VK_LSHIFT", "VK_RSHIFT"}:
            if "Shift" not in modifiers:
                modifiers.append("Shift")
            continue
        if token in {"ALT", "VK_MENU", "VK_LMENU", "VK_RMENU", "MENU"}:
            if "Alt" not in modifiers:
                modifiers.append("Alt")
            continue
        key_token = token

    if not key_token:
        return ""

    if key_token in VK_TO_TK_KEY:
        key_name = VK_TO_TK_KEY[key_token]
    elif key_token.startswith("VK_F") and key_token[4:].isdigit():
        key_name = "F" + key_token[4:]
    elif key_token.startswith("VK_") and len(key_token) == 4 and key_token[-1].isalnum():
        key_name = key_token[-1]
    elif len(key_token) == 1 and key_token.isalnum():
        key_name = key_token
    elif key_token.startswith("F") and key_token[1:].isdigit():
        key_name = key_token
    else:
        return ""

    if len(key_name) == 1:
        key_name = key_name.lower()

    if modifiers or len(key_name) != 1:
        return "<" + "-".join([*modifiers, key_name]) + ">"
    return key_name


def hotkey_sequence_to_display(sequence: str) -> str:
    sequence = hotkey_config_to_sequence(sequence)
    if not sequence:
        return ""

    if sequence.startswith("<") and sequence.endswith(">"):
        inner = sequence[1:-1]
        parts = [part for part in inner.split("-") if part]
        if not parts:
            return ""

        display_parts: list[str] = []
        for part in parts[:-1]:
            if part == "Control":
                display_parts.append("Ctrl")
            elif part == "Shift":
                display_parts.append("Shift")
            elif part == "Alt":
                display_parts.append("Alt")
            else:
                display_parts.append(part)

        key_name = _normalize_hotkey_key_name(parts[-1]) or parts[-1]
        display_parts.append(key_name)
        return " + ".join(display_parts)

    return sequence.upper() if len(sequence) == 1 else sequence


def hotkey_event_to_config(event) -> str | None:
    keysym = event.keysym
    ignored = {
        "Control_L",
        "Control_R",
        "Shift_L",
        "Shift_R",
        "Alt_L",
        "Alt_R",
        "Meta_L",
        "Meta_R",
        "Super_L",
        "Super_R",
    }
    if not keysym or keysym in ignored:
        return None

    special_tokens = {
        "Return": "VK_RETURN",
        "Escape": "VK_ESCAPE",
        "space": "VK_SPACE",
        "BackSpace": "VK_BACK",
        "Delete": "VK_DELETE",
        "Tab": "VK_TAB",
        "Home": "VK_HOME",
        "End": "VK_END",
        "Prior": "VK_PRIOR",
        "Next": "VK_NEXT",
        "Insert": "VK_INSERT",
        "Up": "VK_UP",
        "Down": "VK_DOWN",
        "Left": "VK_LEFT",
        "Right": "VK_RIGHT",
        "Print": "VK_SNAPSHOT",
        "Help": "VK_HELP",
    }

    if keysym in special_tokens:
        key_token = special_tokens[keysym]
    elif len(keysym) == 1 and keysym.isprintable():
        key_token = keysym.upper()
    elif keysym.startswith("F") and keysym[1:].isdigit():
        key_token = "VK_" + keysym.upper()
    else:
        normalized = _normalize_hotkey_key_name(keysym)
        if not normalized:
            return None
        compact = normalized.replace(" ", "_").upper()
        key_token = compact if len(compact) == 1 else f"VK_{compact}"

    ctrl_pressed = bool(event.state & MODIFIER_STATE_MASKS["Ctrl"])
    alt_pressed = bool(event.state & MODIFIER_STATE_MASKS["Alt"])
    shift_pressed = bool(event.state & MODIFIER_STATE_MASKS["Shift"])

    modifiers: list[str] = []
    if ctrl_pressed:
        modifiers.append("CTRL")
    # Some Windows/Tk setups report Alt for normal typing in Entry widgets.
    # Treat plain printable keys as plain keys unless Alt is clearly intentional.
    if alt_pressed and (ctrl_pressed or not (len(key_token) == 1 and key_token.isalnum())):
        modifiers.append("ALT")
    if shift_pressed and not (len(key_token) == 1 and key_token.isalpha()):
        modifiers.append("SHIFT")

    return "+".join([*modifiers, key_token]) if modifiers else key_token


def format_duration_label(seconds: int) -> str:
    seconds = max(1, int(seconds))
    if seconds % 60 == 0 and seconds >= 60:
        minutes = seconds // 60
        return f"{minutes} minute" if minutes == 1 else f"{minutes} minutes"
    return f"{seconds} second" if seconds == 1 else f"{seconds} seconds"


def decode_config_url(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    if not value.lower().startswith("base64:"):
        return value
    try:
        encoded = value.split(":", 1)[1].strip()
        return base64.b64decode(encoded).decode("utf-8").strip()
    except Exception:
        return ""


class App(tk.Tk):
    def __init__(self, is_admin_user: bool, config: dict):
        super().__init__()
        self.title("Internet Disconnector")
        self.is_admin = is_admin_user
        self.config_data = config
        self.ui_mode = str(self.config_data.get("ui_mode", "old")).lower()
        self._all_items: list[dict] = []
        self._item_by_exe: dict[str, dict] = {}
        self._icon_cache: OrderedDict[str, tk.PhotoImage] = OrderedDict()
        self._auto_refresh_job: str | None = None
        self._refresh_token = 0
        self._bound_hotkeys: set[str] = set()
        self._auto_reconnect_jobs: dict[str, str] = {}
        self._target_apps: OrderedDict[str, dict] = OrderedDict()
        self._current_process_exe: str | None = None
        self._preset_autoload_pending = bool(self.config_data.get("preset_autosave", False))
        self.only_active_var = tk.BooleanVar(value=self.config_data.get("only_active_default", False))

        min_width, min_height = (430, 330) if self.ui_mode == "new" else (720, 420)
        geometry = clamp_geometry(
            self.config_data.get("window_geometry", CONFIG_DEFAULTS["window_geometry"]),
            CONFIG_DEFAULTS["window_geometry"],
            min_width,
            min_height,
        )
        self.geometry(geometry)
        self.minsize(min_width, min_height)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._build_ui()
        self._bind_hotkeys()
        self.refresh_app_list()

        if self.config_data.get("auto_refresh", False):
            self._start_auto_refresh()

    def _build_ui(self) -> None:
        if self.ui_mode == "new":
            self._build_new_ui()
        else:
            self._build_old_ui()

    def _build_old_ui(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            outer,
            text="Select an application below, then block or unblock its internet access.",
            font=(None, 12, "bold"),
        ).pack(anchor=tk.W, pady=(0, 8))

        search_frame = ttk.Frame(outer)
        search_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 6))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.search_entry.bind("<KeyRelease>", lambda _event: self._apply_filter())

        content = ttk.Frame(outer)
        content.pack(fill=tk.BOTH, expand=True)

        list_frame = ttk.Frame(content)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(list_frame, show="tree", selectmode="extended")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Button-3>", self.on_right_click)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(content, padding=(12, 0, 0, 0))
        button_frame.pack(side=tk.LEFT, fill=tk.Y)

        self.block_btn = ttk.Button(button_frame, text="Block", command=self.on_block)
        self.block_btn.pack(fill=tk.X, pady=(0, 6))

        self.unblock_btn = ttk.Button(button_frame, text="Unblock", command=self.on_unblock)
        self.unblock_btn.pack(fill=tk.X, pady=(0, 6))

        self.unblock_all_btn = ttk.Button(button_frame, text="Unblock all apps", command=self.on_unblock_all)
        self.unblock_all_btn.pack(fill=tk.X, pady=(0, 6))

        self.block_except_btn = ttk.Button(button_frame, text="Block all except", command=self.on_block_all_except)
        self.block_except_btn.pack(fill=tk.X, pady=(0, 6))

        self.stop_block_except_btn = ttk.Button(button_frame, text="Stop blocking all", command=self.on_stop_block_all)
        self.stop_block_except_btn.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(button_frame, text="Refresh", command=self.refresh_app_list).pack(fill=tk.X, pady=(0, 6))

        self.auto_refresh_btn = ttk.Button(button_frame, text="Auto refresh: Off", command=self._toggle_auto_refresh)
        self.auto_refresh_btn.pack(fill=tk.X, pady=(0, 6))
        self._update_auto_refresh_button()

        self.select_all_btn = ttk.Button(button_frame, text="Select all apps", command=self._toggle_select_all)
        self.select_all_btn.pack(fill=tk.X, pady=(0, 6))

        ttk.Button(button_frame, text="Settings", command=self._open_settings).pack(fill=tk.X, pady=(0, 6))

        ttk.Checkbutton(
            button_frame,
            text="Active connections only",
            variable=self.only_active_var,
            command=self.refresh_app_list,
        ).pack(fill=tk.X, pady=(0, 6))

        self.run_as_admin_btn = ttk.Button(button_frame, text="Run as administrator", command=self.on_run_as_admin)
        self.run_as_admin_btn.pack(fill=tk.X, pady=(6, 0))

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(outer, textvariable=self.status_var, anchor=tk.W).pack(fill=tk.X, pady=(10, 0))

        self.admin_var = tk.StringVar(value="Admin: Yes" if self.is_admin else "Admin: No")
        ttk.Label(outer, textvariable=self.admin_var, anchor=tk.W).pack(fill=tk.X, pady=(4, 0))

        if self.is_admin:
            self.run_as_admin_btn.state(["disabled"])
        else:
            for button in (
                self.block_btn,
                self.unblock_btn,
                self.unblock_all_btn,
                self.block_except_btn,
                self.stop_block_except_btn,
            ):
                button.state(["disabled"])

    def _build_new_ui(self) -> None:
        self.configure(bg="#1786e5")

        outer = tk.Frame(self, bg="#1786e5", padx=5, pady=5)
        outer.pack(fill=tk.BOTH, expand=True)

        title = tk.Label(
            outer,
            text="Internet Disconnector",
            bg="#1786e5",
            fg="white",
            font=("Segoe UI", 11, "bold"),
        )
        title.pack(anchor="w", pady=(0, 3))

        panel = tk.Frame(outer, bg="#1786e5")
        panel.pack(fill=tk.BOTH, expand=True)

        action_frame = tk.LabelFrame(
            panel,
            text="Block List",
            bg="#1786e5",
            fg="white",
            font=("Segoe UI", 9, "bold"),
            padx=5,
            pady=5,
        )
        action_frame.pack(fill=tk.BOTH, expand=True)

        left_actions = tk.Frame(action_frame, bg="#1786e5")
        left_actions.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))

        ttk.Button(left_actions, text="Add App", command=lambda: self._open_process_selector(add_to_list=True)).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(left_actions, text="Remove", command=self._remove_target_selection).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(left_actions, text="Clear", command=self._clear_target_selection).pack(fill=tk.X, pady=(0, 5))
        self.toggle_selected_btn = ttk.Button(left_actions, text="Block Selected", command=self.on_toggle_target_block_state)
        self.toggle_selected_btn.pack(fill=tk.X, pady=(0, 5))
        self.timer_toggle_btn = ttk.Button(left_actions, text="Start Timer", command=self.on_toggle_selected_timer)
        self.timer_toggle_btn.pack(fill=tk.X)

        list_frame = tk.Frame(action_frame, bg="#1786e5")
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.target_tree = ttk.Treeview(list_frame, columns=("status", "path"), show="headings", selectmode="extended")
        self.target_tree.heading("status", text="Status")
        self.target_tree.heading("path", text="App")
        self.target_tree.column("status", width=56, stretch=False, anchor=tk.CENTER)
        self.target_tree.column("path", width=190, stretch=True, anchor=tk.W)
        self.target_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.target_tree.bind("<<TreeviewSelect>>", lambda _event: self._update_new_ui_action_button())
        self.target_tree.bind("<Button-3>", self.on_target_right_click)

        target_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.target_tree.yview)
        target_scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.target_tree.configure(yscrollcommand=target_scrollbar.set)

        bottom_actions = tk.Frame(outer, bg="#1786e5")
        bottom_actions.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(bottom_actions, text="About", command=self._open_about).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(bottom_actions, text="Settings", command=self._open_settings).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        ttk.Button(bottom_actions, text="Refresh", command=self.refresh_app_list).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.inject_btn = ttk.Button(bottom_actions, text="Block All", command=self.on_toggle_all_target_apps)
        self.inject_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(outer, textvariable=self.status_var, bg="#1786e5", fg="white", anchor="w").pack(fill=tk.X, pady=(4, 0))

        if not self.is_admin:
            self.toggle_selected_btn.state(["disabled"])
            self.timer_toggle_btn.state(["disabled"])
            self.inject_btn.state(["disabled"])

    def _open_settings(self) -> None:
        window = tk.Toplevel(self)
        window.title("Settings")

        geometry = clamp_geometry(
            self.config_data.get("settings_window_geometry", CONFIG_DEFAULTS["settings_window_geometry"]),
            CONFIG_DEFAULTS["settings_window_geometry"],
            520,
            620,
        )
        window.geometry(geometry)
        window.minsize(520, 620)
        window.resizable(True, True)

        outer = ttk.Frame(window)
        outer.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        frame = ttk.Frame(canvas, padding=12)
        window_frame = canvas.create_window((0, 0), window=frame, anchor="nw")

        def update_scroll_region(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def update_frame_width(event) -> None:
            canvas.itemconfigure(window_frame, width=event.width)

        frame.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", update_frame_width)

        def on_mousewheel(event) -> None:
            if event.delta:
                canvas.yview_scroll(int(-event.delta / 120), "units")

        window.bind("<MouseWheel>", on_mousewheel)

        def close_settings() -> None:
            window.unbind("<MouseWheel>")
            window.destroy()

        window.protocol("WM_DELETE_WINDOW", close_settings)

        def create_hotkey_capture(parent, initial_value: str):
            value_var = tk.StringVar(value=initial_value)
            display_var = tk.StringVar(value=hotkey_sequence_to_display(initial_value))
            entry = ttk.Entry(parent, textvariable=display_var)

            def handle_keypress(event):
                if event.keysym in {"BackSpace", "Delete"} and not any(
                    event.state & MODIFIER_STATE_MASKS[key] for key in MODIFIER_STATE_MASKS
                ):
                    value_var.set("")
                    display_var.set("")
                    return "break"

                config_value = hotkey_event_to_config(event)
                if not config_value:
                    return "break"
                value_var.set(config_value)
                display_var.set(hotkey_sequence_to_display(config_value))
                return "break"

            entry.bind("<KeyPress>", handle_keypress)
            return value_var, entry

        hide_console_var = tk.BooleanVar(value=self.config_data.get("hide_console", False))
        auto_refresh_var = tk.BooleanVar(value=self.config_data.get("auto_refresh", False))
        preset_autosave_var = tk.BooleanVar(value=self.config_data.get("preset_autosave", False))
        only_active_default_var = tk.BooleanVar(value=self.config_data.get("only_active_default", False))
        show_block_status_var = tk.BooleanVar(value=self.config_data.get("show_block_status", True))
        show_info_notifications_var = tk.BooleanVar(value=self.config_data.get("show_info_notifications", True))
        ui_mode_var = tk.StringVar(value=self.config_data.get("ui_mode", "old"))
        ask_admin_var = tk.BooleanVar(value=self.config_data.get("ask_for_admin_on_startup", True))
        warn_not_admin_var = tk.BooleanVar(value=self.config_data.get("warn_when_not_admin", True))
        confirm_block_all_var = tk.BooleanVar(value=self.config_data.get("confirm_block_all_except", True))
        remember_geometry_var = tk.BooleanVar(value=self.config_data.get("remember_window_geometry", True))
        hotkeys_enabled_var = tk.BooleanVar(value=self.config_data.get("hotkeys_enabled", True))
        refresh_interval_var = tk.StringVar(value=str(self.config_data.get("refresh_interval", 5)))
        max_cached_icons_var = tk.StringVar(value=str(self.config_data.get("max_cached_icons", 192)))
        about_github_url_var = tk.StringVar(
            value=self.config_data.get(
                "about_github_url",
                "base64:aHR0cHM6Ly9naXRodWIuY29tL3NpcmFwYXRwaXBvL0Etc2VsZWN0YWJsZS1kaXNjb25uZWN0",
            )
        )
        window_geometry_var = tk.StringVar(value=self.config_data.get("window_geometry", CONFIG_DEFAULTS["window_geometry"]))
        settings_geometry_var = tk.StringVar(
            value=self.config_data.get("settings_window_geometry", CONFIG_DEFAULTS["settings_window_geometry"])
        )
        hotkey_focus_search_var, hotkey_focus_search_entry = create_hotkey_capture(
            frame, self.config_data.get("hotkey_focus_search", "CTRL+F")
        )
        hotkey_refresh_var, hotkey_refresh_entry = create_hotkey_capture(
            frame, self.config_data.get("hotkey_refresh", "VK_F5")
        )
        hotkey_block_var, hotkey_block_entry = create_hotkey_capture(
            frame, self.config_data.get("hotkey_block_selected", "CTRL+B")
        )
        hotkey_unblock_var, hotkey_unblock_entry = create_hotkey_capture(
            frame, self.config_data.get("hotkey_unblock_selected", "CTRL+U")
        )
        hotkey_unblock_all_var, hotkey_unblock_all_entry = create_hotkey_capture(
            frame, self.config_data.get("hotkey_unblock_all", "CTRL+SHIFT+U")
        )
        hotkey_toggle_selected_var, hotkey_toggle_selected_entry = create_hotkey_capture(
            frame, self.config_data.get("hotkey_toggle_selected", "CTRL+VK_RETURN")
        )

        ttk.Label(frame, text="UI mode:").pack(anchor=tk.W, pady=(0, 2))
        ttk.Combobox(frame, textvariable=ui_mode_var, values=("old", "new"), state="readonly").pack(fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(frame, text="Hide console window", variable=hide_console_var).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(frame, text="Auto refresh app list", variable=auto_refresh_var).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(frame, text="Auto save/load preset", variable=preset_autosave_var).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(
            frame,
            text="Show only active connections by default",
            variable=only_active_default_var,
        ).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(frame, text="Show block status in the list", variable=show_block_status_var).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(frame, text="Show info notifications", variable=show_info_notifications_var).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(frame, text="Ask for admin rights on startup", variable=ask_admin_var).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(frame, text="Warn when running without admin", variable=warn_not_admin_var).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(frame, text="Confirm block-all-except actions", variable=confirm_block_all_var).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(frame, text="Remember window size on close", variable=remember_geometry_var).pack(anchor=tk.W, pady=(0, 6))
        ttk.Checkbutton(frame, text="Enable app hotkeys", variable=hotkeys_enabled_var).pack(anchor=tk.W, pady=(0, 6))

        ttk.Label(frame, text="Refresh interval (seconds):").pack(anchor=tk.W, pady=(8, 2))
        ttk.Entry(frame, textvariable=refresh_interval_var).pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="Max cached icons:").pack(anchor=tk.W, pady=(6, 2))
        ttk.Entry(frame, textvariable=max_cached_icons_var).pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="Hotkey: focus search").pack(anchor=tk.W, pady=(6, 2))
        hotkey_focus_search_entry.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="Hotkey: refresh").pack(anchor=tk.W, pady=(6, 2))
        hotkey_refresh_entry.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="Hotkey: block selected").pack(anchor=tk.W, pady=(6, 2))
        hotkey_block_entry.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="Hotkey: unblock selected").pack(anchor=tk.W, pady=(6, 2))
        hotkey_unblock_entry.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="Hotkey: unblock all").pack(anchor=tk.W, pady=(6, 2))
        hotkey_unblock_all_entry.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="Hotkey: quick toggle selected").pack(anchor=tk.W, pady=(6, 2))
        hotkey_toggle_selected_entry.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="About GitHub link").pack(anchor=tk.W, pady=(6, 2))
        ttk.Entry(frame, textvariable=about_github_url_var).pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="Main window size (WxH):").pack(anchor=tk.W, pady=(6, 2))
        ttk.Entry(frame, textvariable=window_geometry_var).pack(fill=tk.X, pady=(0, 6))

        ttk.Label(frame, text="Settings window size (WxH):").pack(anchor=tk.W, pady=(6, 2))
        ttk.Entry(frame, textvariable=settings_geometry_var).pack(fill=tk.X, pady=(0, 10))

        button_row = ttk.Frame(frame)
        button_row.pack(fill=tk.X)

        def save_settings() -> None:
            previous_ui_mode = self.ui_mode
            try:
                interval = int(refresh_interval_var.get())
                if interval < 1:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid", "Refresh interval must be a positive integer.")
                return

            try:
                max_cached_icons = int(max_cached_icons_var.get())
                if max_cached_icons < 32:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid", "Max cached icons must be an integer of at least 32.")
                return

            window_geometry = window_geometry_var.get().strip()
            settings_geometry = settings_geometry_var.get().strip()
            if not is_valid_geometry(window_geometry):
                messagebox.showerror("Invalid", "Main window size must use the format WxH.")
                return
            if not is_valid_geometry(settings_geometry):
                messagebox.showerror("Invalid", "Settings window size must use the format WxH.")
                return

            requested_ui_mode = ui_mode_var.get().strip().lower() or "old"

            self.config_data.update(
                {
                    "ui_mode": requested_ui_mode,
                    "hide_console": hide_console_var.get(),
                    "auto_refresh": auto_refresh_var.get(),
                    "preset_autosave": preset_autosave_var.get(),
                    "refresh_interval": interval,
                    "only_active_default": only_active_default_var.get(),
                    "show_block_status": show_block_status_var.get(),
                    "show_info_notifications": show_info_notifications_var.get(),
                    "ask_for_admin_on_startup": ask_admin_var.get(),
                    "warn_when_not_admin": warn_not_admin_var.get(),
                    "confirm_block_all_except": confirm_block_all_var.get(),
                    "remember_window_geometry": remember_geometry_var.get(),
                    "max_cached_icons": max_cached_icons,
                    "hotkeys_enabled": hotkeys_enabled_var.get(),
                    "hotkey_focus_search": hotkey_focus_search_var.get().strip(),
                    "hotkey_refresh": hotkey_refresh_var.get().strip(),
                    "hotkey_block_selected": hotkey_block_var.get().strip(),
                    "hotkey_unblock_selected": hotkey_unblock_var.get().strip(),
                    "hotkey_unblock_all": hotkey_unblock_all_var.get().strip(),
                    "hotkey_toggle_selected": hotkey_toggle_selected_var.get().strip(),
                    "about_github_url": about_github_url_var.get().strip(),
                    "window_geometry": window_geometry,
                    "settings_window_geometry": settings_geometry,
                }
            )

            save_config(self.config_data)
            self._trim_icon_cache()
            self._bind_hotkeys()
            if hasattr(self, "hotkey_hint_var"):
                self.hotkey_hint_var.set(self._build_hotkey_hint_text())
            self.only_active_var.set(self.config_data["only_active_default"])
            self.geometry(window_geometry)
            self._update_auto_refresh_button()

            if self.config_data.get("auto_refresh", False):
                self._start_auto_refresh()
            else:
                self._stop_auto_refresh()

            if self.config_data.get("preset_autosave", False):
                self._save_autosave_preset()

            self.refresh_app_list()
            close_settings()
            if previous_ui_mode != requested_ui_mode:
                self._show_info_message("UI mode changed", "Restart the app to apply the new UI mode.")

        preset_row = ttk.Frame(frame)
        preset_row.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(preset_row, text="Save preset", command=self._save_preset_dialog).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(preset_row, text="Load preset", command=self._load_preset_dialog).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(preset_row, text="Open presets folder", command=self._open_presets_folder).pack(side=tk.LEFT)

        ttk.Button(button_row, text="Save", command=save_settings).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(button_row, text="Cancel", command=close_settings).pack(side=tk.LEFT)
        ttk.Button(button_row, text="Open config file", command=self._open_config_file).pack(side=tk.RIGHT)

    def _open_config_file(self) -> None:
        try:
            subprocess.Popen(["notepad", str(CONFIG_PATH)])
        except Exception as error:
            messagebox.showerror("Error", f"Failed to open config: {error}")

    def _autosave_preset_path(self) -> Path:
        return get_preset_dir() / "_autosave.json"

    def _current_preset_items(self) -> list[dict]:
        if self.ui_mode == "new":
            return [
                {
                    "exe": item["exe"],
                    "name": item.get("name") or os.path.basename(item["exe"]),
                }
                for item in self._target_apps.values()
            ]

        return [
            {
                "exe": item["exe"],
                "name": item.get("name") or os.path.basename(item["exe"]),
            }
            for item in self.get_selected_items()
        ]

    def _read_preset_file(self, path: Path) -> list[dict]:
        raw = json.loads(path.read_text(encoding="utf-8"))
        items = raw.get("apps", raw) if isinstance(raw, dict) else raw
        if not isinstance(items, list):
            raise ValueError("Preset file must contain a list of apps.")

        parsed: list[dict] = []
        for item in items:
            if isinstance(item, str):
                exe_path = item.strip()
                if exe_path:
                    parsed.append({"exe": exe_path, "name": os.path.basename(exe_path)})
                continue
            if isinstance(item, dict):
                exe_path = str(item.get("exe", "")).strip()
                if exe_path:
                    parsed.append(
                        {
                            "exe": exe_path,
                            "name": str(item.get("name") or os.path.basename(exe_path)),
                        }
                    )
        return parsed

    def _write_preset_file(self, path: Path, items: list[dict]) -> None:
        payload = {
            "apps": items,
            "ui_mode": self.ui_mode,
        }
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _autosave_preset_if_enabled(self) -> None:
        if self.config_data.get("preset_autosave", False):
            self._save_autosave_preset()

    def _apply_preset_items(self, items: list[dict]) -> None:
        unique: OrderedDict[str, dict] = OrderedDict()
        for item in items:
            exe_path = item.get("exe")
            if not exe_path:
                continue
            live_item = self._item_by_exe.get(exe_path)
            unique[exe_path] = {
                "exe": exe_path,
                "name": (live_item or item).get("name") or os.path.basename(exe_path),
                "pid": (live_item or {}).get("pid"),
                "blocked": (live_item or {}).get("blocked", False),
            }

        if self.ui_mode == "new":
            self._target_apps = unique
            self._refresh_target_tree()
            if unique and hasattr(self, "target_tree"):
                self.target_tree.selection_set(next(iter(unique)))
            self._autosave_preset_if_enabled()
            self.set_status(f"Loaded preset with {len(unique)} app(s).")
            return

        matches = [exe for exe in unique if exe in self._item_by_exe]
        if hasattr(self, "tree"):
            self.tree.selection_remove(self.tree.selection())
            if matches:
                self.tree.selection_set(matches)
        self.set_status(f"Loaded preset with {len(matches)} running app(s).")

    def _save_preset_dialog(self) -> None:
        items = self._current_preset_items()
        if not items:
            self._show_info_message("Save preset", "There is nothing to save right now.")
            return

        path_text = filedialog.asksaveasfilename(
            parent=self,
            title="Save preset",
            initialdir=str(get_preset_dir()),
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path_text:
            return

        try:
            path = Path(path_text)
            self._write_preset_file(path, items)
            self.set_status(f"Preset saved: {path.name}")
        except Exception as error:
            messagebox.showerror("Save preset failed", str(error))

    def _load_preset_dialog(self) -> None:
        path_text = filedialog.askopenfilename(
            parent=self,
            title="Load preset",
            initialdir=str(get_preset_dir()),
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path_text:
            return

        try:
            items = self._read_preset_file(Path(path_text))
            self._apply_preset_items(items)
        except Exception as error:
            messagebox.showerror("Load preset failed", str(error))

    def _open_presets_folder(self) -> None:
        preset_dir = get_preset_dir()
        try:
            subprocess.Popen(["explorer", str(preset_dir)])
        except Exception as error:
            messagebox.showerror("Error", f"Failed to open presets folder: {error}")

    def _save_autosave_preset(self) -> None:
        if not self.config_data.get("preset_autosave", False):
            return
        items = self._current_preset_items()
        try:
            self._write_preset_file(self._autosave_preset_path(), items)
        except Exception:
            pass

    def _load_autosave_preset(self) -> None:
        autosave_path = self._autosave_preset_path()
        if not autosave_path.exists():
            return
        try:
            items = self._read_preset_file(autosave_path)
            if items:
                self._apply_preset_items(items)
        except Exception:
            pass

    def _build_hotkey_hint_text(self) -> str:
        if not self.config_data.get("hotkeys_enabled", True):
            return "Hotkeys: disabled"

        focus_label = "Select" if self.ui_mode == "new" else "Search"
        parts = [
            f"{focus_label} {hotkey_sequence_to_display(self.config_data.get('hotkey_focus_search', 'CTRL+F'))}",
            f"Refresh {hotkey_sequence_to_display(self.config_data.get('hotkey_refresh', 'VK_F5'))}",
            f"Toggle {hotkey_sequence_to_display(self.config_data.get('hotkey_toggle_selected', 'CTRL+VK_RETURN'))}",
            f"Unblock all {hotkey_sequence_to_display(self.config_data.get('hotkey_unblock_all', 'CTRL+SHIFT+U'))}",
        ]
        return " | ".join(parts)

    def _set_current_process(self, item: dict | None) -> None:
        self._current_process_exe = item["exe"] if item else None

    def _open_process_selector(self, add_to_list: bool = False) -> None:
        if not self._all_items:
            self.refresh_app_list()
            self._show_info_message("Process list", "Refreshing the app list first. Try again in a moment.")
            return

        dialog = tk.Toplevel(self)
        dialog.title("Select App")
        dialog.geometry("620x420")
        dialog.minsize(520, 320)

        wrapper = ttk.Frame(dialog, padding=12)
        wrapper.pack(fill=tk.BOTH, expand=True)

        ttk.Label(wrapper, text="Search running apps:").pack(anchor=tk.W, pady=(0, 4))
        filter_var = tk.StringVar()
        filter_entry = ttk.Entry(wrapper, textvariable=filter_var)
        filter_entry.pack(fill=tk.X, pady=(0, 8))

        active_only_var = tk.BooleanVar(value=self.only_active_var.get())
        ttk.Checkbutton(
            wrapper,
            text="Only show active connections",
            variable=active_only_var,
        ).pack(anchor=tk.W, pady=(0, 8))

        list_frame = ttk.Frame(wrapper)
        list_frame.pack(fill=tk.BOTH, expand=True)

        tree = ttk.Treeview(list_frame, columns=("path",), show="tree", selectmode="browse")
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        tree.configure(yscrollcommand=scrollbar.set)

        items_by_iid: dict[str, dict] = {}
        selector_items: list[dict] = list(self._all_items)

        def load_items() -> None:
            nonlocal selector_items
            if active_only_var.get():
                try:
                    selector_items = find_running_apps(True)
                except Exception:
                    selector_items = list(self._all_items)
            else:
                selector_items = list(self._all_items)
            populate(filter_var.get())

        def populate(term: str = "") -> None:
            tree.delete(*tree.get_children())
            items_by_iid.clear()
            lowered = term.strip().lower()
            for item in selector_items:
                if lowered and lowered not in item["name"].lower() and lowered not in item["exe"].lower():
                    continue
                items_by_iid[item["exe"]] = item
                tree.insert("", "end", iid=item["exe"], text=self._format_item_text(item), image=self._get_icon(item["exe"]))

        def choose_selected() -> None:
            selection = tree.selection()
            if not selection:
                return
            item = items_by_iid.get(selection[0])
            self._set_current_process(item)
            if add_to_list:
                self._add_current_process_to_target_list()
            dialog.destroy()

        filter_entry.bind("<KeyRelease>", lambda _event: populate(filter_var.get()))
        active_only_var.trace_add("write", lambda *_args: load_items())
        tree.bind("<Double-1>", lambda _event: choose_selected())

        def refresh_selector() -> None:
            try:
                selector_items[:] = find_running_apps(active_only_var.get())
                populate(filter_var.get())
                self.set_status(f"Refreshed app selector: {len(selector_items)} app(s).")
            except Exception as error:
                messagebox.showerror("Refresh failed", f"Could not refresh the app selector.\n\n{error}")

        button_row = ttk.Frame(wrapper)
        button_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(button_row, text="Select", command=choose_selected).pack(side=tk.LEFT)
        ttk.Button(button_row, text="Refresh", command=refresh_selector).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(button_row, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT)

        load_items()
        filter_entry.focus_set()

    def _refresh_target_tree(self) -> None:
        if not hasattr(self, "target_tree"):
            return

        previous_selection = set(self.target_tree.selection())
        self.target_tree.delete(*self.target_tree.get_children())
        blocked_lookup = {item["exe"]: bool(item.get("blocked")) for item in self._all_items}

        for exe_path, item in self._target_apps.items():
            item["blocked"] = blocked_lookup.get(exe_path, item.get("blocked", False))
            status = STATUS_BLOCKED if item.get("blocked") else "Ready"
            live_item = self._item_by_exe.get(exe_path)
            if live_item and live_item.get("pid") is not None:
                item["pid"] = live_item.get("pid")
            name = self._format_advanced_app_label(item.get("name") or os.path.basename(exe_path))
            self.target_tree.insert("", "end", iid=exe_path, values=(status, name))

        restored = [exe for exe in previous_selection if exe in self._target_apps]
        if restored:
            self.target_tree.selection_set(restored)

        self._update_new_ui_action_button()

    def _add_current_process_to_target_list(self) -> None:
        if not self._current_process_exe:
            self._show_info_message("Select an app", "Choose a running app first.")
            return

        source_item = self._item_by_exe.get(self._current_process_exe)
        if source_item is None:
            source_item = {
                "name": os.path.basename(self._current_process_exe),
                "exe": self._current_process_exe,
                "blocked": False,
            }

        self._target_apps[self._current_process_exe] = {
            "name": source_item["name"],
            "exe": source_item["exe"],
            "pid": source_item.get("pid"),
            "blocked": source_item.get("blocked", False),
        }
        self._refresh_target_tree()
        self.target_tree.selection_set(self._current_process_exe)
        self._autosave_preset_if_enabled()
        self.set_status(f"Added {source_item['name']} to the block list.")

    def _remove_target_selection(self) -> None:
        if not hasattr(self, "target_tree"):
            return
        removed = list(self.target_tree.selection())
        for exe_path in removed:
            self._target_apps.pop(exe_path, None)
        self._refresh_target_tree()
        self._autosave_preset_if_enabled()
        if removed:
            self.set_status(f"Removed {len(removed)} app(s) from the block list.")

    def _clear_target_selection(self) -> None:
        self._target_apps.clear()
        if hasattr(self, "target_tree"):
            self.target_tree.delete(*self.target_tree.get_children())
        self._update_new_ui_action_button()
        self._autosave_preset_if_enabled()
        self.set_status("Block list cleared.")

    def _update_new_ui_action_button(self) -> None:
        if not hasattr(self, "inject_btn"):
            return
        queue_exes = list(self._target_apps.keys())
        selected = self.get_selected_exes()
        if hasattr(self, "toggle_selected_btn"):
            if not selected:
                self.toggle_selected_btn.config(text="Block Selected")
                self.toggle_selected_btn.state(["disabled"])
                if hasattr(self, "timer_toggle_btn"):
                    self.timer_toggle_btn.config(text="Start Timer")
                    self.timer_toggle_btn.state(["disabled"])
            else:
                if self.is_admin:
                    self.toggle_selected_btn.state(["!disabled"])
                else:
                    self.toggle_selected_btn.state(["disabled"])
                selected_blocked = [exe for exe in selected if self._target_apps.get(exe, {}).get("blocked")]
                selected_clear = [exe for exe in selected if exe not in selected_blocked]
                if selected_blocked and not selected_clear:
                    self.toggle_selected_btn.config(text="Unblock Selected")
                elif selected_clear and not selected_blocked:
                    self.toggle_selected_btn.config(text="Block Selected")
                else:
                    self.toggle_selected_btn.config(text="Block / Unblock Selected")

                if hasattr(self, "timer_toggle_btn"):
                    selected_running_timers = [exe for exe in selected if exe in self._auto_reconnect_jobs]
                    if selected_running_timers and len(selected_running_timers) == len(selected):
                        self.timer_toggle_btn.config(text="Stop Timer")
                        self.timer_toggle_btn.state(["!disabled"])
                    else:
                        self.timer_toggle_btn.config(text="Start Timer")
                        if self.is_admin:
                            self.timer_toggle_btn.state(["!disabled"])
                        else:
                            self.timer_toggle_btn.state(["disabled"])

        if not queue_exes:
            self.inject_btn.config(text="Block All")
            self.inject_btn.state(["disabled"])
            return

        if self.is_admin:
            self.inject_btn.state(["!disabled"])
        else:
            self.inject_btn.state(["disabled"])
        if all(self._target_apps.get(exe, {}).get("blocked") for exe in queue_exes):
            self.inject_btn.config(text="Unblock All")
        else:
            self.inject_btn.config(text="Block All")

    def _open_about(self) -> None:
        raw_github_url = (self.config_data.get("about_github_url", "") or "").strip()
        github_url = decode_config_url(raw_github_url)
        if github_url:
            try:
                webbrowser.open(github_url)
                self._show_info_message("About", f"Opening:\n{github_url}")
            except Exception as error:
                messagebox.showerror("About", f"Could not open the GitHub link.\n\n{error}")
            return

        self._show_info_message(
            "About",
            "Add your GitHub link in Settings or in internet_disconnector.config.json.\n\n"
            "Field: about_github_url\n"
            "You can store a plain URL or use base64:...",
        )

    def refresh_app_list(self) -> None:
        self._refresh_token += 1
        refresh_token = self._refresh_token
        self.set_status("Refreshing list of running apps...")
        if self.ui_mode == "new" and hasattr(self, "target_tree"):
            self.target_tree.delete(*self.target_tree.get_children())
        elif hasattr(self, "tree"):
            self.tree.delete(*self.tree.get_children())

        only_active = self.only_active_var.get()
        show_status = self.config_data.get("show_block_status", True)
        result_queue: queue.Queue = queue.Queue()

        def worker() -> None:
            try:
                items = find_running_apps(only_active)
                result_queue.put((load_app_statuses(items, show_status), None))
            except Exception as error:
                result_queue.put((None, error))

        def check_queue() -> None:
            try:
                items, error = result_queue.get_nowait()
            except queue.Empty:
                if refresh_token == self._refresh_token:
                    self.after(100, check_queue)
                return

            if refresh_token != self._refresh_token:
                return

            if error is not None:
                self._refresh_failed(error)
            else:
                self._refresh_done(items or [])

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, check_queue)

    def _refresh_failed(self, error: Exception) -> None:
        messagebox.showerror("Error", f"Failed to list processes: {error}")
        self.set_status("Failed to refresh the app list.")

    def _refresh_done(self, items: list[dict]) -> None:
        self._all_items = items
        if self.ui_mode == "new":
            self._item_by_exe = {item["exe"]: item for item in items}
            self._refresh_target_tree()
            self.set_status(f"Loaded {len(items)} running apps.")
        else:
            self._apply_filter()
            self.select_all_btn.config(text="Select all apps")

        if self._preset_autoload_pending:
            self._preset_autoload_pending = False
            self._load_autosave_preset()

    def _start_auto_refresh(self) -> None:
        self._stop_auto_refresh()
        interval = max(1, int(self.config_data.get("refresh_interval", 5)))
        self._auto_refresh_job = self.after(interval * 1000, self._auto_refresh_tick)

    def _stop_auto_refresh(self) -> None:
        if self._auto_refresh_job is not None:
            self.after_cancel(self._auto_refresh_job)
            self._auto_refresh_job = None

    def _auto_refresh_tick(self) -> None:
        self.refresh_app_list()
        if self.config_data.get("auto_refresh", False):
            self._start_auto_refresh()

    def _update_auto_refresh_button(self) -> None:
        state = "On" if self.config_data.get("auto_refresh", False) else "Off"
        self.auto_refresh_btn.config(text=f"Auto refresh: {state}")

    def _toggle_auto_refresh(self) -> None:
        self.config_data["auto_refresh"] = not self.config_data.get("auto_refresh", False)
        save_config(self.config_data)
        self._update_auto_refresh_button()

        if self.config_data["auto_refresh"]:
            self._start_auto_refresh()
            self.set_status("Auto refresh enabled.")
        else:
            self._stop_auto_refresh()
            self.set_status("Auto refresh disabled.")

    def _toggle_select_all(self) -> None:
        if self.ui_mode == "new":
            return
        all_items = self.tree.get_children()
        selected = self.tree.selection()
        if len(selected) == len(all_items):
            self.tree.selection_remove(all_items)
            self.select_all_btn.config(text="Select all apps")
            self.set_status("Selection cleared.")
            return

        self.tree.selection_set(all_items)
        self.select_all_btn.config(text="Clear selection")
        self.set_status(f"Selected {len(all_items)} apps.")

    def _apply_filter(self) -> None:
        if self.ui_mode == "new":
            self._refresh_target_tree()
            return
        term = self.search_var.get().strip().lower()
        if term:
            items = [
                item
                for item in self._all_items
                if term in item["name"].lower()
                or term in item["exe"].lower()
                or term in self._display_status_label(item).lower()
            ]
        else:
            items = list(self._all_items)

        self._populate_tree(items)
        self.set_status(f"Loaded {len(items)} apps (of {len(self._all_items)}).")

    def _display_status_label(self, item: dict) -> str:
        blocked = item.get("blocked")
        if blocked is True:
            return STATUS_BLOCKED
        return ""

    def _format_advanced_app_label(self, label: str) -> str:
        return f"{label} ..."

    def _format_item_text(self, item: dict) -> str:
        label = self._format_advanced_app_label(item["name"])
        if not self.config_data.get("show_block_status", True):
            return label

        status_label = self._display_status_label(item)
        if not status_label:
            return label
        return f"{label} [{status_label}]"

    def _populate_tree(self, items: list[dict]) -> None:
        self.tree.delete(*self.tree.get_children())
        self._item_by_exe = {}

        for item in items:
            exe_path = item["exe"]
            self._item_by_exe[exe_path] = item
            self.tree.insert(
                "",
                "end",
                iid=exe_path,
                text=self._format_item_text(item),
                image=self._get_icon(exe_path),
            )

    def _get_icon(self, exe_path: str) -> tk.PhotoImage:
        if exe_path in self._icon_cache:
            self._icon_cache.move_to_end(exe_path)
            return self._icon_cache[exe_path]

        image = self._get_icon_from_exe(exe_path)
        if image is None:
            fallback = tk.PhotoImage(width=16, height=16)
            fallback.put("#" + format(abs(hash(exe_path)) & 0xFFFFFF, "06x"), to=(0, 0, 16, 16))
            image = fallback

        self._icon_cache[exe_path] = image
        self._trim_icon_cache()
        return image

    def _trim_icon_cache(self) -> None:
        max_cached_icons = max(32, int(self.config_data.get("max_cached_icons", 192) or 192))
        while len(self._icon_cache) > max_cached_icons:
            self._icon_cache.popitem(last=False)

    def _get_icon_from_exe(self, exe_path: str, size: int = 16) -> tk.PhotoImage | None:
        try:
            bgra = self._extract_icon_bgra(exe_path, size)
            if not bgra:
                return None

            image = tk.PhotoImage(width=size, height=size)
            for y in range(size):
                row = []
                for x in range(size):
                    index = (y * size + x) * 4
                    blue, green, red, _alpha = bgra[index : index + 4]
                    row.append(f"#{red:02x}{green:02x}{blue:02x}")
                image.put("{" + " ".join(row) + "}", to=(0, y))
            return image
        except Exception:
            return None

    def _extract_icon_bgra(self, exe_path: str, size: int = 16) -> bytes | None:
        try:
            import ctypes
            from ctypes import wintypes

            SHGFI_ICON = 0x000000100
            SHGFI_SMALLICON = 0x000000001
            SHGFI_USEFILEATTRIBUTES = 0x000000010

            class SHFILEINFO(ctypes.Structure):
                _fields_ = [
                    ("hIcon", wintypes.HICON),
                    ("iIcon", wintypes.INT),
                    ("dwAttributes", wintypes.DWORD),
                    ("szDisplayName", wintypes.WCHAR * 260),
                    ("szTypeName", wintypes.WCHAR * 80),
                ]

            class BITMAPINFOHEADER(ctypes.Structure):
                _fields_ = [
                    ("biSize", wintypes.DWORD),
                    ("biWidth", wintypes.LONG),
                    ("biHeight", wintypes.LONG),
                    ("biPlanes", wintypes.WORD),
                    ("biBitCount", wintypes.WORD),
                    ("biCompression", wintypes.DWORD),
                    ("biSizeImage", wintypes.DWORD),
                    ("biXPelsPerMeter", wintypes.LONG),
                    ("biYPelsPerMeter", wintypes.LONG),
                    ("biClrUsed", wintypes.DWORD),
                    ("biClrImportant", wintypes.DWORD),
                ]

            file_info = SHFILEINFO()
            result = ctypes.windll.shell32.SHGetFileInfoW(
                str(exe_path),
                0,
                ctypes.byref(file_info),
                ctypes.sizeof(file_info),
                SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES,
            )
            if result == 0 or not file_info.hIcon:
                return None

            bitmap_info = BITMAPINFOHEADER()
            bitmap_info.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bitmap_info.biWidth = size
            bitmap_info.biHeight = size
            bitmap_info.biPlanes = 1
            bitmap_info.biBitCount = 32
            bitmap_info.biCompression = 0

            device_context = ctypes.windll.user32.GetDC(None)
            memory_context = ctypes.windll.gdi32.CreateCompatibleDC(device_context)
            bits = ctypes.c_void_p()
            bitmap = ctypes.windll.gdi32.CreateDIBSection(
                memory_context,
                ctypes.byref(bitmap_info),
                0,
                ctypes.byref(bits),
                None,
                0,
            )
            old_bitmap = ctypes.windll.gdi32.SelectObject(memory_context, bitmap)

            ctypes.windll.user32.DrawIconEx(memory_context, 0, 0, file_info.hIcon, size, size, 0, None, 3)

            buffer = (ctypes.c_ubyte * (size * size * 4)).from_address(bits.value)
            data = bytes(buffer)

            ctypes.windll.gdi32.SelectObject(memory_context, old_bitmap)
            ctypes.windll.gdi32.DeleteObject(bitmap)
            ctypes.windll.gdi32.DeleteDC(memory_context)
            ctypes.windll.user32.ReleaseDC(None, device_context)
            ctypes.windll.user32.DestroyIcon(file_info.hIcon)
            return data
        except Exception:
            return None

    def get_selected_exes(self) -> list[str]:
        if self.ui_mode == "new" and hasattr(self, "target_tree"):
            return list(self.target_tree.selection())
        return list(self.tree.selection())

    def get_selected_items(self) -> list[dict]:
        if self.ui_mode == "new":
            return [self._target_apps[exe] for exe in self.get_selected_exes() if exe in self._target_apps]
        return [self._item_by_exe[exe] for exe in self.get_selected_exes() if exe in self._item_by_exe]

    def _is_app_blocked(self, exe_path: str) -> bool:
        target_item = self._target_apps.get(exe_path)
        if target_item is not None and target_item.get("blocked") is not None:
            return bool(target_item.get("blocked"))

        source_item = self._item_by_exe.get(exe_path)
        if source_item is not None and source_item.get("blocked") is not None:
            return bool(source_item.get("blocked"))

        try:
            return is_blocked(exe_path)
        except Exception:
            return False

    def _saved_reconnect_timer_map(self) -> dict[str, int]:
        raw = self.config_data.get("saved_reconnect_timers", {})
        if not isinstance(raw, dict):
            raw = {}
            self.config_data["saved_reconnect_timers"] = raw
        return raw

    def _disabled_reconnect_timer_list(self) -> list[str]:
        raw = self.config_data.get("disabled_reconnect_timers", [])
        if not isinstance(raw, list):
            raw = []
            self.config_data["disabled_reconnect_timers"] = raw
        return raw

    def _saved_reconnect_timer_for(self, exe_path: str) -> int | None:
        raw = self._saved_reconnect_timer_map().get(_normalize_path(exe_path))
        try:
            return max(1, int(raw)) if raw is not None else None
        except (TypeError, ValueError):
            return None

    def _saved_reconnect_timer_enabled(self, exe_path: str) -> bool:
        normalized = _normalize_path(exe_path)
        if self._saved_reconnect_timer_for(exe_path) is None:
            return False
        return normalized not in {str(item).lower() for item in self._disabled_reconnect_timer_list()}

    def _set_saved_reconnect_timer(self, exe_path: str, seconds: int) -> None:
        self.config_data["auto_reconnect_seconds"] = max(1, int(seconds))
        normalized = _normalize_path(exe_path)
        self._saved_reconnect_timer_map()[normalized] = max(1, int(seconds))
        disabled = self._disabled_reconnect_timer_list()
        if normalized in disabled:
            disabled.remove(normalized)
        save_config(self.config_data)

    def _set_saved_reconnect_timer_enabled(self, exe_path: str, enabled: bool) -> None:
        normalized = _normalize_path(exe_path)
        disabled = self._disabled_reconnect_timer_list()
        if enabled:
            while normalized in disabled:
                disabled.remove(normalized)
        elif normalized not in disabled:
            disabled.append(normalized)
        save_config(self.config_data)

    def _clear_saved_reconnect_timer(self, exe_path: str) -> None:
        normalized = _normalize_path(exe_path)
        timers = self._saved_reconnect_timer_map()
        timers.pop(normalized, None)
        disabled = self._disabled_reconnect_timer_list()
        while normalized in disabled:
            disabled.remove(normalized)
        save_config(self.config_data)

    def _clear_saved_reconnect_timer_for_app(self, exe_path: str) -> None:
        self._clear_saved_reconnect_timer(exe_path)
        self.set_status(f"Cleared saved reconnect timer for {os.path.basename(exe_path)}.")

    def _toggle_saved_reconnect_timer_for_app(self, exe_path: str) -> None:
        enabled = not self._saved_reconnect_timer_enabled(exe_path)
        self._set_saved_reconnect_timer_enabled(exe_path, enabled)
        if not enabled:
            self._cancel_auto_reconnect(exe_path)
        state = "enabled" if enabled else "disabled"
        self.set_status(f"Saved reconnect timer {state} for {os.path.basename(exe_path)}.")
        self._update_new_ui_action_button()

    def _cancel_auto_reconnect_for_app(self, exe_path: str) -> None:
        self._cancel_auto_reconnect(exe_path)
        self.set_status(f"Cancelled auto reconnect for {os.path.basename(exe_path)}.")
        self._update_new_ui_action_button()

    def _prompt_app_reconnect_timer(self, exe_path: str) -> bool:
        app_name = os.path.basename(exe_path)
        default_seconds = self._saved_reconnect_timer_for(exe_path) or self.config_data.get("auto_reconnect_seconds", 60)
        seconds = simpledialog.askinteger(
            "Reconnect Timer",
            f"Save the reconnect timer for {app_name}.\n\n"
            "Enter seconds.\n"
            "Use 0 to clear the saved timer.",
            parent=self,
            initialvalue=default_seconds,
            minvalue=0,
        )
        if seconds is None:
            return False

        if seconds == 0:
            self._clear_saved_reconnect_timer(exe_path)
            self.set_status(f"Cleared saved reconnect timer for {app_name}.")
            return True

        self._set_saved_reconnect_timer(exe_path, seconds)
        self._show_info_message(
            "Reconnect timer saved",
            f"{app_name} will use {format_duration_label(seconds)} when you block with timer.",
        )
        self.set_status(f"Saved reconnect timer for {app_name}: {format_duration_label(seconds)}.")
        return True

    def _block_with_saved_timer(self, exes: list[str], *, empty_title: str, empty_message: str, prompt_single: bool = False) -> None:
        timer_overrides: dict[str, int] = {}
        missing_timers: list[str] = []
        disabled_timers: list[str] = []

        for exe_path in exes:
            saved_seconds = self._saved_reconnect_timer_for(exe_path)
            if saved_seconds is None and prompt_single and len(exes) == 1:
                if self._prompt_app_reconnect_timer(exe_path):
                    saved_seconds = self._saved_reconnect_timer_for(exe_path)
            if saved_seconds is None:
                missing_timers.append(exe_path)
                continue
            if not self._saved_reconnect_timer_enabled(exe_path):
                disabled_timers.append(exe_path)
                continue
            timer_overrides[exe_path] = saved_seconds

        if not timer_overrides:
            heading = "Could not start reconnect timer:"
            extra_sections = [("Saved timer disabled:", disabled_timers)] if disabled_timers else None
            self._show_action_summary("Reconnect timer", heading, missing_timers, extra_sections=extra_sections)
            self.set_status("Reconnect timer is missing or disabled for the selected app(s).")
            return

        self._block_exes(
            list(timer_overrides.keys()),
            empty_title=empty_title,
            empty_message=empty_message,
            timer_overrides=timer_overrides,
            skipped_timer_exes=missing_timers,
            disabled_timer_exes=disabled_timers,
        )

    def _start_or_stop_timers(self, exes: list[str], *, empty_title: str, empty_message: str) -> None:
        if not exes:
            self._show_info_message(empty_title, empty_message)
            return

        running = [exe for exe in exes if exe in self._auto_reconnect_jobs]
        pending = [exe for exe in exes if exe not in self._auto_reconnect_jobs]

        if running and not pending:
            for exe_path in running:
                self._cancel_auto_reconnect(exe_path)
            self._show_action_summary("Reconnect timer stopped", f"Stopped reconnect timer for {len(running)} app(s):", running)
            self.set_status(f"Stopped reconnect timer for {len(running)} app(s).")
            self._update_new_ui_action_button()
            return

        self._block_with_saved_timer(pending, empty_title=empty_title, empty_message=empty_message)

    def _show_app_context_menu(self, event, exe_path: str, pid: int | None = None) -> None:
        if not exe_path:
            return

        blocked = self._is_app_blocked(exe_path)
        timer_pending = exe_path in self._auto_reconnect_jobs
        saved_timer = self._saved_reconnect_timer_for(exe_path)
        timer_enabled = self._saved_reconnect_timer_enabled(exe_path)

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(
            label="Block",
            command=lambda: self._block_exes([exe_path], empty_title="Select an app", empty_message="Please select at least one application first."),
            state=tk.DISABLED if blocked else tk.NORMAL,
        )
        menu.add_command(
            label="Unblock",
            command=lambda: self._unblock_exes([exe_path], empty_title="Select an app", empty_message="Please select at least one application first."),
            state=tk.NORMAL if blocked else tk.DISABLED,
        )
        menu.add_separator()
        menu.add_command(
            label="Block + reconnect timer",
            command=lambda: self._block_with_saved_timer(
                [exe_path],
                empty_title="Select an app",
                empty_message="Please select at least one application first.",
                prompt_single=True,
            ),
            state=tk.NORMAL if saved_timer is None or timer_enabled else tk.DISABLED,
        )
        menu.add_command(label="Set reconnect timer...", command=lambda: self._prompt_app_reconnect_timer(exe_path))
        if saved_timer is not None:
            menu.add_command(
                label="Disable saved timer" if timer_enabled else "Enable saved timer",
                command=lambda: self._toggle_saved_reconnect_timer_for_app(exe_path),
            )
            menu.add_command(label="Clear saved timer", command=lambda: self._clear_saved_reconnect_timer_for_app(exe_path))
        if timer_pending:
            menu.add_command(
                label="Stop reconnect timer",
                command=lambda: self._cancel_auto_reconnect_for_app(exe_path),
            )
        menu.add_separator()
        menu.add_command(label="Copy path", command=lambda: self._copy_path(exe_path))
        menu.add_command(label="Open folder", command=lambda: self._open_folder(exe_path))
        if pid is not None:
            menu.add_command(label="End task", command=lambda: self._end_task(pid))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def on_right_click(self, event) -> None:
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return

        self.tree.selection_set(item_id)
        item = self._item_by_exe.get(item_id)
        self._show_app_context_menu(event, item_id, item.get("pid") if item else None)

    def on_target_right_click(self, event) -> None:
        if not hasattr(self, "target_tree"):
            return

        item_id = self.target_tree.identify_row(event.y)
        if not item_id:
            return

        self.target_tree.selection_set(item_id)
        self._update_new_ui_action_button()
        item = self._target_apps.get(item_id) or self._item_by_exe.get(item_id) or {}
        self._show_app_context_menu(event, item_id, item.get("pid"))

    def _copy_path(self, exe_path: str | None) -> None:
        if not exe_path:
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(exe_path)
            self.set_status("Copied path to the clipboard.")
        except Exception:
            self.set_status("Could not copy the path.")

    def _open_folder(self, exe_path: str | None) -> None:
        if not exe_path:
            return
        try:
            subprocess.Popen(["explorer", "/select,", exe_path])
        except Exception as error:
            messagebox.showerror("Error", f"Failed to open folder: {error}")

    def _end_task(self, pid: int | None) -> None:
        if pid is None or psutil is None:
            return
        try:
            process = psutil.Process(pid)
            process.terminate()
            process.wait(3)
            self._show_info_message("Ended", f"Process {pid} terminated.")
            self.refresh_app_list()
        except Exception as error:
            messagebox.showerror("Error", f"Failed to end task: {error}")

    def _require_admin(self, action_name: str) -> bool:
        if self.is_admin:
            return True
        messagebox.showwarning(
            "Administrator Required",
            f"{action_name} requires administrator rights.\n"
            "Run the tool as Administrator to manage firewall rules.",
        )
        return False

    def _after_firewall_action(self) -> None:
        self.refresh_app_list()
        if self.ui_mode == "new" and hasattr(self, "target_tree"):
            self.target_tree.selection_remove(self.target_tree.selection())
            self._update_new_ui_action_button()
        else:
            self.tree.selection_remove(self.tree.selection())
            self.select_all_btn.config(text="Select all apps")

    def _show_info_message(self, title: str, message: str) -> None:
        if self.config_data.get("show_info_notifications", True):
            messagebox.showinfo(title, message)
        else:
            summary = message.splitlines()[0].strip() if message else title
            self.set_status(summary)

    def _text_input_has_focus(self) -> bool:
        widget = self.focus_get()
        if widget is None:
            return False
        return widget.winfo_class() in {"Entry", "TEntry", "Text", "Spinbox", "TCombobox", "Combobox"}

    def _bind_hotkey(self, sequence: str, callback, allow_when_typing: bool = False) -> None:
        sequence = (sequence or "").strip()
        if not sequence:
            return

        def handler(_event=None):
            if not allow_when_typing and self._text_input_has_focus():
                return None
            callback()
            return "break"

        try:
            self.bind_all(sequence, handler)
            self._bound_hotkeys.add(sequence)
        except tk.TclError:
            self.set_status(f"Invalid hotkey skipped: {sequence}")

    def _bind_hotkeys(self) -> None:
        for sequence in self._bound_hotkeys:
            self.unbind_all(sequence)
        self._bound_hotkeys.clear()

        if not self.config_data.get("hotkeys_enabled", True):
            return

        self._bind_hotkey(hotkey_config_to_sequence(self.config_data.get("hotkey_focus_search", "CTRL+F")), self._focus_search_hotkey, True)
        self._bind_hotkey(hotkey_config_to_sequence(self.config_data.get("hotkey_refresh", "VK_F5")), self.refresh_app_list, True)
        self._bind_hotkey(hotkey_config_to_sequence(self.config_data.get("hotkey_block_selected", "CTRL+B")), self.on_block)
        self._bind_hotkey(hotkey_config_to_sequence(self.config_data.get("hotkey_unblock_selected", "CTRL+U")), self.on_unblock)
        self._bind_hotkey(hotkey_config_to_sequence(self.config_data.get("hotkey_unblock_all", "CTRL+SHIFT+U")), self.on_unblock_all)
        self._bind_hotkey(hotkey_config_to_sequence(self.config_data.get("hotkey_toggle_selected", "CTRL+VK_RETURN")), self.on_toggle_target_block_state)

    def _focus_search_hotkey(self) -> None:
        if self.ui_mode == "new":
            self._open_process_selector()
            return
        if hasattr(self, "search_entry"):
            self.search_entry.focus_set()
            self.search_entry.selection_range(0, tk.END)

    def _format_app_lines(self, exes: list[str], max_items: int = 15) -> list[str]:
        if not exes:
            return []

        normalized = list(dict.fromkeys(exes))
        name_counts: dict[str, int] = {}
        for exe_path in normalized:
            name = os.path.basename(exe_path)
            name_counts[name] = name_counts.get(name, 0) + 1

        lines: list[str] = []
        for exe_path in normalized[:max_items]:
            name = os.path.basename(exe_path)
            if name_counts[name] > 1:
                lines.append(f"- {name} | {exe_path}")
            else:
                lines.append(f"- {name}")

        remaining = len(normalized) - max_items
        if remaining > 0:
            lines.append(f"- ...and {remaining} more")
        return lines

    def _show_action_summary(
        self,
        title: str,
        heading: str,
        exes: list[str],
        extra_sections: list[tuple[str, list[str]]] | None = None,
    ) -> None:
        lines = [heading]
        lines.extend(self._format_app_lines(exes))

        for section_title, section_exes in extra_sections or []:
            if not section_exes:
                continue
            lines.append("")
            lines.append(section_title)
            lines.extend(self._format_app_lines(section_exes))

        self._show_info_message(title, "\n".join(lines))

    def _get_auto_reconnect_seconds(self) -> int:
        try:
            return max(1, int(self.config_data.get("auto_reconnect_seconds", 60) or 60))
        except (TypeError, ValueError):
            return 60

    def _cancel_auto_reconnect(self, exe_path: str) -> None:
        job_id = self._auto_reconnect_jobs.pop(exe_path, None)
        if not job_id:
            return
        try:
            self.after_cancel(job_id)
        except Exception:
            pass

    def _cancel_all_auto_reconnects(self) -> None:
        for exe_path in list(self._auto_reconnect_jobs):
            self._cancel_auto_reconnect(exe_path)

    def _schedule_auto_reconnects(self, exes: list[str], delay_seconds: int | None = None) -> list[str]:
        delay_seconds = max(1, int(delay_seconds if delay_seconds is not None else self._get_auto_reconnect_seconds()))
        scheduled: list[str] = []
        for exe_path in dict.fromkeys(exes):
            self._cancel_auto_reconnect(exe_path)
            job_id = self.after(delay_seconds * 1000, lambda path=exe_path: self._auto_reconnect_exe(path))
            self._auto_reconnect_jobs[exe_path] = job_id
            scheduled.append(exe_path)
        return scheduled

    def _auto_reconnect_exe(self, exe_path: str) -> None:
        self._auto_reconnect_jobs.pop(exe_path, None)
        success, message = unblock_app(exe_path)
        self.refresh_app_list()
        app_name = os.path.basename(exe_path)

        if success:
            if "No matching firewall rules were found." in message:
                self.set_status(f"Auto reconnect skipped: {app_name} was already clear.")
            else:
                self.set_status(f"Auto reconnected {app_name}.")
            return

        messagebox.showerror("Auto reconnect failed", f"{exe_path}: {message}")
        self.set_status(f"Auto reconnect failed for {app_name}.")

    def on_toggle_target_block_state(self) -> None:
        selected = self.get_selected_exes()
        if not selected:
            self._show_info_message("Select an app", "Please select at least one application first.")
            return

        blocked = [exe for exe in selected if self._target_apps.get(exe, {}).get("blocked")]
        clear = [exe for exe in selected if exe not in blocked]

        if blocked and not clear:
            self.on_unblock()
        else:
            self.on_block()

    def on_toggle_selected_timer(self) -> None:
        self._start_or_stop_timers(
            self.get_selected_exes(),
            empty_title="Select an app",
            empty_message="Please select at least one application first.",
        )

    def on_toggle_all_target_apps(self) -> None:
        queue_exes = list(self._target_apps.keys())
        if not queue_exes:
            self._show_info_message("Block list", "Add at least one application to the block list first.")
            return

        blocked = [exe for exe in queue_exes if self._target_apps.get(exe, {}).get("blocked")]
        clear = [exe for exe in queue_exes if exe not in blocked]

        if blocked and not clear:
            self._unblock_exes(queue_exes, empty_title="Block list", empty_message="Add at least one application to the block list first.")
        else:
            self._block_exes(queue_exes, empty_title="Block list", empty_message="Add at least one application to the block list first.")

    def _block_exes(
        self,
        exes: list[str],
        *,
        empty_title: str,
        empty_message: str,
        timer_overrides: dict[str, int] | None = None,
        skipped_timer_exes: list[str] | None = None,
        disabled_timer_exes: list[str] | None = None,
    ) -> None:
        if not self._require_admin("Blocking"):
            return

        if not exes:
            self._show_info_message(empty_title, empty_message)
            return

        self.set_status(f"Blocking {len(exes)} app(s)...")
        failures: list[str] = []
        blocked_now: list[str] = []
        already_blocked: list[str] = []
        reconnect_started: list[str] = []

        for exe_path in exes:
            success, message = block_app(exe_path)
            if success and "already blocked" in message.lower():
                already_blocked.append(exe_path)
            elif success:
                blocked_now.append(exe_path)
            elif not success:
                failures.append(f"{exe_path}: {message}")

            if success and timer_overrides and exe_path in timer_overrides:
                self._schedule_auto_reconnects([exe_path], delay_seconds=timer_overrides[exe_path])
                reconnect_started.append(exe_path)

        self._after_firewall_action()

        if failures:
            messagebox.showerror("Failed", "\n".join(failures))
            self.set_status("Some apps could not be blocked.")
            return

        heading = f"Blocked {len(blocked_now)} app(s):" if blocked_now else "No new apps were blocked."
        extra_sections: list[tuple[str, list[str]]] = [("Already blocked:", already_blocked)]
        if reconnect_started:
            extra_sections.append(("Reconnect timer started:", reconnect_started))
        if skipped_timer_exes:
            extra_sections.append(("No saved reconnect timer:", skipped_timer_exes))
        if disabled_timer_exes:
            extra_sections.append(("Saved reconnect timer disabled:", disabled_timer_exes))
        self._show_action_summary("Blocked", heading, blocked_now, extra_sections=extra_sections)
        status_message = f"Blocked {len(exes)} app(s)."
        if reconnect_started:
            status_message += f" Reconnect timer started for {len(reconnect_started)} app(s)."
        self.set_status(status_message)

    def _unblock_exes(self, exes: list[str], *, empty_title: str, empty_message: str) -> None:
        if not self._require_admin("Unblocking"):
            return

        if not exes:
            self._show_info_message(empty_title, empty_message)
            return

        for exe_path in exes:
            self._cancel_auto_reconnect(exe_path)

        self.set_status(f"Unblocking {len(exes)} app(s)...")
        failures: list[str] = []
        unblocked: list[str] = []
        already_clear: list[str] = []

        for exe_path in exes:
            success, message = unblock_app(exe_path)
            if not success:
                failures.append(f"{exe_path}: {message}")
            elif "No matching firewall rules were found." in message:
                already_clear.append(exe_path)
            else:
                unblocked.append(exe_path)

        self._after_firewall_action()

        if failures:
            messagebox.showerror("Failed", "\n".join(failures))
            self.set_status("Some apps could not be unblocked.")
            return

        heading = f"Unblocked {len(unblocked)} app(s):" if unblocked else "No new apps were unblocked."
        self._show_action_summary("Unblocked", heading, unblocked, extra_sections=[("Already unblocked:", already_clear)])
        self.set_status(f"Processed {len(exes)} unblock request(s).")

    def on_block(self) -> None:
        exes = self.get_selected_exes()
        self._block_exes(exes, empty_title="Select an app", empty_message="Please select at least one application first.")

    def on_unblock(self) -> None:
        exes = self.get_selected_exes()
        self._unblock_exes(exes, empty_title="Select an app", empty_message="Please select at least one application first.")

    def on_unblock_all(self) -> None:
        if not self._require_admin("Unblocking all apps"):
            return

        blocked_programs = list_blocked_programs()
        if blocked_programs is None:
            messagebox.showerror("Failed", "Could not read the current blocked app list from Windows Firewall.")
            self.set_status("Failed to load blocked apps.")
            return

        if not blocked_programs:
            self._show_info_message("Unblock all apps", "No blocked apps were found.")
            self.set_status("No blocked apps found.")
            return

        self.set_status(f"Unblocking all blocked apps ({len(blocked_programs)})...")
        failures: list[str] = []
        unblocked: list[str] = []
        self._cancel_all_auto_reconnects()

        for exe_path in blocked_programs:
            success, message = unblock_app(exe_path)
            if success:
                unblocked.append(exe_path)
            else:
                failures.append(f"{exe_path}: {message}")

        self._after_firewall_action()

        if failures:
            messagebox.showerror("Failed", "\n".join(failures))
            self.set_status("Some blocked apps could not be unblocked.")
            return

        self._show_action_summary(
            "Unblocked All",
            f"Unblocked {len(unblocked)} app(s):",
            unblocked,
        )
        self.set_status(f"Unblocked {len(unblocked)} app(s).")

    def on_block_all_except(self) -> None:
        if not self._require_admin("Block-all mode"):
            return

        items = self.get_selected_items()
        if not items:
            self._show_info_message("Select an app", "Please select at least one application first.")
            return

        if self.config_data.get("confirm_block_all_except", True):
            confirmed = messagebox.askyesno(
                "Block All Except",
                "This will block all outbound traffic except for the selected app(s).\n"
                "Continue?",
            )
            if not confirmed:
                return

        self.set_status(f"Blocking all outbound traffic except {len(items)} app(s)...")

        failures: list[str] = []
        for item in items:
            success, message = allow_app(item["exe"])
            if not success:
                failures.append(f"{item['exe']}: {message}")

        blocked_all, block_message = block_all()
        self._after_firewall_action()

        if failures or not blocked_all:
            if not blocked_all:
                failures.append(f"Block-all rule: {block_message}")
            messagebox.showerror("Failed", "\n".join(failures))
            self.set_status("Block-all-except action failed.")
            return

        self._show_info_message(
            "Blocked All Except",
            "All outbound traffic is blocked except for the selected app(s).",
        )
        self.set_status("Block-all mode is active.")

    def on_stop_block_all(self) -> None:
        if not self._require_admin("Stopping block-all mode"):
            return

        self.set_status("Stopping block-all mode...")
        unblock_ok, unblock_message = unblock_all()

        selected_items = self.get_selected_items()
        allow_failures: list[str] = []
        for item in selected_items:
            success, message = remove_allow_rule(item["exe"])
            if not success:
                allow_failures.append(f"{item['exe']}: {message}")

        self._after_firewall_action()

        if not unblock_ok or allow_failures:
            errors = []
            if not unblock_ok:
                errors.append(f"Block-all rule: {unblock_message}")
            errors.extend(allow_failures)
            messagebox.showerror("Failed", "\n".join(errors))
            self.set_status("Could not fully stop block-all mode.")
            return

        message = "Stopped blocking all outbound traffic."
        if selected_items:
            message += "\nRemoved allow rules for the selected app(s)."
        else:
            message += "\nSelect allowed apps before stopping if you also want their allow rules removed."
        self._show_info_message("Stopped", message)
        self.set_status("Block-all mode stopped.")

    def on_run_as_admin(self) -> None:
        if relaunch_as_admin():
            self._show_info_message(
                "Elevated",
                "A new administrator window has been started.\n"
                "This window will now close.",
            )
            self.destroy()
            sys.exit(0)

        messagebox.showerror("Failed", "Could not launch the script as Administrator.")

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def on_close(self) -> None:
        self._stop_auto_refresh()
        self._cancel_all_auto_reconnects()
        self._save_autosave_preset()
        for sequence in self._bound_hotkeys:
            self.unbind_all(sequence)
        self._bound_hotkeys.clear()
        if self.config_data.get("remember_window_geometry", True):
            self.config_data["window_geometry"] = self.geometry().split("+", 1)[0]
            save_config(self.config_data)
        self.destroy()


def hide_console_window() -> None:
    try:
        import ctypes

        handle = ctypes.windll.kernel32.GetConsoleWindow()
        if handle:
            ctypes.windll.user32.ShowWindow(handle, 0)
    except Exception:
        pass


def main() -> None:
    config = load_config()

    if config.get("hide_console", False):
        hide_console_window()

    if config.get("ask_for_admin_on_startup", True) and not is_admin():
        root = tk.Tk()
        root.withdraw()
        run_as_admin = messagebox.askyesno(
            "Administrator rights",
            "Run with administrator rights?\n\n"
            "Yes = restart as admin now\n"
            "No = open in view-only mode",
        )
        root.destroy()

        if run_as_admin and relaunch_as_admin():
            return

    if not ensure_psutil():
        messagebox.showerror(
            "Missing requirement",
            "psutil is required for this tool.\n"
            "Install it manually with:\npython -m pip install psutil",
        )
        return

    admin_mode = is_admin()
    app = App(is_admin_user=admin_mode, config=config)

    if not admin_mode and config.get("warn_when_not_admin", True):
        messagebox.showwarning(
            "View-only mode",
            "The app is running without administrator rights.\n"
            "You can still view processes and rule status, but block and unblock actions are disabled.",
        )

    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        try:
            messagebox.showerror("Fatal error", str(error))
        except Exception:
            pass
        raise

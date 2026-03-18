# Internet Disconnector

Internet Disconnector is a Windows app that lets you block or unblock internet access for running `.exe` files by using Windows Firewall.

## Run

```bash
python internet_disconnector.py
```

Run it as Administrator if you want to block or unblock apps.

## Main Use

- `old` UI: select app(s) from the running list, then click `Block` or `Unblock`
- `new` UI: use `Add App`, build your block list, then use the buttons on the left and bottom
- Right-click any app in either UI for more actions

## Presets

- Open `Settings`
- Use `Save preset` to save the current app list
- Use `Load preset` to load it back
- Turn on `Auto save/load preset` if you want the app to remember the last preset automatically

## Reconnect Timer

- Right-click an app
- Click `Set reconnect timer...`
- This only saves the timer for that app, it does not start the timer yet
- Use `Block + reconnect timer` to block that app and start its saved timer
- Use `Disable saved timer` or `Enable saved timer` to turn the saved timer off or on for one app
- In the new UI, the button under `Block Selected` can start or stop the timer for the selected queued apps
- `Stop reconnect timer` only stops the running timer
- `Clear saved timer` removes the saved timer value for that app

## Config

The app uses `internet_disconnector.config.json`.

Simple example:

```json
{
  "ui_mode": "new",
  "hide_console": true,
  "auto_refresh": false,
  "refresh_interval": 5,
  "auto_reconnect_seconds": 60,
  "saved_reconnect_timers": {},
  "disabled_reconnect_timers": [],
  "preset_autosave": false,
  "only_active_default": false,
  "show_block_status": true,
  "show_info_notifications": true,
  "ask_for_admin_on_startup": true,
  "warn_when_not_admin": true,
  "confirm_block_all_except": true,
  "remember_window_geometry": true,
  "max_cached_icons": 192,
  "hotkeys_enabled": true,
  "hotkey_focus_search": "CTRL+F",
  "hotkey_refresh": "VK_F5",
  "hotkey_block_selected": "CTRL+B",
  "hotkey_unblock_selected": "CTRL+U",
  "hotkey_unblock_all": "CTRL+SHIFT+U",
  "hotkey_toggle_selected": "CTRL+VK_RETURN",
  "about_github_url": "https://github.com/your-github-link-here",
  "window_geometry": "460x360",
  "settings_window_geometry": "520x620"
}
```

Useful config keys:

- `ui_mode`: `old` or `new`
- `auto_refresh`: refresh the running app list automatically
- `refresh_interval`: refresh delay in seconds
- `auto_reconnect_seconds`: default value shown when you save a reconnect timer
- `saved_reconnect_timers`: saved per-app timers
- `disabled_reconnect_timers`: apps whose saved timer is turned off
- `preset_autosave`: save the current preset automatically and load it on next start
- `only_active_default`: start with only active network apps
- `hotkeys_enabled`: turn hotkeys on or off
- `window_geometry`: main window size like `460x360`

## Hotkeys

You can set hotkeys in the Settings window by pressing the keys directly.

You can also edit them in `internet_disconnector.config.json`.

Format:

- one key: `A`
- modifier + key: `CTRL+U`
- modifier + virtual key: `CTRL+VK_HOME`
- function key: `VK_F5`

Examples:

- `CTRL+F`
- `CTRL+SHIFT+U`
- `CTRL+VK_RETURN`
- `ALT+VK_INSERT`
- `VK_HOME`

Windows Virtual-Key Codes:

- Microsoft list: https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
- Examples:
  - `VK_HOME`
  - `VK_LEFT`
  - `VK_UP`
  - `VK_RIGHT`
  - `VK_DOWN`
  - `VK_INSERT`
  - `VK_DELETE`
  - `A`
  - `B`
  - `C`
  - `0`
  - `1`

## Notes

- The app creates outbound firewall rules
- Block status loading uses one firewall snapshot for better speed
- Right-click menu actions include block, unblock, copy path, open folder, and reconnect timer tools
- Presets can be saved and loaded from the Settings window

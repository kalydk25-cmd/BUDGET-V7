import os
import shutil
import sys
from pathlib import Path

import cost_calc_v95 as base


APP_SUPPORT_DIRNAME = "RETEC-CostCalc-V95-Mac"
TASK_CONFIG_NAME = "task_config_overrides.json"
RATIO_FILE_NAME = "dept_ratios.json"


def is_macos():
    return sys.platform == "darwin"


def app_support_dir():
    return Path.home() / "Library" / "Application Support" / APP_SUPPORT_DIRNAME


def user_data_path(filename):
    return app_support_dir() / filename


def bundled_resource_path(filename):
    return Path(base.resource_path(filename))


def ensure_user_data_dir():
    app_support_dir().mkdir(parents=True, exist_ok=True)


def seed_user_file(filename):
    source = bundled_resource_path(filename)
    target = user_data_path(filename)
    if target.exists() or not source.exists():
        return
    shutil.copy2(source, target)


def task_config_search_paths():
    user_path = user_data_path(TASK_CONFIG_NAME)
    bundled_path = bundled_resource_path(TASK_CONFIG_NAME)

    paths = [str(user_path)]
    if os.path.abspath(str(bundled_path)) != os.path.abspath(str(user_path)):
        paths.append(str(bundled_path))
    return paths


def task_config_save_path():
    return str(user_data_path(TASK_CONFIG_NAME))


def mousewheel_units(event):
    delta = getattr(event, "delta", 0)

    if is_macos():
        if delta == 0:
            return 0
        return -1 if delta > 0 else 1

    if base.platform.system() == "Windows":
        if delta == 0:
            return 0
        return int(-1 * (delta / 120))

    event_num = getattr(event, "num", None)
    if event_num == 4:
        return -1
    if event_num == 5:
        return 1
    return 0


def patched_checkbox_mousewheel(self, event):
    units = mousewheel_units(event)
    if units:
        self.canvas.yview_scroll(units, "units")


def patched_section_mousewheel(self, event):
    units = mousewheel_units(event)
    if units:
        self.canvas.yview_scroll(units, "units")


def patch_mac_fonts():
    base.UI_FONT = ("PingFang SC", 11)
    base.UI_FONT_BOLD = ("PingFang SC", 11, "bold")
    base.UI_FONT_SMALL = ("PingFang SC", 10)
    base.UI_FONT_TITLE = ("PingFang SC", 18, "bold")


def configure_mac_runtime():
    if not is_macos():
        return

    ensure_user_data_dir()
    seed_user_file(TASK_CONFIG_NAME)

    base.RATIO_FILE = str(user_data_path(RATIO_FILE_NAME))
    base.get_task_config_search_paths = task_config_search_paths
    base.get_task_config_save_path = task_config_save_path
    base.CURRENT_RATIO_DB = base.load_json(base.RATIO_FILE, base.DEFAULT_RATIOS)
    base.load_task_config_overrides_from_file()

    base.ScrollableCheckBoxFrame.on_mousewheel = patched_checkbox_mousewheel
    base.SectionScrollArea._on_mousewheel = patched_section_mousewheel
    patch_mac_fonts()


def main():
    configure_mac_runtime()
    root = base.tk.Tk()
    base.App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

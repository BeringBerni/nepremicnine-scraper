"""
launcher.py – Zaženjalnik za Nepremičnine GUI
==============================================
Samodejno namesti vse potrebne knjižnice in zažene gui.py.

Zagon:
    py launcher.py
Ali pa použij Nepremicnine_GUI.exe (zgrajen z Build_EXE.bat).
"""

import sys
import os
import subprocess

# ── Barve za konzolo ──────────────────────────────────────────────────────────
def _c(code, text):
    """ANSI barva (deluje v Windows Terminal / PowerShell)."""
    return f"\033[{code}m{text}\033[0m"

GREEN  = lambda t: _c("92", t)
YELLOW = lambda t: _c("93", t)
RED    = lambda t: _c("91", t)
CYAN   = lambda t: _c("96", t)
BOLD   = lambda t: _c("1",  t)

# ── Zahtevane knjižnice: {ime_za_import: ime_za_pip} ─────────────────────────
REQUIRED: dict[str, str] = {
    "DrissionPage": "DrissionPage",
    "bs4":          "beautifulsoup4",
    "lxml":         "lxml",
    "docx":         "python-docx",
}

# ── Preveri in namesti manjkajoče knjižnice ───────────────────────────────────
def check_and_install():
    missing = []
    print(BOLD("\n══════════════════════════════════════════════════"))
    print(BOLD("   🏠  Nepremičnine GUI  –  Zaženjalnik"))
    print(BOLD("══════════════════════════════════════════════════"))
    print(f"   Python: {CYAN(sys.version.split()[0])}")
    print()

    print("   Preverjam knjižnice …")
    for import_name, pip_name in REQUIRED.items():
        try:
            __import__(import_name)
            print(f"   {GREEN('✓')}  {pip_name}")
        except ImportError:
            print(f"   {YELLOW('○')}  {pip_name}  (manjka – bo nameščena)")
            missing.append(pip_name)

    if missing:
        print()
        print(f"   {YELLOW('⚙  Nameščam:')} {', '.join(missing)}")
        print()
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "--upgrade"] + missing,
                stdout=sys.stdout,
                stderr=sys.stderr,
            )
            print()
            print(f"   {GREEN('✓  Vse knjižnice so uspešno nameščene!')}")
        except subprocess.CalledProcessError as e:
            print()
            print(f"   {RED('✗  Napaka pri namestitvi:')} {e}")
            print(f"   {YELLOW('➜  Poskusi ročno:')} py -m pip install {' '.join(missing)}")
            input("\n   Pritisni Enter za izhod …")
            sys.exit(1)
    else:
        print()
        print(f"   {GREEN('✓  Vse knjižnice so že nameščene.')}")

    print()

# ── Poišči gui.py ─────────────────────────────────────────────────────────────
def find_gui() -> str:
    # Ko teče kot PyInstaller .exe, je sys.frozen = True
    if getattr(sys, "frozen", False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    gui_path = os.path.join(script_dir, "gui.py")
    if not os.path.isfile(gui_path):
        print(f"   {RED('✗  gui.py ni najdena v:')} {script_dir}")
        input("\n   Pritisni Enter za izhod …")
        sys.exit(1)
    return gui_path

# ── Zaženi gui.py ─────────────────────────────────────────────────────────────
def launch_gui(gui_path: str):
    print(f"   {GREEN('▶  Zaganjam:')} gui.py")
    print()

    # Zapremo konzolo in zaženemo GUI (ne čakamo na konec)
    try:
        if sys.platform == "win32":
            # DETACHED_PROCESS – konzolno okno se zapre takoj ko se GUI odpre
            DETACHED = 0x00000008
            CREATE_NO_WINDOW = 0x08000000
            subprocess.Popen(
                [sys.executable, "-X", "utf8", gui_path],
                creationflags=DETACHED | CREATE_NO_WINDOW,
                close_fds=True,
            )
        else:
            subprocess.Popen([sys.executable, "-X", "utf8", gui_path])
    except Exception as e:
        print(f"   {RED('✗  Napaka pri zagonu gui.py:')} {e}")
        input("\n   Pritisni Enter za izhod …")
        sys.exit(1)


# ── Vstopna točka ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Omogoči UTF-8 v konzoli
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    # Omogoči ANSI barve v starejšem Windows cmd
    os.system("color")

    # Ko teče kot PyInstaller .exe, ne kličemo subprocess(sys.executable, gui.py)
    # ker bi poklicali sam sebe.  Namesto tega uvozimo in zaženemo gui neposredno.
    if getattr(sys, "frozen", False):
        check_and_install()
        try:
            import gui as _gui
            _gui.ScraperGUI().mainloop()
        except Exception as e:
            print(f"   {RED('✗  Napaka pri zagonu GUI:')} {e}")
            import traceback; traceback.print_exc()
            input("\n   Pritisni Enter za izhod …")
            sys.exit(1)
    else:
        check_and_install()
        gui_path = find_gui()
        launch_gui(gui_path)


#!/usr/bin/env python3
# pylint: disable=line-too-long
# Copyright (C) 2022-2026 The MIO-KITCHEN-SOURCE Project
#
# Licensed under the GNU AFFERO GENERAL PUBLIC LICENSE, Version 3.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      https://www.gnu.org/licenses/agpl-3.0.en.html#license-text
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
import json
import os
import platform
import shutil
import subprocess
import sys
import zipfile
from platform import system
import importlib.util
from pip._internal.cli.main import main as _main

class Builder:
    def __init__(self):
        ostype = system()
        if ostype == 'Linux':
            name = 'MIO-KITCHEN-linux.zip'
        elif ostype == 'Darwin':
            if platform.machine() == 'x86_64':
                name = 'MIO-KITCHEN-macos-intel.zip'
            else:
                name = 'MIO-KITCHEN-macos.zip'
        else:
            name = 'MIO-KITCHEN-win.zip'
        self.name = name
        self.local = os.getcwd()
        self.ostype = ostype

    def build(self):
        print('Building...')
        self.install_package()
        self.unit_test()
        self.patch_libs()
        self.fix_pywin32()
        self.pyinstaller_build()
        self.config_folder()
        self.pack_zip(f'{self.local}/dist', self.name)

    def fix_pywin32(self):
        # pywin32 installed via pip needs its post-install step run so that
        # pywintypes/pythoncom are registered and importable by *any* python
        # process (including the isolated child processes PyInstaller spawns
        # to inspect hidden imports). Without this, PyInstaller's
        # hook-pythoncom.py fails with:
        #   ModuleNotFoundError: No module named 'pywintypes'
        if os.name != 'nt':
            return
        print("Fixing pywin32 (running post-install step)")

        import site
        import glob
        site_packages = [p for p in site.getsitepackages() if os.path.isdir(p)]
        try:
            user_site = site.getusersitepackages()
            if user_site and os.path.isdir(user_site):
                site_packages.append(user_site)
        except Exception:
            pass

        # 0) Make sure pywin32's own .pth file is actually in place and
        # correct. This file is what makes every *future* python process
        # (including PyInstaller's isolated child workers, which inherit
        # environment but not our sys.path edits) automatically add
        # win32/win32\lib/pythonwin to sys.path on startup - which is what
        # ultimately fixes the hook failure, not just this process.
        pth_content = (
            "# .pth file for the PyWin32 extensions (restored by build.py)\n"
            "win32\n"
            "win32\\lib\n"
            "pythonwin\n"
        )
        for base in site_packages:
            if os.path.isdir(os.path.join(base, "win32")):
                pth_path = os.path.join(base, "pywin32.pth")
                try:
                    write_it = True
                    if os.path.exists(pth_path):
                        with open(pth_path, "r", encoding="utf-8", errors="ignore") as f:
                            if "win32\\lib" in f.read():
                                write_it = False
                    if write_it:
                        with open(pth_path, "w", encoding="utf-8") as f:
                            f.write(pth_content)
                        print(f"Wrote {pth_path}")
                except Exception as e:
                    print(f"Could not write pywin32.pth at {pth_path}: {e}")

        # 1) Try the real post-install script, wherever it landed (search
        # broadly since its location varies by pip/pywin32 version).
        candidates = []
        for base in site_packages:
            candidates.append(os.path.join(base, "win32", "scripts", "pywin32_postinstall.py"))
            candidates.append(os.path.join(base, "Scripts", "pywin32_postinstall.py"))
            candidates.append(os.path.normpath(os.path.join(base, "..", "Scripts", "pywin32_postinstall.py")))
            candidates.extend(glob.glob(os.path.join(base, "pywin32*.data", "scripts", "pywin32_postinstall.py")))
        # also search near the interpreter (Scripts dir next to python.exe)
        py_dir = os.path.dirname(sys.executable)
        candidates.append(os.path.join(py_dir, "Scripts", "pywin32_postinstall.py"))

        ran_postinstall = False
        for script in candidates:
            if script and os.path.exists(script):
                try:
                    subprocess.run([sys.executable, script, "-install", "-silent"], check=True)
                    ran_postinstall = True
                    print(f"Ran pywin32 post-install script: {script}")
                    break
                except Exception as e:
                    print(f"Running {script} failed: {e}")
        if not ran_postinstall:
            try:
                subprocess.run([sys.executable, "-m", "pywin32_postinstall", "-install", "-silent"], check=True)
                ran_postinstall = True
            except Exception as e:
                print(f"pywin32_postinstall via -m failed ({e})")

        # 2) Regardless of the above, make sure the pywin32 extension
        # directories are on sys.path AND on PATH for this process and any
        # child processes (PyInstaller's isolated workers included), since
        # that's what actually lets `import pywintypes` succeed and lets the
        # loader find the accompanying DLLs. Critically this must include
        # pywin32_system32, where pythoncomXXX.dll / pywintypesXXX.dll
        # actually live, and win32/lib, where the pywintypes.py shim lives.
        pywin32_dirs = []
        for base in site_packages:
            for sub in ("win32", os.path.join("win32", "lib"), "win32com", "win32comext",
                        "pywin32_system32", "Pythonwin"):
                d = os.path.join(base, sub)
                if os.path.isdir(d):
                    pywin32_dirs.append(d)

        for d in pywin32_dirs:
            if d not in sys.path:
                sys.path.insert(0, d)

        if pywin32_dirs:
            os.environ["PATH"] = os.pathsep.join(pywin32_dirs) + os.pathsep + os.environ.get("PATH", "")
            # Python 3.8+ ignores PATH for DLL resolution unless the
            # directory is registered explicitly.
            if hasattr(os, "add_dll_directory"):
                for d in pywin32_dirs:
                    try:
                        os.add_dll_directory(d)
                    except (OSError, ValueError):
                        pass

        # 2b) Belt-and-braces: also copy the pywin32 DLLs next to the
        # interpreter (alongside python.exe). This directory is always on
        # the DLL search path for every process using this interpreter,
        # including PyInstaller's isolated child workers, regardless of
        # PATH/env inheritance quirks.
        for base in site_packages:
            dll_dir = os.path.join(base, "pywin32_system32")
            if os.path.isdir(dll_dir):
                for dll in glob.glob(os.path.join(dll_dir, "*.dll")):
                    try:
                        dest = os.path.join(py_dir, os.path.basename(dll))
                        if not os.path.exists(dest):
                            shutil.copy(dll, dest)
                            print(f"Copied {dll} -> {dest}")
                    except Exception as e:
                        print(f"Could not copy {dll}: {e}")

        # 3) Verify in a *fresh subprocess* (not this process). This is the
        # only reliable way to know whether PyInstaller's own isolated child
        # workers will be able to import pythoncom/pywintypes too, since
        # they're fresh interpreters just like this check is.
        check = subprocess.run(
            [sys.executable, "-c", "import pywintypes, pythoncom, win32com; print('OK')"],
            capture_output=True, text=True,
        )
        if check.returncode == 0 and "OK" in check.stdout:
            print("pywin32 is importable from a fresh interpreter, proceeding")
        else:
            print(f"WARNING: pywin32 still not importable from a fresh interpreter after fix attempts.\n"
                  f"stdout: {check.stdout}\nstderr: {check.stderr}")

    def patch_libs(self):
        print("Patching libs")
        spec = importlib.util.find_spec('qfluentwidgets')
        target_file = os.path.join(os.path.dirname(spec.origin), "common", "config.py")
        with open(target_file, "r", encoding="utf-8", newline="\n") as f:
            data = f.readlines()
            data_index = 14
            data[data_index] = "ALERT = None"
        with open(target_file, "w", encoding="utf-8", newline="\n") as f:
            f.writelines(data)

    def run_command(self, command: list[str], strip: bool = False):
        try:
            result = subprocess.run(command, capture_output=True, text=True, check=True)
            return result.stdout.strip() if strip else result.stdout
        except subprocess.CalledProcessError:
            return None

    def generate_release_body(self):
        print('Generating Release Body...')
        # load config
        with open('bin/settings.json', 'r', encoding='utf-8') as f:
            ver = json.load(f)
            ver = ver['Tool']['Version']
        with open('body.md', 'w', encoding='utf-8', newline='\n') as f:
            f.write(f"Build times: {os.getenv('GITHUB_RUN_NUMBER')}\n")
            f.write(f"Actor: {os.getenv('GITHUB_TRIGGERING_ACTOR')}\n")
            f.write(f"Repository: {os.getenv('GITHUB_REPOSITORY')}\n")
            f.write(f'Version: {ver}\n')
            f.write(f'Changelog:\n')
            f.write(f'```\n')
            head = self.run_command(['git', 'rev-parse', 'HEAD'], strip=True)
            f.write(self.run_command(['git', "log", "-1", "--pretty=%B", head]))
            f.write(f'```\n')

    def move_artifacts(self):
        with open('bin/settings.json', 'r', encoding='utf-8') as f:
            ver = json.load(f)
            ver = ver['Tool']['Version']
        for i in ['MIO-KITCHEN-win', 'MIO-KITCHEN-linux', 'MIO-KITCHEN-macos', 'MIO-KITCHEN-macos-intel']:
            name_list = i.rsplit('-')
            name_list.insert(2, ver)
            name = '-'.join(name_list)
            os.rename(f'{i}/{i}.zip', f'{name}.zip')
        # write ver to github env
        with open(os.getenv('GITHUB_ENV'), 'a', encoding='utf-8') as f:
            f.write(f'ver={ver}\n')

    def unit_test(self):
        from src.tool_tester import test_main, Test

        if Test:
            test_main(exit=False)

    def install_package(self):
        with open('requirements.txt', 'r', encoding='utf-8') as l:
            for i in l.read().split("\n"):
                print(f"Installing {i}")
                _main(['install', i])
        for ext in os.listdir("src/c_extension"):
            print(f"[Ext] Installing {ext}")
            _main(['install', f"src/c_extension/{ext}"])

    def pyinstaller_build(self):
        import PyInstaller.__main__
        if self.ostype == 'Darwin':
            PyInstaller.__main__.run([
                'tool.py',
                '-Dw',
                '--exclude-module', 'tkinter',
                '--exclude-module',
                'numpy',
                '-i',
                'icon.ico',
                '--collect-data',
                'androguard',
                '--hidden-import',
                'PIL',
            ])
        elif os.name == 'posix':

            PyInstaller.__main__.run([
                'tool.py',
                '-Dw',
                '--exclude-module', 'tkinter',
                '--exclude-module',
                'numpy',
                '-i',
                'icon.ico',
                '--collect-data',
                'androguard',
                '--hidden-import',
                'PIL',
                '--splash',
                'splash_loongarch.png' if platform.machine() == 'loongarch64' else 'splash.png'
            ])
        elif os.name == 'nt':
            PyInstaller.__main__.run([
                'tool.py',
                '-Dw',
                '--exclude-module', 'numpy',
                '--exclude-module', 'tkinter',
                '-i',
                'icon.ico',
                '--collect-data',
                'androguard',
                '--collect-all', 'win32',
                '--collect-all', 'win32api',
                "--hidden-import", "win32api",
                "--hidden-import", "win32com",
                "--hidden-import", "win32",
                "--hidden-import", "win32timezone",
                "--hidden-import", "pywintypes",
                "--hidden-import", "pythoncom",
                '--splash',
                'splash.png'
            ])

    def config_folder(self):
        if not os.path.exists('dist/bin'):
            os.makedirs('dist/bin', exist_ok=True)
        while_list = ['images', 'languages', 'licenses', 'module', 'temp', 'extra_flash', 'settings.json', self.ostype,
                      'kemiaojiang.png', 'License_kemiaojiang.txt', 'help_document.json', "exec.sh", 'update.json']
        for i in os.listdir(self.local + "/bin"):
            if i in while_list:
                if os.path.isdir(f"{self.local}/bin/{i}"):
                    shutil.copytree(f"{self.local}/bin/{i}", f"{self.local}/dist/bin/{i}", dirs_exist_ok=True)
                else:
                    shutil.copy(f"{self.local}/bin/{i}", f"{self.local}/dist/bin/{i}")
        if not os.path.exists('dist/LICENSE'):
            shutil.copy(f'{self.local}/LICENSE', f"{self.local}/dist/LICENSE")

        if os.name == 'posix':
            if platform.machine() == 'x86_64' and os.path.exists(f'{self.local}/dist/bin/Linux/aarch64'):
                try:
                    shutil.rmtree(f'{self.local}/dist/bin/Linux/aarch64')
                except Exception as e:
                    print(e)
            for root, dirs, files in os.walk(f'{self.local}/dist/bin', topdown=True):
                for i in files:
                    print(f"Chmod {os.path.join(root, i)}")
                    os.chmod(os.path.join(root, i), 0o7777, follow_symlinks=False)
        os.rename(f'{self.local}/dist/tool', f'{self.local}/dist/tool_built')
        shutil.copytree(f'{self.local}/dist/tool_built', f'{self.local}/dist', dirs_exist_ok=True)
        shutil.rmtree(f'{self.local}/dist/tool_built')

    def pack_zip(self, source, name):
        abs_folder_path = os.path.abspath(source)
        zip_file_path = os.path.join(self.local, name)
        with zipfile.ZipFile(zip_file_path, "w", zipfile.ZIP_DEFLATED) as archive:
            for root, _, files in os.walk(abs_folder_path):
                for file in files:
                    if file == name:
                        continue
                    file_path = os.path.join(root, file)
                    if ".git" in file_path:
                        continue
                    print(f"Adding: {file_path}")
                    archive.write(file_path, os.path.relpath(file_path, abs_folder_path))
        print("Pack Zip Done!")


if __name__ == '__main__':
    if len(sys.argv) == 1:
        builder = Builder()
        builder.build()
    else:
        # Generate Release Body
        if sys.argv[1] == 'grb':
            builder = Builder()
            builder.generate_release_body()
        elif sys.argv[1] == 'ma':
            builder = Builder()
            builder.move_artifacts()
        else:
            print('Usage:')
            print('To Build Binary and Pack')
            print('\tpython build.py')
            print('To Move artifacts to local folder')
            print('\tpython build.py ma')
            print('To Generate Release Body')
            print('\tpython build.py grb')

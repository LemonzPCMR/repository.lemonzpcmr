""" 
    Put this script in the root folder of your repo and it will
    zip up all addon folders, create a new zip in your zips folder
    and then update the md5 and addons.xml file
"""

import hashlib
import os
import shutil
import sys
import zipfile

from xml.etree import ElementTree

SCRIPT_VERSION = 5
KODI_VERSIONS = ["krypton", "leia", "matrix", "nexus", "repo"]
IGNORE = [
    ".git",
    ".github",
    ".gitignore",
    ".DS_Store",
    "thumbs.db",
    ".idea",
    "venv",
]
_COLOR_ESCAPE = "\x1b[{}m"
_COLORS = {
    "black": "30",
    "red": "31",
    "green": "4;32",
    "yellow": "3;33",
    "blue": "34",
    "magenta": "35",
    "cyan": "1;36",
    "grey": "37",
    "endc": "0",
}


def _setup_colors():
    """
    Return True if the running system's terminal supports color,
    and False otherwise.
    """

    def vt_codes_enabled_in_windows_registry():
        """
        Check the Windows registry to see if VT code handling has been enabled by default.
        """
        try:
            import winreg
        except:
            return False
        else:
            reg_key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, "Console", access=winreg.KEY_ALL_ACCESS
            )
            try:
                reg_key_value, _ = winreg.QueryValueEx(reg_key, "VirtualTerminalLevel")
            except FileNotFoundError:
                try:
                    winreg.SetValueEx(
                        reg_key, "VirtualTerminalLevel", 0, winreg.KEY_DWORD, 1
                    )
                except:
                    return False
                else:
                    reg_key_value, _ = winreg.QueryValueEx(
                        reg_key, "VirtualTerminalLevel"
                    )
            else:
                return reg_key_value == 1

    def is_a_tty():
        return hasattr(sys.stdout, "isatty") and sys.stdout.isatty()

    def legacy_support():
        console = 0
        color = 0
        if sys.platform in ["linux", "linux2", "darwin"]:
            pass
        elif sys.platform == "win32":
            color = os.system("color")

            from ctypes import windll

            k = windll.kernel32
            console = k.SetConsoleMode(k.GetStdHandle(-11), 7)

        return any([color == 1, console == 1])

    return any(
        [
            is_a_tty(),
            sys.platform != "win32",
            "ANSICON" in os.environ,
            "WT_SESSION" in os.environ,
            os.environ.get("TERM_PROGRAM") == "vscode",
            vt_codes_enabled_in_windows_registry(),
            legacy_support(),
        ]
    )


_SUPPORTS_COLOR = _setup_colors()


def color_text(text, color):
    """
    Return an ANSI-colored string, if supported.
    """

    return (
        '{}{}{}'.format(
            _COLOR_ESCAPE.format(_COLORS[color]),
            text,
            _COLOR_ESCAPE.format(_COLORS["endc"]),
        )
        if _SUPPORTS_COLOR
        else text
    )


def convert_bytes(num):
    """
    this function will convert bytes to MB.... GB... etc
    """
    for x in ['bytes', 'KB', 'MB', 'GB', 'TB']:
        if num < 1024.0:
            return "%3.1f %s" % (num, x)
        num /= 1024.0


class Generator:
    """
    Generates a new addons.xml file from each addons addon.xml file
    and a new addons.xml.md5 hash file. Must be run from the root of
    the checked-out repo.
    """

    def __init__(self, release):
        self.release_path = release
        self.zips_path = os.path.join(self.release_path, "zips")
        # Root of the whole git repo (one level above the Kodi-version folder)
        self.root_path = os.path.dirname(os.path.abspath(__file__))
        addons_xml_path = os.path.join(self.zips_path, "addons.xml")
        md5_path = os.path.join(self.zips_path, "addons.xml.md5")

        if not os.path.exists(self.zips_path):
            os.makedirs(self.zips_path)

        self._remove_binaries()

        if self._generate_addons_file(addons_xml_path):
            print(
                "Successfully updated {}".format(color_text(addons_xml_path, 'yellow'))
            )

            if self._generate_md5_file(addons_xml_path, md5_path):
                print("Successfully updated {}".format(color_text(md5_path, 'yellow')))

        self._cleanup_old_zips()
        self._copy_repo_zip_to_root()
        self._update_index_html()

    def _remove_binaries(self):
        """
        Removes any and all compiled Python files before operations.
        """

        for parent, dirnames, filenames in os.walk(self.release_path):
            for fn in filenames:
                if fn.lower().endswith("pyo") or fn.lower().endswith("pyc"):
                    compiled = os.path.join(parent, fn)
                    try:
                        os.remove(compiled)
                        print(
                            "Removed compiled python file: {}".format(
                                color_text(compiled, 'green')
                            )
                        )
                    except:
                        print(
                            "Failed to remove compiled python file: {}".format(
                                color_text(compiled, 'red')
                            )
                        )
            for dir in dirnames:
                if "pycache" in dir.lower():
                    compiled = os.path.join(parent, dir)
                    try:
                        shutil.rmtree(compiled)
                        print(
                            "Removed __pycache__ cache folder: {}".format(
                                color_text(compiled, 'green')
                            )
                        )
                    except:
                        print(
                            "Failed to remove __pycache__ cache folder:  {}".format(
                                color_text(compiled, 'red')
                            )
                        )

    def _get_addon_folders(self):
        """
        Returns a list of absolute paths to addon folders.
        Supports two layouts:
          1. Direct addon:  repo/<addon-id>/addon.xml
          2. Nested project: repo/<project>/repo/<addon-id>/addon.xml
             (i.e. the submodule is itself a repo-generator project)
        """
        results = []
        for item in os.listdir(self.release_path):
            item_path = os.path.join(self.release_path, item)
            if not os.path.isdir(item_path) or item == "zips" or item.startswith("."):
                continue

            # Case 1: direct addon with addon.xml at root
            if os.path.exists(os.path.join(item_path, "addon.xml")):
                results.append(item_path)

            # Case 2: nested repo-generator project — look inside its repo/ subfolder
            elif os.path.isdir(os.path.join(item_path, "repo")):
                nested_repo = os.path.join(item_path, "repo")
                for sub in os.listdir(nested_repo):
                    sub_path = os.path.join(nested_repo, sub)
                    if (
                        os.path.isdir(sub_path)
                        and sub != "zips"
                        and not sub.startswith(".")
                        and os.path.exists(os.path.join(sub_path, "addon.xml"))
                    ):
                        results.append(sub_path)
        return results

    def _create_zip(self, addon_folder, addon_id, version):
        """
        Creates a zip file in the zips directory for the given addon.
        addon_folder is the absolute path to the addon directory.
        """
        zip_folder = os.path.join(self.zips_path, addon_id)
        if not os.path.exists(zip_folder):
            os.makedirs(zip_folder)

        final_zip = os.path.join(zip_folder, "{0}-{1}.zip".format(addon_id, version))
        if not os.path.exists(final_zip):
            zip = zipfile.ZipFile(final_zip, "w", compression=zipfile.ZIP_DEFLATED)
            root_len = len(os.path.dirname(os.path.abspath(addon_folder)))

            for root, dirs, files in os.walk(addon_folder):
                # remove any unneeded artifacts
                for i in IGNORE:
                    if i in dirs:
                        try:
                            dirs.remove(i)
                        except:
                            pass
                    for f in files:
                        if f.startswith(i):
                            try:
                                files.remove(f)
                            except:
                                pass

                archive_root = os.path.abspath(root)[root_len:]

                for f in files:
                    fullpath = os.path.join(root, f)
                    archive_name = os.path.join(archive_root, f)
                    zip.write(fullpath, archive_name, zipfile.ZIP_DEFLATED)

            zip.close()
            size = convert_bytes(os.path.getsize(final_zip))
            print(
                "Zip created for {} ({}) - {}".format(
                    color_text(addon_id, 'cyan'),
                    color_text(version, 'green'),
                    color_text(size, 'yellow'),
                )
            )

    def _copy_meta_files(self, src_addon_folder, addon_id):
        """
        Copy the addon.xml and relevant art files into the zips folder.
        src_addon_folder is the absolute path to the addon directory.
        """

        tree = ElementTree.parse(os.path.join(src_addon_folder, "addon.xml"))
        root = tree.getroot()

        copyfiles = ["addon.xml"]
        for ext in root.findall("extension"):
            if ext.get("point") in ["xbmc.addon.metadata", "kodi.addon.metadata"]:
                assets = ext.find("assets")
                if not assets:
                    continue
                for art in [a for a in assets if a.text]:
                    copyfiles.append(os.path.normpath(art.text))

        dest_folder = os.path.join(self.zips_path, addon_id)
        for file in copyfiles:
            addon_path = os.path.join(src_addon_folder, file)
            if not os.path.exists(addon_path):
                continue

            zips_path = os.path.join(dest_folder, file)
            asset_path = os.path.split(zips_path)[0]
            if not os.path.exists(asset_path):
                os.makedirs(asset_path)

            shutil.copy(addon_path, zips_path)

    def _generate_addons_file(self, addons_xml_path):
        """
        Generates a zip for each found addon, and updates the addons.xml file accordingly.
        Supports both direct addon folders and nested repo-generator project submodules.
        """
        if not os.path.exists(addons_xml_path):
            addons_root = ElementTree.Element('addons')
            addons_xml = ElementTree.ElementTree(addons_root)
        else:
            addons_xml = ElementTree.parse(addons_xml_path)
            addons_root = addons_xml.getroot()

        addon_folders = self._get_addon_folders()

        addon_xpath = "addon[@id='{}']"
        changed = False
        for addon_path in addon_folders:
            try:
                addon_xml_path = os.path.join(addon_path, "addon.xml")
                addon_xml = ElementTree.parse(addon_xml_path)
                addon_root = addon_xml.getroot()
                id = addon_root.get('id')
                version = addon_root.get('version')

                updated = False
                addon_entry = addons_root.find(addon_xpath.format(id))
                if addon_entry is not None and addon_entry.get('version') != version:
                    index = addons_root.findall('addon').index(addon_entry)
                    addons_root.remove(addon_entry)
                    addons_root.insert(index, addon_root)
                    updated = True
                    changed = True
                elif addon_entry is None:
                    addons_root.append(addon_root)
                    updated = True
                    changed = True

                if updated:
                    self._create_zip(addon_path, id, version)
                    self._copy_meta_files(addon_path, id)
            except Exception as e:
                print(
                    "Excluding {}: {}".format(
                        color_text(addon_path, 'yellow'), color_text(e, 'red')
                    )
                )

        if changed:
            addons_root[:] = sorted(addons_root, key=lambda addon: addon.get('id'))
            try:
                addons_xml.write(
                    addons_xml_path, encoding="utf-8", xml_declaration=True
                )

                return changed
            except Exception as e:
                print(
                    "An error occurred updating {}!\n{}".format(
                        color_text(addons_xml_path, 'yellow'), color_text(e, 'red')
                    )
                )

    def _cleanup_old_zips(self):
        """
        Removes stale versioned zip files from each addon subfolder inside
        the zips directory, keeping only the most recently modified zip.
        """
        if not os.path.exists(self.zips_path):
            return
        for addon_id in os.listdir(self.zips_path):
            addon_zip_folder = os.path.join(self.zips_path, addon_id)
            if not os.path.isdir(addon_zip_folder):
                continue
            zips = sorted(
                [f for f in os.listdir(addon_zip_folder) if f.endswith(".zip")],
                key=lambda f: os.path.getmtime(os.path.join(addon_zip_folder, f)),
            )
            # Keep the newest, delete the rest
            for old_zip in zips[:-1]:
                old_path = os.path.join(addon_zip_folder, old_zip)
                try:
                    os.remove(old_path)
                    print("Removed old zip: {}".format(color_text(old_path, 'yellow')))
                except Exception as e:
                    print("Failed to remove {}: {}".format(color_text(old_path, 'red'), e))

    def _copy_repo_zip_to_root(self):
        """
        Copies the freshly built repository zip from zips/<repo-id>/ to the
        git root, removing any previously copied repo zips first.

        Priority: prefer the addon whose id matches this repo's own root
        folder name (e.g. repository.lemonzpcmr), so that bundled addons
        like repository.jurialmunkey are never promoted by accident.
        """
        if not os.path.exists(self.zips_path):
            return

        # The "home" repo id is the basename of the git root directory.
        home_id = os.path.basename(self.root_path)

        # Collect all candidates: list of (addon_id, zip_name, zip_src)
        candidates = []

        for addon_id in os.listdir(self.zips_path):
            addon_zip_folder = os.path.join(self.zips_path, addon_id)
            if not os.path.isdir(addon_zip_folder):
                continue
            addon_xml_path = os.path.join(addon_zip_folder, "addon.xml")
            if not os.path.exists(addon_xml_path):
                continue
            try:
                tree = ElementTree.parse(addon_xml_path)
                for ext in tree.getroot().findall("extension"):
                    if ext.get("point") == "xbmc.addon.repository":
                        zips = [
                            f for f in os.listdir(addon_zip_folder) if f.endswith(".zip")
                        ]
                        if zips:
                            newest = sorted(
                                zips,
                                key=lambda f: os.path.getmtime(
                                    os.path.join(addon_zip_folder, f)
                                ),
                            )[-1]
                            candidates.append(
                                (addon_id, newest, os.path.join(addon_zip_folder, newest))
                            )
                        break
            except Exception:
                continue

        if not candidates:
            print(color_text("No repository addon zip found to copy to root.", 'yellow'))
            return

        # Prefer the addon whose id matches this repo's own folder name.
        preferred = [c for c in candidates if c[0] == home_id]
        chosen_id, repo_zip_name, repo_zip_src = (preferred or candidates)[0]


        # Remove any old repo zips already sitting at root
        for f in os.listdir(self.root_path):
            if f.endswith(".zip"):
                old = os.path.join(self.root_path, f)
                try:
                    os.remove(old)
                    print("Removed old root zip: {}".format(color_text(old, 'yellow')))
                except Exception as e:
                    print("Failed to remove {}: {}".format(color_text(old, 'red'), e))

        dest = os.path.join(self.root_path, repo_zip_name)
        shutil.copy(repo_zip_src, dest)
        print("Copied {} to root: {}".format(
            color_text(repo_zip_name, 'cyan'), color_text(dest, 'green')
        ))
        self._latest_repo_zip_name = repo_zip_name

    def _update_index_html(self):
        """
        Rewrites the <a> tag in index.html to point at the current repo zip.
        """
        html_path = os.path.join(self.root_path, "index.html")
        zip_name = getattr(self, "_latest_repo_zip_name", None)
        if not zip_name or not os.path.exists(html_path):
            return

        with open(html_path, "r", encoding="utf-8") as f:
            content = f.read()

        import re
        new_link = '<a href="{0}">{0}</a>'.format(zip_name)
        updated = re.sub(
            r'<a href="[^"]*\.zip">[^<]*</a>',
            new_link,
            content,
        )

        if updated != content:
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(updated)
            print("Updated index.html link to {}".format(color_text(zip_name, 'cyan')))
        else:
            print("index.html already up to date.")

    def _generate_md5_file(self, addons_xml_path, md5_path):
        """
        Generates a new addons.xml.md5 file.
        """
        try:
            with open(addons_xml_path, "r", encoding="utf-8") as f:
                m = hashlib.md5(f.read().encode("utf-8")).hexdigest()
                self._save_file(m, file=md5_path)

            return True
        except Exception as e:
            print(
                "An error occurred updating {}!\n{}".format(
                    color_text(md5_path, 'yellow'), color_text(e, 'red')
                )
            )

    def _save_file(self, data, file):
        """
        Saves a file.
        """
        try:
            with open(file, "w") as f:
                f.write(data)
        except Exception as e:
            print(
                "An error occurred saving {}!\n{}".format(
                    color_text(file, 'yellow'), color_text(e, 'red')
                )
            )


if __name__ == "__main__":
    for release in [r for r in KODI_VERSIONS if os.path.exists(r)]:
        Generator(release)

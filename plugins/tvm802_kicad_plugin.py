import os
import csv

try:
    import pcbnew  # type: ignore
    import wx  # type: ignore
except ImportError:
    # These are only available inside KiCad's Python environment.
    pcbnew = None
    wx = None


def _detect_pos_format(header):
    """Detect whether a POS CSV is classic KiCad POS or positions.csv-style.

    Be tolerant of UTF-8 BOMs and minor header variations, especially on the
    first column (which often arrives as "\ufeffDesignator").
    """

    def norm(cell: str) -> str:
        s = cell.strip().lower()
        # Strip UTF-8 BOM if present (both decoded form and raw bytes)
        if s.startswith("\ufeff"):
            s = s[1:]
        if s.startswith("\xef\xbb\xbf"):
            s = s[3:]
        return s

    header_lower = [norm(h) for h in header]

    if (
        len(header_lower) >= 3
        and header_lower[0] in ("ref", "reference")
        and header_lower[1] == "val"
        and header_lower[2] == "package"
    ):
        return "kicad_pos"

    if (
        len(header_lower) >= 5
        and header_lower[0] in ("designator", "ref")
        and header_lower[1].startswith("mid x")
        and header_lower[2].startswith("mid y")
        and header_lower[3].startswith("rotation")
    ):
        return "positions"

    return "unknown"


def gen_components_list(pos_filename):
    components = []
    with open(pos_filename, newline="") as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        header = next(reader, None)
        if header is None:
            return components
        fmt = _detect_pos_format(header)
        for row in reader:
            if not row:
                continue
            if fmt == "kicad_pos":
                if len(row) < 3:
                    continue
                ref = row[0].strip()
                ref_upper = ref.upper()
                if ref_upper.startswith("FID"):
                    continue
                package = row[2].strip()
                val = row[1].strip()
                key = f"{package} {val}"
            elif fmt == "positions":
                if len(row) < 1:
                    continue
                ref = row[0].strip()
                ref_upper = ref.upper()
                if not ref or ref_upper.startswith("FID"):
                    continue
                key = ref
            else:
                if len(row) < 3:
                    continue
                ref = row[0].strip()
                ref_upper = ref.upper()
                if ref_upper.startswith("FID"):
                    continue
                # Unknown format: fall back to per-designator keys instead of
                # accidentally building keys from coordinates.
                key = ref
            components.append(key)
    return sorted(set(components))


def collect_fid_refs(pos_filename):
    refs = []
    with open(pos_filename, newline="") as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        header = next(reader, None)
        if header is None:
            return refs
        fmt = _detect_pos_format(header)
        for row in reader:
            if not row:
                continue
            if fmt == "kicad_pos":
                if len(row) < 1:
                    continue
                ref = row[0].strip()
            elif fmt == "positions":
                if len(row) < 1:
                    continue
                ref = row[0].strip()
            else:
                if len(row) < 1:
                    continue
                ref = row[0].strip()
            if not ref:
                continue
            ref_upper = ref.upper()
            if ref_upper.startswith("FID") and ref not in refs:
                refs.append(ref)
    return refs


def read_bom_ref_to_component_key(bom_filename):
    """Read mapping of reference designator -> component key from a BOM CSV.

    The BOM is expected to have a "Designator" column that may contain
    comma-separated designators, plus "Footprint" and "Value" columns.
    """
    ref_to_key = {}
    with open(bom_filename, newline="") as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        header = next(reader, None)
        if header is None:
            return ref_to_key

        def norm(cell: str) -> str:
            s = cell.strip()
            if s.startswith("\ufeff"):
                s = s[1:]
            if s.startswith("\xef\xbb\xbf"):
                s = s[3:]
            return s

        header_norm = [norm(h).lower() for h in header]

        try:
            idx_designator = header_norm.index("designator")
        except ValueError:
            idx_designator = None
        try:
            idx_footprint = header_norm.index("footprint")
        except ValueError:
            idx_footprint = None
        try:
            idx_value = header_norm.index("value")
        except ValueError:
            idx_value = None

        for row in reader:
            if not row:
                continue
            if idx_designator is None:
                continue
            designators = row[idx_designator].strip() if idx_designator < len(row) else ""
            if not designators:
                continue
            footprint = row[idx_footprint].strip() if idx_footprint is not None and idx_footprint < len(row) else ""
            value = row[idx_value].strip() if idx_value is not None and idx_value < len(row) else ""

            # Use the same idea as the classic POS format: footprint + value.
            component_key = f"{footprint} {value}".strip()
            # Split "C1, C2, C3" into individual refs.
            parts = [d.strip() for d in designators.replace(" ", "").split(",") if d.strip()]
            for ref in parts:
                ref_to_key[ref] = component_key
    return ref_to_key


def gen_components_list_from_bom_and_pos(pos_filename, bom_ref_to_key):
    """Generate component keys using both BOM and POS/positions files.

    Only designators that appear in the POS file are considered. Fiducials
    (FID*) are always skipped. When a BOM mapping exists for a designator it
    is used; otherwise we fall back to the normal POS-derived key.
    """
    components = set()
    with open(pos_filename, newline="") as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        header = next(reader, None)
        if header is None:
            return []
        fmt = _detect_pos_format(header)
        for row in reader:
            if not row:
                continue
            if fmt == "kicad_pos":
                if len(row) < 3:
                    continue
                ref = row[0].strip()
                ref_upper = ref.upper()
                if ref_upper.startswith("FID"):
                    continue
                package = row[2].strip()
                val = row[1].strip()
                default_key = f"{package} {val}"
            elif fmt == "positions":
                if len(row) < 1:
                    continue
                ref = row[0].strip()
                ref_upper = ref.upper()
                if not ref or ref_upper.startswith("FID"):
                    continue
                default_key = ref
            else:
                if len(row) < 1:
                    continue
                ref = row[0].strip()
                ref_upper = ref.upper()
                if not ref or ref_upper.startswith("FID"):
                    continue
                default_key = ref
            key = bom_ref_to_key.get(ref, default_key)
            if key:
                components.add(key)
    return sorted(components)


def _sanitize_explanation(text: str) -> str:
    for ch in ('"', '(', ')', '（', '）'):
        text = text.replace(ch, '')
    return text.strip()


def read_feeder_component_mappings(filename):
    """Read mapping of component key -> [Feeder, Nozzle, Speed, Height] from CSV."""
    cfeeders = {}
    firstline = True
    with open(filename, newline="") as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        for row in reader:
            if firstline:
                firstline = False
                continue
            if not row or len(row) < 5:
                continue
            key = row[0].strip()
            feeder = row[1].strip()
            nozzle = row[2].strip()
            speed = row[3].strip()
            height = row[4].strip()
            if key:
                cfeeders[key] = [feeder, nozzle, speed, height]
    return cfeeders


def gen_machine_data(pos_filename, cfeeders, output_file, mark1_ref=None, mark2_ref=None, bom_ref_to_key=None, skip_no_feeder=False):
    pcb_mark1_x = "0.00"
    pcb_mark1_y = "0.00"
    pcb_mark2_x = "0.00"
    pcb_mark2_y = "0.00"

    rows_total = 0
    rows_with_feeders = 0

    mark1_ref_upper = mark1_ref.upper() if mark1_ref else None
    mark2_ref_upper = mark2_ref.upper() if mark2_ref else None

    with open(pos_filename, newline="") as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        header = next(reader, None)
        if header is None:
            return
        fmt = _detect_pos_format(header)
        with open(output_file, "w", newline="") as outfile:
            # Tab-separated format matching TVM802 "Pick Place" files
            outfile.write("Designator\tNozzleNum\tStackNum\tMid X\tMid Y\tRotation\tHeight\tSpeed\tVision\tCheck\tExplanation\r\n")
            # Empty row after header (10 tabs for 11 columns)
            outfile.write("\t\t\t\t\t\t\t\t\t\t\r\n")

            for row in reader:
                if not row:
                    continue

                if fmt == "kicad_pos":
                    if len(row) < 6:
                        continue
                    ref = row[0].strip()
                    value = row[1].strip()
                    package = row[2].strip()
                    posx = row[3].strip()
                    posy = row[4].strip()
                    rot = row[5].strip()
                elif fmt == "positions":
                    if len(row) < 5:
                        continue
                    ref = row[0].strip()
                    posx = row[1].strip()
                    posy = row[2].strip()
                    rot = row[3].strip()
                    value = ""
                    package = ""
                else:
                    if len(row) < 6:
                        continue
                    ref = row[0].strip()
                    value = row[1].strip()
                    package = row[2].strip()
                    posx = row[3].strip()
                    posy = row[4].strip()
                    rot = row[5].strip()

                if not ref:
                    continue
                ref_upper = ref.upper()
                if mark1_ref_upper and ref_upper == mark1_ref_upper:
                    pcb_mark1_x = posx
                    pcb_mark1_y = posy
                    continue
                if mark2_ref_upper and ref_upper == mark2_ref_upper:
                    pcb_mark2_x = posx
                    pcb_mark2_y = posy
                    continue
                if not mark1_ref_upper and ref_upper in ("FID01", "FID1"):
                    pcb_mark1_x = posx
                    pcb_mark1_y = posy
                    continue
                if not mark2_ref_upper and ref_upper in ("FID02", "FID2"):
                    pcb_mark2_x = posx
                    pcb_mark2_y = posy
                    continue
                if ref_upper.startswith("FID"):
                    continue

                # Determine the component key used for feeder lookups. Prefer
                # BOM-derived keys when available so that BOM/grouping matches
                # the feeders template.
                if bom_ref_to_key is not None:
                    component_key = bom_ref_to_key.get(ref)
                else:
                    component_key = None

                if not component_key:
                    if fmt == "positions":
                        component_key = ref
                    else:
                        component_key = f"{package} {value}"

                feeder_params = cfeeders.get(component_key, ["", "", "", ""])
                feeder = feeder_params[0]
                nozzle = feeder_params[1]
                speed = feeder_params[2]
                height = feeder_params[3]

                if skip_no_feeder and not feeder:
                    continue

                rows_total += 1
                if feeder or nozzle:
                    rows_with_feeders += 1

                # Tab-separated row: Designator, NozzleNum, StackNum, Mid X, Mid Y, Rotation, Height, Speed, Vision, Check, Explanation
                # Use integer 0 for height if empty, matching working files
                h = height if height else "0"
                s = speed if speed else "100"
                n = nozzle if nozzle else "1"
                outfile.write(f"{ref}\t{n}\t{feeder}\t{posx}\t{posy}\t{rot}\t{h}\t{s}\tAccurate\tVision\t{_sanitize_explanation(component_key)}\r\n")

            # Trailing newline to match working example format
            outfile.write("\n")

    return rows_total, rows_with_feeders


class TVM802ActionPlugin(pcbnew.ActionPlugin if pcbnew is not None else object):
    """KiCad pcbnew action plugin to generate TVM802 machine data."""

    def defaults(self):  # type: ignore[override]
        self.name = "Export TVM802 Machine Data"
        self.category = "Fabrication"
        self.description = "Generate TVM802 pick-and-place CSV for TVM802 from KiCad POS + feeders CSV"
        self.show_toolbar_button = True
        self.icon_file_name = ""  # No custom icon by default

    def Run(self):  # type: ignore[override]
        if pcbnew is None or wx is None:
            return

        board = pcbnew.GetBoard()
        board_path = board.GetFileName() if board is not None else ""
        project_dir = os.path.dirname(board_path) if board_path else os.getcwd()

        # Prefer production/positions.csv if it exists
        default_pos = os.path.join(project_dir, "production", "positions.csv")
        default_dir = os.path.dirname(default_pos) if os.path.exists(default_pos) else project_dir

        pos_path = self._ask_open_file(
            "Select KiCad position CSV (POS)",
            default_dir,
            "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
        )
        if not pos_path:
            return

        # Inspect the POS file format so we can decide how strongly to require
        # a BOM (positions.csv-style inputs benefit from a BOM for grouping).
        pos_fmt = None
        try:
            with open(pos_path, newline="") as csvfile:
                reader = csv.reader(csvfile, delimiter=",")
                header = next(reader, None)
                if header is not None:
                    pos_fmt = _detect_pos_format(header)
        except Exception:
            pos_fmt = None

        bom_ref_to_key = None
        # For positions.csv-style inputs, require a BOM; for classic KiCad POS
        # it's optional (cancel to skip).
        if pos_fmt == "positions":
            bom_path = self._ask_open_file(
                "Select BOM CSV (required for positions.csv input)",
                default_dir,
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            )
            if not bom_path:
                wx.MessageBox(
                    "BOM CSV is required when using 'positions.csv' input.",
                    "TVM802 Export Error",
                    wx.OK | wx.ICON_ERROR,
                )
                return
            try:
                bom_ref_to_key = read_bom_ref_to_component_key(bom_path)
            except Exception as exc:  # pragma: no cover - runtime error path
                wx.MessageBox(
                    f"Failed to read BOM CSV:\n{exc}",
                    "TVM802 Export Error",
                    wx.OK | wx.ICON_ERROR,
                )
                return
        else:
            bom_path = self._ask_open_file(
                "Select BOM CSV (optional, cancel to skip)",
                default_dir,
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            )
            if bom_path:
                try:
                    bom_ref_to_key = read_bom_ref_to_component_key(bom_path)
                except Exception as exc:  # pragma: no cover - runtime error path
                    wx.MessageBox(
                        f"Failed to read BOM CSV (continuing without it):\n{exc}",
                        "TVM802 Export Warning",
                        wx.OK | wx.ICON_WARNING,
                    )
                    bom_ref_to_key = None

        # Let the user pick which fiducials to use as Mark1/Mark2
        fid_refs = collect_fid_refs(pos_path)
        mark1_ref = None
        mark2_ref = None
        if fid_refs:
            dlg = wx.SingleChoiceDialog(
                None,
                "Select fiducial for Mark 1 (lower-left)",
                "TVM802 Fiducial Selection",
                fid_refs,
            )
            try:
                if dlg.ShowModal() == wx.ID_OK:
                    mark1_ref = dlg.GetStringSelection()
            finally:
                dlg.Destroy()

            dlg2 = wx.SingleChoiceDialog(
                None,
                "Select fiducial for Mark 2 (second mark, e.g. upper-left)",
                "TVM802 Fiducial Selection",
                fid_refs,
            )
            try:
                if dlg2.ShowModal() == wx.ID_OK:
                    mark2_ref = dlg2.GetStringSelection()
            finally:
                dlg2.Destroy()

        skip_no_feeder = False
        opt_dlg = wx.MessageDialog(
            None,
            "Skip components without assigned feeder?",
            "TVM802 Options",
            wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION,
        )
        try:
            if opt_dlg.ShowModal() == wx.ID_YES:
                skip_no_feeder = True
        finally:
            opt_dlg.Destroy()

        feeders_path = self._ask_open_file(
            "Select feeders CSV configuration", 
            default_dir,
            "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
        )
        if not feeders_path:
            return

        output_default = os.path.join(default_dir, "tvm802-machine.csv")
        output_path = self._ask_save_file(
            "Save TVM802 machine data as",
            default_dir,
            output_default,
            "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
        )
        if not output_path:
            return

        try:
            cfeeders = read_feeder_component_mappings(feeders_path)
            rows_total, rows_with_feeders = gen_machine_data(
                pos_path,
                cfeeders,
                output_path,
                mark1_ref=mark1_ref,
                mark2_ref=mark2_ref,
                bom_ref_to_key=bom_ref_to_key,
                skip_no_feeder=skip_no_feeder,
            )
        except Exception as exc:  # pragma: no cover - runtime error path
            wx.MessageBox(
                f"Error during TVM802 export:\n{exc}",
                "TVM802 Export Error",
                wx.OK | wx.ICON_ERROR,
            )
            return

        msg = f"TVM802 machine data written to:\n{output_path}\n\n"
        msg += f"Placements exported: {rows_total}\n"
        msg += f"With feeders/nozzles: {rows_with_feeders}"
        if rows_with_feeders == 0 and rows_total > 0:
            msg += "\n\nWARNING: No feeders matched! Check that your feeders CSV\nuses the same component keys as the BOM."
        wx.MessageBox(
            msg,
            "TVM802 Export",
            wx.OK | wx.ICON_INFORMATION,
        )

    def _ask_open_file(self, message, default_dir, wildcard):
        dlg = wx.FileDialog(
            None,
            message,
            defaultDir=default_dir,
            wildcard=wildcard,
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        )
        try:
            if dlg.ShowModal() == wx.ID_CANCEL:
                return None
            return dlg.GetPath()
        finally:
            dlg.Destroy()

    def _ask_save_file(self, message, default_dir, default_path, wildcard):
        default_fname = os.path.basename(default_path)
        dlg = wx.FileDialog(
            None,
            message,
            defaultDir=default_dir,
            defaultFile=default_fname,
            wildcard=wildcard,
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
        )
        try:
            if dlg.ShowModal() == wx.ID_CANCEL:
                return None
            return dlg.GetPath()
        finally:
            dlg.Destroy()


class TVM802FeedersTemplatePlugin(TVM802ActionPlugin):
    def defaults(self):
        self.name = "Generate TVM802 Feeders Template"
        self.category = "Fabrication"
        self.description = "Generate unconfigured feeders CSV for TVM802 from KiCad POS CSV"
        self.show_toolbar_button = True
        self.icon_file_name = ""

    def Run(self):
        if pcbnew is None or wx is None:
            return

        board = pcbnew.GetBoard()
        board_path = board.GetFileName() if board is not None else ""
        default_dir = os.path.dirname(board_path) if board_path else os.getcwd()

        pos_path = self._ask_open_file(
            "Select KiCad position CSV (POS) for feeders template",
            default_dir,
            "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
        )
        if not pos_path:
            return

        # Inspect POS format to decide how to handle BOM input.
        pos_fmt = None
        try:
            with open(pos_path, newline="") as csvfile:
                reader = csv.reader(csvfile, delimiter=",")
                header = next(reader, None)
                if header is not None:
                    pos_fmt = _detect_pos_format(header)
        except Exception:
            pos_fmt = None

        bom_ref_to_key = None
        if pos_fmt == "positions":
            bom_path = self._ask_open_file(
                "Select BOM CSV (required for positions.csv input)",
                default_dir,
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            )
            if not bom_path:
                wx.MessageBox(
                    "BOM CSV is required when using 'positions.csv' input.",
                    "TVM802 Feeders Template Error",
                    wx.OK | wx.ICON_ERROR,
                )
                return
            try:
                bom_ref_to_key = read_bom_ref_to_component_key(bom_path)
            except Exception as exc:
                wx.MessageBox(
                    f"Failed to read BOM CSV:\n{exc}",
                    "TVM802 Feeders Template Error",
                    wx.OK | wx.ICON_ERROR,
                )
                return
        else:
            bom_path = self._ask_open_file(
                "Select BOM CSV (optional, cancel to skip)",
                default_dir,
                "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            )
            if bom_path:
                try:
                    bom_ref_to_key = read_bom_ref_to_component_key(bom_path)
                except Exception as exc:
                    wx.MessageBox(
                        f"Failed to read BOM CSV (continuing without it):\n{exc}",
                        "TVM802 Feeders Template Warning",
                        wx.OK | wx.ICON_WARNING,
                    )
                    bom_ref_to_key = None

        output_default = os.path.join(default_dir, "feeders-unconfigged.csv")
        output_path = self._ask_save_file(
            "Save TVM802 feeders template as",
            default_dir,
            output_default,
            "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
        )
        if not output_path:
            return

        try:
            if bom_ref_to_key is not None:
                components = gen_components_list_from_bom_and_pos(pos_path, bom_ref_to_key)
            else:
                components = gen_components_list(pos_path)
            with open(output_path, "w", newline="") as result_file:
                writer = csv.writer(result_file, delimiter=",", lineterminator="\r\n")
                writer.writerow(["Component", "Feeder", "Nozzle", "Speed", "Height"])
                for comp in components:
                    writer.writerow([comp, "", "1/2", "100", "0.5"])
        except Exception as exc:
            wx.MessageBox(
                f"Error during TVM802 feeders template generation:\n{exc}",
                "TVM802 Feeders Template Error",
                wx.OK | wx.ICON_ERROR,
            )
            return

        wx.MessageBox(
            f"TVM802 feeders template written to:\n{output_path}",
            "TVM802 Feeders Template",
            wx.OK | wx.ICON_INFORMATION,
        )


class TVM802ToolsPlugin(pcbnew.ActionPlugin if pcbnew is not None else object):
    def defaults(self):  # type: ignore[override]
        self.name = "TVM802 Tools"
        self.category = "Fabrication"
        self.description = "Generate TVM802 feeders template and/or machine data"
        self.show_toolbar_button = True
        self.icon_file_name = ""

    def Run(self):  # type: ignore[override]
        if pcbnew is None or wx is None:
            return

        choices = [
            "Generate TVM802 Feeders Template",
            "Export TVM802 Machine Data",
        ]

        dlg = wx.MultiChoiceDialog(
            None,
            "Select TVM802 actions to perform",
            "TVM802 Tools",
            choices,
        )
        try:
            if dlg.ShowModal() != wx.ID_OK:
                return
            selections = dlg.GetSelections()
        finally:
            dlg.Destroy()

        if 0 in selections:
            TVM802FeedersTemplatePlugin().Run()
        if 1 in selections:
            TVM802ActionPlugin().Run()


if pcbnew is not None:
    TVM802ToolsPlugin().register()

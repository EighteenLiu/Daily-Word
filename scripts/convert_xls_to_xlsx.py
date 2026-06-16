from __future__ import annotations

import argparse
import shutil
import sys
import time
import zipfile
from contextlib import contextmanager
from pathlib import Path


XLSX_FILE_FORMAT = 51
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
PICTURE_COMPRESSION_VALUE = "AutomaticPictureCompressionDefault"
OFFICE_VERSIONS = ("12.0", "14.0", "15.0", "16.0")
RPC_E_CALL_REJECTED = -2147418111
RPC_S_SERVER_UNAVAILABLE = -2147023174
EXCEL_COM_RETRY_ATTEMPTS = 8
EXCEL_COM_RETRY_DELAY_SECONDS = 1.5
EXCEL_CONVERT_TEMP_DIR = Path(r"D:\temp_excel_convert")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Use Microsoft Excel to convert .xls files to .xlsx while preserving "
            "formatting, formulas, merged cells, drawings, and pictures as much as Excel allows."
        )
    )
    parser.add_argument("input", type=Path, help="A .xls file or a directory containing .xls files.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help=(
            "Output .xlsx file for single-file input, or output directory for directory input. "
            "Defaults to the source file's directory."
        ),
    )
    parser.add_argument("-r", "--recursive", action="store_true", help="Search directories recursively.")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing .xlsx files.")
    parser.add_argument("--visible", action="store_true", help="Show the Excel window during conversion.")
    parser.add_argument(
        "--keep-registry-setting",
        action="store_true",
        help=(
            "Keep Excel's image-compression registry setting disabled after conversion. "
            "By default, the previous registry values are restored."
        ),
    )
    parser.add_argument(
        "--no-registry-protection",
        action="store_true",
        help=(
            "Do not temporarily disable Excel's default picture compression in the registry. "
            "Use only if your environment blocks registry changes."
        ),
    )
    parser.add_argument(
        "--verify-media",
        action="store_true",
        help="After saving, print the number and total size of media files embedded in each .xlsx.",
    )
    return parser.parse_args()


def discover_sources(input_path: Path, recursive: bool) -> list[Path]:
    input_path = input_path.resolve()
    if input_path.is_file():
        if input_path.suffix.lower() != ".xls":
            raise ValueError(f"Input file is not .xls: {input_path}")
        return [input_path]

    if not input_path.is_dir():
        raise FileNotFoundError(f"Input path does not exist: {input_path}")

    pattern = "**/*.xls" if recursive else "*.xls"
    return sorted(
        path.resolve()
        for path in input_path.glob(pattern)
        if not path.name.startswith("~$") and path.suffix.lower() == ".xls"
    )


def output_path_for(source: Path, input_root: Path, output: Path | None) -> Path:
    if source.is_file() and input_root.is_file():
        if output is None:
            return source.with_suffix(".xlsx")
        if output.suffix.lower() == ".xlsx":
            return output.resolve()
        return (output.resolve() / source.with_suffix(".xlsx").name)

    if output is None:
        return source.with_suffix(".xlsx")

    output_root = output.resolve()
    try:
        relative = source.relative_to(input_root.resolve())
    except ValueError:
        relative = Path(source.name)
    return (output_root / relative).with_suffix(".xlsx")


@contextmanager
def temporary_excel_picture_compression_setting(enabled: bool, keep_setting: bool):
    if not enabled or sys.platform != "win32":
        yield
        return

    import winreg

    saved_values: list[tuple[str, int | None, int | None]] = []
    for version in OFFICE_VERSIONS:
        key_path = rf"Software\Microsoft\Office\{version}\Excel\Options"
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        try:
            try:
                old_value, old_type = winreg.QueryValueEx(key, PICTURE_COMPRESSION_VALUE)
            except FileNotFoundError:
                old_value, old_type = None, None
            saved_values.append((key_path, old_value, old_type))
            winreg.SetValueEx(key, PICTURE_COMPRESSION_VALUE, 0, winreg.REG_DWORD, 0)
        finally:
            winreg.CloseKey(key)

    try:
        yield
    finally:
        if keep_setting:
            return
        for key_path, old_value, old_type in saved_values:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
            try:
                if old_value is None:
                    try:
                        winreg.DeleteValue(key, PICTURE_COMPRESSION_VALUE)
                    except FileNotFoundError:
                        pass
                else:
                    winreg.SetValueEx(key, PICTURE_COMPRESSION_VALUE, 0, old_type, old_value)
            finally:
                winreg.CloseKey(key)


def set_workbook_no_picture_compression(workbook) -> None:
    # Newer Office builds expose this workbook-level flag through COM.
    for property_name in ("DoNotCompressPictures", "DoNotCompressImages"):
        try:
            setattr(workbook, property_name, True)
        except Exception:
            continue


def embedded_media_summary(xlsx_path: Path) -> tuple[int, int]:
    with zipfile.ZipFile(xlsx_path) as archive:
        media_files = [info for info in archive.infolist() if info.filename.startswith("xl/media/")]
    return len(media_files), sum(info.file_size for info in media_files)


def safe_remove(path: Path) -> None:
    path = Path(path)
    if not path.exists():
        return
    try:
        path.unlink()
    except PermissionError as exc:
        raise RuntimeError(f"目标文件被占用，无法删除：{path}") from exc


def is_excel_call_rejected(exc: Exception) -> bool:
    return bool(getattr(exc, "args", ())) and getattr(exc, "args", ())[0] == RPC_E_CALL_REJECTED


def is_excel_disconnect(exc: Exception) -> bool:
    return bool(getattr(exc, "args", ())) and getattr(exc, "args", ())[0] in {
        RPC_E_CALL_REJECTED,
        RPC_S_SERVER_UNAVAILABLE,
    }


def close_workbook_quietly(workbook) -> None:
    if workbook is None:
        return
    try:
        workbook.Close(SaveChanges=False)
    except Exception:
        # SaveAs 失败后 Excel COM 对象可能断开；关闭失败不应覆盖真正的保存错误。
        pass


def quit_excel_quietly(excel) -> None:
    if excel is None:
        return
    try:
        excel.Quit()
    except Exception:
        # Excel.Application 已断开时，Quit 本身也可能不可用。
        pass


def call_excel_with_retry(action: str, func):
    last_exc: Exception | None = None
    for attempt in range(1, EXCEL_COM_RETRY_ATTEMPTS + 1):
        try:
            return func()
        except Exception as exc:
            if not is_excel_call_rejected(exc):
                raise
            last_exc = exc
            print(
                f"[wait] Excel 正忙，正在重试 {action} "
                f"({attempt}/{EXCEL_COM_RETRY_ATTEMPTS})...",
                file=sys.stderr,
            )
            try:
                import pythoncom

                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(EXCEL_COM_RETRY_DELAY_SECONDS)
    raise RuntimeError(
        "Excel 一直拒绝接收自动化调用。请关闭所有 Excel 窗口、弹窗和保护视图后重试；"
        "也可以手动打开该 .xls，另存为 .xlsx 后再选择 .xlsx 台账。"
    ) from last_exc


def lock_file_for_xlsx(xlsx_path: Path) -> Path:
    return xlsx_path.parent / f"~${xlsx_path.name}"


def temporary_xlsx_path_for(xlsx_path: Path) -> Path:
    EXCEL_CONVERT_TEMP_DIR.mkdir(parents=True, exist_ok=True)
    return EXCEL_CONVERT_TEMP_DIR / f"{xlsx_path.stem}_{int(time.time())}.xlsx"


def prepare_xlsx_destination(xlsx_path: Path) -> Path:
    xlsx_path = Path(xlsx_path)
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    safe_remove(xlsx_path)
    safe_remove(lock_file_for_xlsx(xlsx_path))
    tmp_xlsx_path = temporary_xlsx_path_for(xlsx_path)
    safe_remove(tmp_xlsx_path)
    safe_remove(lock_file_for_xlsx(tmp_xlsx_path))
    return tmp_xlsx_path


def replace_xlsx_destination(tmp_xlsx_path: Path, xlsx_path: Path) -> None:
    safe_remove(xlsx_path)
    safe_remove(lock_file_for_xlsx(xlsx_path))
    shutil.move(str(tmp_xlsx_path), str(xlsx_path))


def convert_with_excel(
    jobs: list[tuple[Path, Path]],
    overwrite: bool,
    visible: bool,
    verify_media: bool,
) -> int:
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        print(
            "pywin32 is required. Install it with: pip install pywin32",
            file=sys.stderr,
        )
        print(f"Import error: {exc}", file=sys.stderr)
        return len(jobs)

    pythoncom.CoInitialize()
    excel = None
    failures = 0
    try:
        try:
            excel = win32.DispatchEx("Excel.Application")
        except Exception as exc:
            print(
                "Cannot start Microsoft Excel through COM automation. "
                "Run this script in an interactive Windows user session with Excel installed.",
                file=sys.stderr,
            )
            print(f"Excel startup error: {exc}", file=sys.stderr)
            return len(jobs)

        excel.Visible = visible
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        for source, destination in jobs:
            if destination.exists() and not overwrite:
                print(f"[skip] exists: {destination}")
                continue

            workbook = None
            tmp_destination: Path | None = None
            try:
                tmp_destination = prepare_xlsx_destination(destination)
                workbook = call_excel_with_retry(
                    f"打开 {source.name}",
                    lambda: excel.Workbooks.Open(
                        Filename=str(source),
                        UpdateLinks=0,
                        ReadOnly=True,
                        AddToMru=False,
                        IgnoreReadOnlyRecommended=True,
                        CorruptLoad=1,
                    ),
                )
                workbook.CheckCompatibility = False
                set_workbook_no_picture_compression(workbook)
                call_excel_with_retry(
                    f"保存 {tmp_destination.name}",
                    lambda: workbook.SaveAs(
                        Filename=str(tmp_destination),
                        FileFormat=XLSX_FILE_FORMAT,
                        CreateBackup=False,
                        ConflictResolution=2,
                        Local=True,
                    ),
                )
                close_workbook_quietly(workbook)
                workbook = None
                replace_xlsx_destination(tmp_destination, destination)
                tmp_destination = None
                if verify_media:
                    count, total_size = embedded_media_summary(destination)
                    print(f"[ok] {source} -> {destination} | media: {count} files, {total_size} bytes")
                else:
                    print(f"[ok] {source} -> {destination}")
            except Exception as exc:
                failures += 1
                print(f"[fail] {source}: {exc}", file=sys.stderr)
            finally:
                if workbook is not None:
                    close_workbook_quietly(workbook)
                if tmp_destination is not None:
                    try:
                        safe_remove(tmp_destination)
                    except Exception:
                        pass
    finally:
        if excel is not None:
            quit_excel_quietly(excel)
        pythoncom.CoUninitialize()

    return failures


def main() -> int:
    args = parse_args()
    input_path = args.input.resolve()
    sources = discover_sources(input_path, args.recursive)
    if not sources:
        print(f"No .xls files found: {input_path}")
        return 0

    jobs = [(source, output_path_for(source, input_path, args.output)) for source in sources]
    with temporary_excel_picture_compression_setting(
        enabled=not args.no_registry_protection,
        keep_setting=args.keep_registry_setting,
    ):
        failures = convert_with_excel(
            jobs=jobs,
            overwrite=args.overwrite,
            visible=args.visible,
            verify_media=args.verify_media,
        )

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())

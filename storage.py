"""SQLite persistence and project-file storage helpers."""

import json
import os
import shutil
import sqlite3
import tempfile
import time
import zipfile
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("AILYN_DATA_DIR", APP_DIR / "data")).expanduser().resolve()
DB_FILE = str(DATA_DIR / "ailyn_state.db")
BACKUP_DIR = str(DATA_DIR / "backups")
PHOTO_DIR = DATA_DIR / "scanner_photos"
STATE_TABLE = "app_state"
SCHEMA_VERSION = 1
MAX_BACKUP_MEMBERS = 10_000
MAX_BACKUP_UNCOMPRESSED_BYTES = 250 * 1024 * 1024
LEGACY_DB_FILE = APP_DIR / "ailyn_state.db"
LEGACY_DATA_NAMES = (
    "ailyn_project_ledger.xlsx",
    "materials_ledger.xlsx",
    "labor_ledger.xlsx",
    "backups",
    "archive",
    "scanner_photos",
)


def _migrate_legacy_database():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not os.path.exists(DB_FILE) and LEGACY_DB_FILE.exists() and LEGACY_DB_FILE != Path(DB_FILE):
        shutil.copy2(LEGACY_DB_FILE, DB_FILE)
    for name in LEGACY_DATA_NAMES:
        source = APP_DIR / name
        destination = DATA_DIR / name
        if source.exists() and not destination.exists() and source.resolve() != destination.resolve():
            if source.is_dir():
                shutil.copytree(source, destination)
            else:
                shutil.copy2(source, destination)


def _connect(path=None):
    path = path or DB_FILE
    connection = sqlite3.connect(path)
    connection.execute(
        f"CREATE TABLE IF NOT EXISTS {STATE_TABLE} "
        "(key TEXT PRIMARY KEY, value TEXT NOT NULL)"
    )
    connection.execute("CREATE TABLE IF NOT EXISTS schema_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
    current_version = connection.execute(
        "SELECT value FROM schema_meta WHERE key = 'schema_version'"
    ).fetchone()
    if current_version is None:
        connection.execute(
            "INSERT INTO schema_meta(key, value) VALUES ('schema_version', ?)",
            (str(SCHEMA_VERSION),),
        )
    elif int(current_version[0]) > SCHEMA_VERSION:
        connection.close()
        raise sqlite3.DatabaseError("The database was created by a newer app version.")
    connection.commit()
    return connection


def load_state():
    """Load persisted JSON values, returning an empty state when absent."""
    _migrate_legacy_database()
    if not os.path.exists(DB_FILE):
        return {}
    try:
        with _connect() as connection:
            rows = connection.execute(f"SELECT key, value FROM {STATE_TABLE}").fetchall()
        return {key: json.loads(value) for key, value in rows}
    except (OSError, sqlite3.Error, json.JSONDecodeError):
        return {}


def save_state(state):
    """Atomically persist JSON-serializable application state."""
    values = {}
    for key, value in dict(state).items():
        try:
            values[key] = json.dumps(value, ensure_ascii=True)
        except (TypeError, ValueError):
            continue
    temporary_path = None
    database_directory = os.path.dirname(os.path.abspath(DB_FILE))
    os.makedirs(database_directory, exist_ok=True)
    try:
        with tempfile.NamedTemporaryFile(delete=False, dir=database_directory, suffix=".db") as temporary:
            temporary_path = temporary.name
        with _connect(temporary_path) as connection:
            connection.executemany(
                f"INSERT INTO {STATE_TABLE}(key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                values.items(),
            )
            connection.commit()
        os.replace(temporary_path, DB_FILE)
        os.chmod(DB_FILE, 0o600)
    finally:
        if temporary_path and os.path.exists(temporary_path):
            os.remove(temporary_path)
    return True


def create_backup():
    """Create a complete, immutable ZIP backup of all managed project data."""
    if not os.path.exists(DB_FILE):
        save_state({})
    os.makedirs(BACKUP_DIR, exist_ok=True)
    destination = os.path.join(BACKUP_DIR, f"ailyn_backup_{time.time_ns()}.zip")
    data_root = Path(DATA_DIR)
    backup_root = Path(BACKUP_DIR).resolve()
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in data_root.rglob("*"):
            if path.is_file() and backup_root not in path.resolve().parents:
                archive.write(path, path.relative_to(data_root).as_posix())
    os.chmod(destination, 0o600)
    return destination


def restore_backup(path):
    """Validate a backup and restore its managed files without deleting current data."""
    source_path = Path(path)
    temporary_directory = None
    try:
        if zipfile.is_zipfile(source_path):
            temporary_directory = tempfile.mkdtemp(prefix="ailyn-restore-")
            with zipfile.ZipFile(source_path) as archive:
                members = archive.namelist()
                if any(member.startswith("/") or ".." in Path(member).parts for member in members):
                    raise ValueError("The backup contains unsafe file paths.")
                if len(members) > MAX_BACKUP_MEMBERS or sum(info.file_size for info in archive.infolist()) > MAX_BACKUP_UNCOMPRESSED_BYTES:
                    raise ValueError("The backup is too large to restore safely.")
                archive.extractall(temporary_directory)
            database_path = os.path.join(temporary_directory, "ailyn_state.db")
        else:
            database_path = str(source_path)
        if not os.path.isfile(database_path):
            raise ValueError("The backup does not contain a database.")
        with sqlite3.connect(database_path) as validation_connection:
            has_state_table = validation_connection.execute(
                "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
                (STATE_TABLE,),
            ).fetchone()
        if not has_state_table:
            raise ValueError("The backup does not contain application state.")
        with _connect(database_path) as connection:
            rows = connection.execute(f"SELECT key, value FROM {STATE_TABLE}").fetchall()
        state = {key: json.loads(value) for key, value in rows}
    except (OSError, sqlite3.Error, json.JSONDecodeError) as error:
        raise ValueError("The backup file is invalid or unreadable.") from error
    finally:
        if temporary_directory:
            shutil.rmtree(temporary_directory, ignore_errors=True)
    if zipfile.is_zipfile(source_path):
        with zipfile.ZipFile(source_path) as archive:
            if len(archive.infolist()) > MAX_BACKUP_MEMBERS or sum(info.file_size for info in archive.infolist()) > MAX_BACKUP_UNCOMPRESSED_BYTES:
                raise ValueError("The backup is too large to restore safely.")
            for member in archive.namelist():
                target = (DATA_DIR / member).resolve()
                if target.parent != DATA_DIR.resolve() and DATA_DIR.resolve() not in target.parents:
                    raise ValueError("The backup contains unsafe file paths.")
                if member.endswith("/"):
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source, open(target, "wb") as destination:
                    shutil.copyfileobj(source, destination)
    return state


def save_scanner_photo(photo_bytes, mime_type, photo_id):
    """Store a scanner image and return its app-relative path."""
    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    extension = ".png" if mime_type == "image/png" else ".jpg"
    path = PHOTO_DIR / f"{photo_id}{extension}"
    path.write_bytes(photo_bytes)
    return os.path.relpath(path, DATA_DIR)


def delete_scanner_photo(relative_path):
    """Delete only a photo inside the managed scanner directory."""
    if not relative_path:
        return
    path = (DATA_DIR / relative_path).resolve()
    if not path.is_file():
        path = (APP_DIR / relative_path).resolve()
    if PHOTO_DIR.resolve() in path.parents and path.is_file():
        path.unlink()


def history_count():
    """Return the number of persisted state keys for compatibility."""
    with _connect() as connection:
        return connection.execute(f"SELECT COUNT(*) FROM {STATE_TABLE}").fetchone()[0]
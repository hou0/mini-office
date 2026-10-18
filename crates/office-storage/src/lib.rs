use office_core::{sha256_hex, FileInfo, OfficeError, OfficeKind};
use rusqlite::{params, Connection, OptionalExtension};
use serde::{Deserialize, Serialize};
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SaveOutcome {
    pub path: String,
    pub backup_path: Option<String>,
    pub bytes_written: u64,
    pub sha256: String,
}

fn backup_path_for(path: &Path) -> PathBuf {
    let stamp = chrono::Local::now().format("%Y%m%d-%H%M%S-%3f");
    let suffix = uuid::Uuid::new_v4().simple().to_string();
    let file_name = path
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("document");
    path.with_file_name(format!(
        "{file_name}.mini-office-backup-{stamp}-{}",
        &suffix[..6]
    ))
}

fn sync_parent(path: &Path) -> Result<(), OfficeError> {
    #[cfg(unix)]
    {
        if let Some(parent) = path.parent() {
            if let Ok(directory) = fs::File::open(parent) {
                let _ = directory.sync_all();
            }
        }
    }
    Ok(())
}

/// Write to a sibling temporary file, fsync, then atomically replace the target.
/// The previous file is copied to a timestamped backup before replacement.
pub fn atomic_write(
    path: impl AsRef<Path>,
    bytes: &[u8],
    create_backup: bool,
) -> Result<SaveOutcome, OfficeError> {
    let path = path.as_ref();
    let parent = path
        .parent()
        .ok_or_else(|| OfficeError::Storage("保存路径缺少父目录。".to_string()))?;
    if !parent.exists() {
        fs::create_dir_all(parent)?;
    }
    let mut temp = tempfile::Builder::new()
        .prefix(".mini-office-tmp-")
        .suffix(".part")
        .tempfile_in(parent)
        .map_err(|error| OfficeError::Storage(format!("无法创建临时文件：{error}")))?;
    temp.write_all(bytes)?;
    temp.as_file_mut().sync_all()?;

    let backup = if create_backup && path.exists() {
        let backup = backup_path_for(path);
        fs::copy(path, &backup)?;
        Some(backup)
    } else {
        None
    };

    let file = match temp.persist(path) {
        Ok(file) => file,
        Err(error) => {
            // Windows cannot rename over an existing target. The backup above
            // still protects the original file if this fallback is needed.
            let temp = error.file;
            if path.exists() {
                std::fs::remove_file(path)?;
            }
            temp.persist(path)
                .map_err(|error| OfficeError::Storage(format!("无法替换目标文件：{error}")))?
        }
    };
    file.sync_all()?;
    sync_parent(path)?;

    Ok(SaveOutcome {
        path: path.to_string_lossy().to_string(),
        backup_path: backup.map(|path| path.to_string_lossy().to_string()),
        bytes_written: bytes.len() as u64,
        sha256: sha256_hex(bytes),
    })
}

pub fn inspect_file(path: impl AsRef<Path>) -> Result<FileInfo, OfficeError> {
    FileInfo::from_path(path)
}

pub fn detect_kind(path: impl AsRef<Path>) -> Option<OfficeKind> {
    OfficeKind::from_extension(path)
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RecentFile {
    pub id: String,
    pub path: String,
    pub name: String,
    pub kind: String,
    pub size: u64,
    pub updated_at: i64,
    pub level: u8,
}

pub struct RecentStore {
    connection: Arc<Mutex<Connection>>,
}

impl RecentStore {
    pub fn open(app_data_dir: impl AsRef<Path>) -> Result<Self, OfficeError> {
        let directory = app_data_dir.as_ref();
        fs::create_dir_all(directory)?;
        let connection = Connection::open(directory.join("mini-office.sqlite3"))
            .map_err(|error| OfficeError::Storage(format!("无法打开数据库：{error}")))?;
        connection
            .execute_batch(
                "PRAGMA journal_mode = WAL;
                 CREATE TABLE IF NOT EXISTS recent_files (
                     id TEXT PRIMARY KEY,
                     path TEXT NOT NULL UNIQUE,
                     name TEXT NOT NULL,
                     kind TEXT NOT NULL,
                     size INTEGER NOT NULL,
                     updated_at INTEGER NOT NULL,
                     level INTEGER NOT NULL DEFAULT 3
                 );
                 CREATE INDEX IF NOT EXISTS idx_recent_updated ON recent_files(updated_at DESC);",
            )
            .map_err(|error| OfficeError::Storage(format!("无法初始化数据库：{error}")))?;
        Ok(Self {
            connection: Arc::new(Mutex::new(connection)),
        })
    }

    pub fn record(&self, file: &RecentFile) -> Result<(), OfficeError> {
        let connection = self
            .connection
            .lock()
            .map_err(|_| OfficeError::Storage("数据库繁忙。".to_string()))?;
        connection
            .execute(
                "INSERT INTO recent_files (id, path, name, kind, size, updated_at, level)
                 VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7)
                 ON CONFLICT(path) DO UPDATE SET
                     name = excluded.name,
                     kind = excluded.kind,
                     size = excluded.size,
                     updated_at = excluded.updated_at,
                     level = excluded.level",
                params![
                    file.id,
                    file.path,
                    file.name,
                    file.kind,
                    file.size as i64,
                    file.updated_at,
                    file.level as i64
                ],
            )
            .map_err(|error| OfficeError::Storage(format!("无法写入最近文件：{error}")))?;
        Ok(())
    }

    pub fn list(&self, limit: usize) -> Result<Vec<RecentFile>, OfficeError> {
        let connection = self
            .connection
            .lock()
            .map_err(|_| OfficeError::Storage("数据库繁忙。".to_string()))?;
        let mut statement = connection
            .prepare(
                "SELECT id, path, name, kind, size, updated_at, level
                 FROM recent_files ORDER BY updated_at DESC LIMIT ?1",
            )
            .map_err(|error| OfficeError::Storage(error.to_string()))?;
        let rows = statement
            .query_map(params![limit as i64], |row| {
                Ok(RecentFile {
                    id: row.get(0)?,
                    path: row.get(1)?,
                    name: row.get(2)?,
                    kind: row.get(3)?,
                    size: row.get::<_, i64>(4)? as u64,
                    updated_at: row.get(5)?,
                    level: row.get::<_, i64>(6)? as u8,
                })
            })
            .map_err(|error| OfficeError::Storage(error.to_string()))?;
        let mut result = Vec::new();
        for row in rows {
            result.push(row.map_err(|error| OfficeError::Storage(error.to_string()))?);
        }
        Ok(result)
    }

    pub fn find_by_path(&self, path: &str) -> Result<Option<RecentFile>, OfficeError> {
        let connection = self
            .connection
            .lock()
            .map_err(|_| OfficeError::Storage("数据库繁忙。".to_string()))?;
        connection
            .query_row(
                "SELECT id, path, name, kind, size, updated_at, level FROM recent_files WHERE path = ?1",
                params![path],
                |row| {
                    Ok(RecentFile {
                        id: row.get(0)?,
                        path: row.get(1)?,
                        name: row.get(2)?,
                        kind: row.get(3)?,
                        size: row.get::<_, i64>(4)? as u64,
                        updated_at: row.get(5)?,
                        level: row.get::<_, i64>(6)? as u8,
                    })
                },
            )
            .optional()
            .map_err(|error| OfficeError::Storage(error.to_string()))
    }

    pub fn remove(&self, id: &str) -> Result<(), OfficeError> {
        let connection = self
            .connection
            .lock()
            .map_err(|_| OfficeError::Storage("数据库繁忙。".to_string()))?;
        connection
            .execute("DELETE FROM recent_files WHERE id = ?1", params![id])
            .map_err(|error| OfficeError::Storage(error.to_string()))?;
        Ok(())
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RecoverySnapshot {
    pub id: String,
    pub display_name: String,
    pub kind: String,
    pub updated_at: i64,
    pub size: u64,
    pub file_name: String,
}

pub struct RecoveryStore {
    root: PathBuf,
}

impl RecoveryStore {
    pub fn open(app_data_dir: impl AsRef<Path>) -> Result<Self, OfficeError> {
        let root = app_data_dir.as_ref().join("recovery");
        fs::create_dir_all(&root)?;
        Ok(Self { root })
    }

    pub fn save(
        &self,
        display_name: &str,
        kind: &str,
        bytes: &[u8],
    ) -> Result<RecoverySnapshot, OfficeError> {
        let id = uuid::Uuid::new_v4().to_string();
        let file_name = format!("{id}.snapshot");
        atomic_write(self.root.join(&file_name), bytes, false)?;
        let snapshot = RecoverySnapshot {
            id,
            display_name: display_name.to_string(),
            kind: kind.to_string(),
            updated_at: chrono::Utc::now().timestamp_millis(),
            size: bytes.len() as u64,
            file_name,
        };
        let meta = serde_json::to_vec_pretty(&snapshot)
            .map_err(|error| OfficeError::Storage(error.to_string()))?;
        atomic_write(
            self.root.join(format!("{}.json", snapshot.id)),
            &meta,
            false,
        )?;
        Ok(snapshot)
    }

    pub fn list(&self) -> Result<Vec<RecoverySnapshot>, OfficeError> {
        let mut snapshots = Vec::new();
        for entry in fs::read_dir(&self.root)? {
            let entry = entry?;
            let path = entry.path();
            if path.extension().and_then(|value| value.to_str()) != Some("json") {
                continue;
            }
            if let Ok(data) = fs::read(&path) {
                if let Ok(snapshot) = serde_json::from_slice::<RecoverySnapshot>(&data) {
                    snapshots.push(snapshot);
                }
            }
        }
        snapshots.sort_by(|a, b| b.updated_at.cmp(&a.updated_at));
        Ok(snapshots)
    }

    pub fn load(&self, id: &str) -> Result<Vec<u8>, OfficeError> {
        let meta_path = self.root.join(format!("{id}.json"));
        let data = fs::read(&meta_path)?;
        let snapshot: RecoverySnapshot = serde_json::from_slice(&data)
            .map_err(|error| OfficeError::Storage(error.to_string()))?;
        Ok(fs::read(self.root.join(snapshot.file_name))?)
    }

    pub fn delete(&self, id: &str) -> Result<(), OfficeError> {
        let meta_path = self.root.join(format!("{id}.json"));
        if let Ok(data) = fs::read(&meta_path) {
            if let Ok(snapshot) = serde_json::from_slice::<RecoverySnapshot>(&data) {
                let _ = fs::remove_file(self.root.join(snapshot.file_name));
            }
        }
        let _ = fs::remove_file(meta_path);
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn atomic_write_creates_backup_and_content() {
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().join("document.docx");
        fs::write(&target, b"old").unwrap();
        let outcome = atomic_write(&target, b"new-content", true).unwrap();
        assert_eq!(fs::read(&target).unwrap(), b"new-content");
        assert_eq!(outcome.bytes_written, 11);
        assert!(outcome.backup_path.is_some());
        let backup = outcome.backup_path.unwrap();
        assert_eq!(fs::read(backup).unwrap(), b"old");
    }

    #[test]
    fn recent_store_round_trips() {
        let dir = tempfile::tempdir().unwrap();
        let store = RecentStore::open(dir.path()).unwrap();
        let file = RecentFile {
            id: "1".to_string(),
            path: "/tmp/a.docx".to_string(),
            name: "a.docx".to_string(),
            kind: "docx".to_string(),
            size: 12,
            updated_at: 1,
            level: 3,
        };
        store.record(&file).unwrap();
        let list = store.list(10).unwrap();
        assert_eq!(list.len(), 1);
        assert_eq!(list[0].name, "a.docx");
        assert!(store.find_by_path("/tmp/a.docx").unwrap().is_some());
        store.remove("1").unwrap();
        assert!(store.list(10).unwrap().is_empty());
    }

    #[test]
    fn recovery_store_round_trips() {
        let dir = tempfile::tempdir().unwrap();
        let store = RecoveryStore::open(dir.path()).unwrap();
        let snapshot = store.save("合同.docx", "docx", b"recovery").unwrap();
        assert_eq!(store.list().unwrap().len(), 1);
        assert_eq!(store.load(&snapshot.id).unwrap(), b"recovery");
        store.delete(&snapshot.id).unwrap();
        assert!(store.list().unwrap().is_empty());
    }
}

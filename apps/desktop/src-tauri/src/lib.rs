use office_core::{FileInfo, OfficeKind, PlatformInfo, MAX_INPUT_BYTES};
use office_storage::{
    atomic_write, RecentFile, RecentStore, RecoverySnapshot, RecoveryStore, SaveOutcome,
};
use serde::Serialize;
use std::path::Path;
use tauri::{Manager, State};

struct AppState {
    recent: RecentStore,
    recovery: RecoveryStore,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct DocumentInspectionDto {
    kind: OfficeKind,
    level: u8,
    main_part: String,
    entry_count: usize,
    warnings: Vec<office_core::CompatibilityWarning>,
    features: Vec<office_core::CapabilityFeature>,
}

fn allowed_output(path: &str) -> bool {
    let extension = Path::new(path)
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    matches!(
        extension.as_str(),
        "docx" | "xlsx" | "pptx" | "csv" | "pdf" | "zip" | "png" | "jpg" | "jpeg"
    )
}

#[tauri::command]
fn read_file(path: String) -> Result<Vec<u8>, String> {
    let metadata = std::fs::metadata(&path).map_err(|error| format!("无法读取文件：{error}"))?;
    if !metadata.is_file() {
        return Err("目标不是文件。".to_string());
    }
    if metadata.len() > MAX_INPUT_BYTES {
        return Err("文件超过 250 MB 的 V1 上限。".to_string());
    }
    std::fs::read(&path).map_err(|error| format!("无法读取文件：{error}"))
}

#[tauri::command]
fn write_file(path: String, bytes: Vec<u8>) -> Result<SaveOutcome, String> {
    if bytes.is_empty() {
        return Err("拒绝写入空文件。".to_string());
    }
    if !allowed_output(&path) {
        return Err("不允许写入该文件类型。".to_string());
    }
    if OfficeKind::from_extension(&path).is_some() {
        office_ooxml::validate_package(&bytes)
            .map_err(|error| format!("保存前容器校验失败：{error}"))?;
    }
    atomic_write(&path, &bytes, true).map_err(|error| error.to_string())
}

#[tauri::command]
fn list_fonts() -> Vec<office_font::FontInfo> {
    office_font::cached_fonts().to_vec()
}

#[tauri::command]
fn inspect_document(path: String) -> Result<DocumentInspectionDto, String> {
    let bytes = read_file(path.clone())?;
    let name = Path::new(&path)
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("document");
    let inspection = office_ooxml::analyze(&bytes, name).map_err(|error| error.to_string())?;
    Ok(DocumentInspectionDto {
        kind: inspection.kind,
        level: inspection.capability.level,
        main_part: inspection.main_part,
        entry_count: inspection.entry_count,
        warnings: inspection.capability.warnings,
        features: inspection.capability.features,
    })
}

#[tauri::command]
fn platform_info() -> PlatformInfo {
    office_platform::platform_info()
}

#[tauri::command]
fn app_version() -> String {
    office_platform::app_version().to_string()
}

#[tauri::command]
fn recent_list(state: State<'_, AppState>, limit: Option<usize>) -> Result<Vec<RecentFile>, String> {
    state
        .recent
        .list(limit.unwrap_or(30))
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn recent_record(state: State<'_, AppState>, path: String) -> Result<RecentFile, String> {
    let info: FileInfo = office_storage::inspect_file(&path).map_err(|error| error.to_string())?;
    let kind = OfficeKind::from_extension(&path)
        .map(|kind| match kind {
            OfficeKind::Docx => "docx",
            OfficeKind::Xlsx => "xlsx",
            OfficeKind::Pptx => "pptx",
        })
        .unwrap_or("bin")
        .to_string();
    let entry = RecentFile {
        id: uuid::Uuid::new_v4().to_string(),
        path: info.path,
        name: info.name,
        kind,
        size: info.size,
        updated_at: chrono::Utc::now().timestamp_millis(),
        level: 3,
    };
    state
        .recent
        .record(&entry)
        .map_err(|error| error.to_string())?;
    Ok(entry)
}

#[tauri::command]
fn recovery_save(
    state: State<'_, AppState>,
    display_name: String,
    kind: String,
    bytes: Vec<u8>,
) -> Result<RecoverySnapshot, String> {
    if bytes.is_empty() || bytes.len() as u64 > MAX_INPUT_BYTES {
        return Err("快照大小无效。".to_string());
    }
    state
        .recovery
        .save(&display_name, &kind, &bytes)
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn recovery_list(state: State<'_, AppState>) -> Result<Vec<RecoverySnapshot>, String> {
    state.recovery.list().map_err(|error| error.to_string())
}

#[tauri::command]
fn recovery_load(state: State<'_, AppState>, id: String) -> Result<Vec<u8>, String> {
    state.recovery.load(&id).map_err(|error| error.to_string())
}

#[tauri::command]
fn recovery_delete(state: State<'_, AppState>, id: String) -> Result<(), String> {
    state
        .recovery
        .delete(&id)
        .map_err(|error| error.to_string())
}

#[tauri::command]
fn print_window(window: tauri::WebviewWindow) -> Result<(), String> {
    window.print().map_err(|error| format!("打印失败：{error}"))
}

#[tauri::command]
fn export_pdf(
    input: String,
    output_directory: String,
) -> Result<office_pdf::PdfExportResult, String> {
    let converter = office_pdf::LibreOfficeConverter::discover().ok_or_else(|| {
        "未找到 LibreOffice。Mini Office 前端会使用内置 PDF 导出；LibreOffice 转换仅作为 PoC 对照。"
            .to_string()
    })?;
    converter
        .convert_to_pdf(&input, &output_directory)
        .map_err(|error| error.to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .setup(|app| {
            office_platform::ensure_directories()?;
            let app_data = office_platform::app_data_dir();
            let state = AppState {
                recent: RecentStore::open(&app_data)?,
                recovery: RecoveryStore::open(&app_data)?,
            };
            app.manage(state);
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            read_file,
            write_file,
            list_fonts,
            inspect_document,
            platform_info,
            app_version,
            recent_list,
            recent_record,
            recovery_save,
            recovery_list,
            recovery_load,
            recovery_delete,
            export_pdf,
            print_window
        ])
        .run(tauri::generate_context!())
        .expect("Mini Office failed to start");
}

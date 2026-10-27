use office_core::{FileInfo, OfficeKind, PlatformInfo, MAX_INPUT_BYTES};
use office_storage::{
    atomic_write, RecentFile, RecentStore, RecoverySnapshot, RecoveryStore, SaveOutcome,
};
use serde::Serialize;
use std::path::Path;
use std::sync::{atomic::{AtomicBool, Ordering}, Mutex};
use tauri::{Emitter, Manager, State};

struct AppState {
    recent: RecentStore,
    recovery: RecoveryStore,
    // 功能：在前端就绪前暂存系统文件事件，并统一保护窗口关闭和退出；时间：2026-10-18；作者：lq。
    pending_files: Mutex<Vec<String>>,
    close_authorized: AtomicBool,
}

/// 功能：原子领取待打开文件，避免冷启动时系统事件先于页面监听而丢失；时间：2026-10-18；作者：lq。
#[tauri::command]
fn take_pending_files(state: State<'_, AppState>) -> Result<Vec<String>, String> {
    let mut pending = state.pending_files.lock().map_err(|error| error.to_string())?;
    Ok(std::mem::take(&mut *pending))
}

/// 功能：前端逐份完成保存确认后允许应用退出；时间：2026-10-18；作者：lq。
#[tauri::command]
fn confirm_app_close(app: tauri::AppHandle, state: State<'_, AppState>) {
    state.close_authorized.store(true, Ordering::SeqCst);
    app.exit(0);
}

/// 功能：接收 macOS 文件 URL 并唤醒应用，中文与空格路径由 URL 库解码；时间：2026-10-18；作者：lq。
#[cfg(target_os = "macos")]
fn queue_open_files(app: &tauri::AppHandle, urls: Vec<tauri::Url>) {
    let state = app.state::<AppState>();
    if let Ok(mut pending) = state.pending_files.lock() {
        for url in urls {
            let Ok(path) = url.to_file_path() else { continue };
            if !path.is_file() { continue; }
            let path = path.canonicalize().unwrap_or(path).to_string_lossy().into_owned();
            if !pending.contains(&path) { pending.push(path); }
        }
    }
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.unminimize();
        let _ = window.show();
        let _ = window.set_focus();
    }
    let _ = app.emit("mo:open-files", ());
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
    theme_colors: Vec<(String, String)>,
    major_font: Option<String>,
    minor_font: Option<String>,
    relationship_count: usize,
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

/// 功能：在线程池扫描字体，避免冷启动字体文件读取阻塞原生窗口事件和加载动画；时间：2026-10-24；作者：lq。
#[tauri::command]
async fn list_fonts() -> Result<Vec<office_font::FontInfo>, String> {
    tauri::async_runtime::spawn_blocking(|| office_font::cached_fonts().to_vec())
        .await
        .map_err(|error| format!("字体扫描失败：{error}"))
}

#[tauri::command]
fn inspect_document(path: String) -> Result<DocumentInspectionDto, String> {
    let bytes = read_file(path.clone())?;
    let name = Path::new(&path)
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("document");
    let inspection = office_ooxml::analyze(&bytes, name).map_err(|error| error.to_string())?;
    // 功能：把主题和关系摘要暴露给兼容性报告，便于后续编辑器复用统一的主题继承结果；时间：2026-10-28；作者：lq。
    let theme_part = match inspection.kind {
        OfficeKind::Docx => "word/theme/theme1.xml",
        OfficeKind::Xlsx => "xl/theme/theme1.xml",
        OfficeKind::Pptx => "ppt/theme/theme1.xml",
    };
    let theme = office_ooxml::read_text_entry(&bytes, theme_part)
        .map(|xml| office_ooxml::parse_theme(&xml))
        .unwrap_or_default();
    let relationship_count = office_ooxml::entry_names(&bytes)
        .unwrap_or_default()
        .into_iter()
        .filter(|entry| entry.ends_with(".rels"))
        .filter_map(|entry| office_ooxml::read_text_entry(&bytes, &entry).ok())
        .map(|xml| office_ooxml::relationship_records(&xml).len())
        .sum();
    Ok(DocumentInspectionDto {
        kind: inspection.kind,
        level: inspection.capability.level,
        main_part: inspection.main_part,
        entry_count: inspection.entry_count,
        warnings: inspection.capability.warnings,
        features: inspection.capability.features,
        theme_colors: theme.colors.into_iter().collect(),
        major_font: theme.major_font,
        minor_font: theme.minor_font,
        relationship_count,
    })
}

/// 功能：将旧版 DOC/XLS/PPT 转换为可编辑的 OOXML 副本，供前端兼容模式调用；时间：2026-10-27；作者：lq。
#[tauri::command]
fn convert_legacy_document(input: String) -> Result<String, String> {
    if !Path::new(&input).is_file() {
        return Err("源文件不存在。".to_string());
    }
    let output = office_compat::default_output_path(&input)
        .ok_or_else(|| "无法根据源文件扩展名生成转换路径。".to_string())?;
    office_compat::convert_legacy(&input, &output).map_err(|error| error.to_string())?;
    Ok(output.to_string_lossy().into_owned())
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
                pending_files: Mutex::new(Vec::new()),
                close_authorized: AtomicBool::new(false),
            };
            app.manage(state);
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            read_file,
            write_file,
            list_fonts,
            inspect_document,
            convert_legacy_document,
            platform_info,
            app_version,
            recent_list,
            recent_record,
            recovery_save,
            recovery_list,
            recovery_load,
            recovery_delete,
            export_pdf,
            print_window,
            take_pending_files,
            confirm_app_close
        ])
        .build(tauri::generate_context!())
        .expect("Mini Office failed to start")
        .run(|app, event| {
            // 功能：原生关窗和 Cmd+Q 均等待前端保存确认，不直接结束进程；时间：2026-10-18；作者：lq。
            match event {
                tauri::RunEvent::WindowEvent { event: tauri::WindowEvent::CloseRequested { api, .. }, .. } => {
                    if !app.state::<AppState>().close_authorized.load(Ordering::SeqCst) {
                        api.prevent_close();
                        let _ = app.emit("mo:request-close", ());
                    }
                }
                tauri::RunEvent::ExitRequested { api, .. } => {
                    if !app.state::<AppState>().close_authorized.load(Ordering::SeqCst) {
                        api.prevent_exit();
                        let _ = app.emit("mo:request-close", ());
                    }
                }
                #[cfg(target_os = "macos")]
                tauri::RunEvent::Opened { urls } => queue_open_files(app, urls),
                _ => {}
            }
        });
}

use office_core::{
    CompatibilityWarning, OfficeError, Severity, MAX_ENTRIES, MAX_SINGLE_ENTRY_BYTES,
    MAX_UNCOMPRESSED_BYTES,
};
use serde::{Deserialize, Serialize};
use std::io::{Cursor, Read};
use std::path::{Component, Path};
use zip::ZipArchive;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SecurityReport {
    pub ok: bool,
    pub entry_count: usize,
    pub uncompressed_bytes: u64,
    pub has_macros: bool,
    pub has_signature: bool,
    pub has_external_links: bool,
    pub warnings: Vec<CompatibilityWarning>,
}

#[derive(Debug, Clone, Copy)]
pub struct ScanLimits {
    pub max_entries: usize,
    pub max_uncompressed_bytes: u64,
    pub max_single_entry_bytes: u64,
    pub max_compression_ratio: u64,
}

impl Default for ScanLimits {
    fn default() -> Self {
        Self {
            max_entries: MAX_ENTRIES,
            max_uncompressed_bytes: MAX_UNCOMPRESSED_BYTES,
            max_single_entry_bytes: MAX_SINGLE_ENTRY_BYTES,
            max_compression_ratio: 2_000,
        }
    }
}

pub fn is_zip(bytes: &[u8]) -> bool {
    bytes.len() >= 4 && bytes[0] == 0x50 && bytes[1] == 0x4b
}

pub fn is_legacy_ole(bytes: &[u8]) -> bool {
    bytes.len() >= 8
        && bytes[0] == 0xd0
        && bytes[1] == 0xcf
        && bytes[2] == 0x11
        && bytes[3] == 0xe0
        && bytes[4] == 0xa1
        && bytes[5] == 0xb1
        && bytes[6] == 0x1a
        && bytes[7] == 0xe1
}

fn unsafe_entry_name(name: &str) -> bool {
    if name.contains('\0') || name.contains('\\') {
        return true;
    }
    let path = Path::new(name);
    path.is_absolute()
        || path.components().any(|component| matches!(component, Component::ParentDir))
        || name.starts_with('/')
}

/// Scan an OOXML container before parsing. This is intentionally conservative:
/// hostile archives are rejected rather than partially loaded.
pub fn inspect_zip(bytes: &[u8], file_name: &str) -> Result<SecurityReport, OfficeError> {
    if is_legacy_ole(bytes) {
        return Err(OfficeError::Unsupported(
            "这是旧版二进制 Office 格式（.doc/.xls/.ppt），V1 暂不支持。".to_string(),
        ));
    }
    if !is_zip(bytes) {
        return Err(OfficeError::InvalidPackage(
            "文件不是有效的 OOXML 压缩包。".to_string(),
        ));
    }
    let limits = ScanLimits::default();
    let mut archive = ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| OfficeError::InvalidPackage(format!("无法解压：{error}")))?;
    if archive.len() > limits.max_entries {
        return Err(OfficeError::Security(format!(
            "文件包含过多部件（{}），超过安全上限。",
            archive.len()
        )));
    }
    let mut uncompressed_bytes = 0u64;
    let mut warnings = Vec::new();
    let mut has_macros = false;
    let mut has_signature = false;
    let mut has_external_links = false;
    let mut has_content_types = false;

    for index in 0..archive.len() {
        let mut entry = archive
            .by_index(index)
            .map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
        let name = entry.name().to_string();
        if unsafe_entry_name(&name) {
            return Err(OfficeError::Security(format!(
                "压缩包包含不安全路径：{name}"
            )));
        }
        if entry.size() > limits.max_single_entry_bytes {
            return Err(OfficeError::Security(format!(
                "部件 {name} 解压后过大，已阻止打开。"
            )));
        }
        if name == "[Content_Types].xml" {
            has_content_types = true;
        }
        uncompressed_bytes = uncompressed_bytes.saturating_add(entry.size());
        if uncompressed_bytes > limits.max_uncompressed_bytes {
            return Err(OfficeError::Security(
                "文件解压后体积超过安全上限，已阻止打开。".to_string(),
            ));
        }
        if entry.size() > 0 {
            let ratio = entry.size() / entry.compressed_size().max(1);
            if ratio > limits.max_compression_ratio {
                return Err(OfficeError::Security(format!(
                    "部件 {name} 的压缩比异常（{ratio}:1），疑似压缩炸弹。"
                )));
            }
        }
        let lower = name.to_ascii_lowercase();
        if lower.ends_with("vbaproject.bin") {
            has_macros = true;
        }
        if lower.starts_with("_xmlsignatures/") {
            has_signature = true;
        }
        if lower.ends_with(".rels") {
            let mut xml = String::new();
            if entry.read_to_string(&mut xml).is_ok()
                && (xml.contains("TargetMode=\"External\"")
                    || xml.contains("TargetMode='External'"))
            {
                has_external_links = true;
            }
        }
    }

    if !has_content_types {
        return Err(OfficeError::InvalidPackage(
            "缺少 [Content_Types].xml，无法确认文件类型。".to_string(),
        ));
    }
    if !file_name.is_empty() && file_name.to_ascii_lowercase().ends_with(".xlsm") && !has_macros {
        warnings.push(CompatibilityWarning::new(
            "xlsm.macroMissing",
            "文件名标记为启用宏，但未发现 VBA 部件。",
            Severity::Info,
        ));
    }
    if has_macros {
        warnings.push(CompatibilityWarning::new(
            "security.macro",
            "检测到 VBA 宏。宏不会执行，文档将以只读或副本模式处理。",
            Severity::Blocking,
        ));
    }
    if has_signature {
        warnings.push(CompatibilityWarning::new(
            "security.signature",
            "检测到数字签名。编辑会使签名失效。",
            Severity::Blocking,
        ));
    }
    if has_external_links {
        warnings.push(CompatibilityWarning::new(
            "security.externalLinks",
            "检测到外部链接。外部资源不会被自动加载。",
            Severity::Warning,
        ));
    }

    Ok(SecurityReport {
        ok: true,
        entry_count: archive.len(),
        uncompressed_bytes,
        has_macros,
        has_signature,
        has_external_links,
        warnings,
    })
}

pub fn validate_zip(bytes: &[u8]) -> Result<(), OfficeError> {
    inspect_zip(bytes, "")?;
    Ok(())
}

pub fn ensure_safe_output_path(
    path: impl AsRef<Path>,
    allowed_extensions: &[&str],
) -> Result<(), OfficeError> {
    let path = path.as_ref();
    if path.as_os_str().is_empty() {
        return Err(OfficeError::Storage("保存路径为空。".to_string()));
    }
    if path.is_dir() {
        return Err(OfficeError::Storage("保存路径是目录，无法写入。".to_string()));
    }
    let extension = path
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    if !allowed_extensions
        .iter()
        .any(|allowed| allowed.trim_start_matches('.').eq_ignore_ascii_case(&extension))
    {
        return Err(OfficeError::Security(format!(
            "不允许写入 .{extension} 文件。"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn package(entries: &[(&str, &str)]) -> Vec<u8> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        for (name, content) in entries {
            writer
                .start_file(name.to_string(), SimpleFileOptions::default())
                .unwrap();
            writer.write_all(content.as_bytes()).unwrap();
        }
        writer.finish().unwrap().into_inner()
    }

    #[test]
    fn accepts_clean_package() {
        let bytes = package(&[
            ("[Content_Types].xml", "<Types/>"),
            ("word/document.xml", "<w:document/>"),
        ]);
        let report = inspect_zip(&bytes, "clean.docx").unwrap();
        assert!(report.ok);
        assert!(!report.has_macros);
    }

    #[test]
    fn rejects_path_traversal() {
        let bytes = package(&[("[Content_Types].xml", "<Types/>"), ("../evil.xml", "x")]);
        let error = inspect_zip(&bytes, "bad.docx").unwrap_err();
        assert!(matches!(error, OfficeError::Security(_)));
    }

    #[test]
    fn detects_macros_and_signatures() {
        let bytes = package(&[
            ("[Content_Types].xml", "<Types/>"),
            ("word/document.xml", "<w:document/>"),
            ("word/vbaProject.bin", "macro"),
            ("_xmlsignatures/sig1.xml", "<sig/>"),
        ]);
        let report = inspect_zip(&bytes, "macro.docm").unwrap();
        assert!(report.has_macros);
        assert!(report.has_signature);
        assert!(report.warnings.iter().any(|warning| warning.severity == Severity::Blocking));
    }

    #[test]
    fn validates_output_extension() {
        assert!(ensure_safe_output_path("/tmp/a.docx", &["docx", "xlsx", "pptx"]).is_ok());
        assert!(ensure_safe_output_path("/tmp/a.sh", &["docx"]).is_err());
    }
}

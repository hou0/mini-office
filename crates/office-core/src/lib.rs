use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::path::Path;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum OfficeKind {
    Docx,
    Xlsx,
    Pptx,
}

impl OfficeKind {
    pub fn extension(self) -> &'static str {
        match self {
            OfficeKind::Docx => ".docx",
            OfficeKind::Xlsx => ".xlsx",
            OfficeKind::Pptx => ".pptx",
        }
    }

    pub fn mime(self) -> &'static str {
        match self {
            OfficeKind::Docx => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
            OfficeKind::Xlsx => {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
            OfficeKind::Pptx => {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            }
        }
    }

    pub fn from_extension(path: impl AsRef<Path>) -> Option<Self> {
        let extension = path
            .as_ref()
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        match extension.as_str() {
            "docx" | "docm" | "dotx" => Some(OfficeKind::Docx),
            "xlsx" | "xlsm" | "xltx" => Some(OfficeKind::Xlsx),
            "pptx" | "pptm" | "potx" => Some(OfficeKind::Pptx),
            _ => None,
        }
    }

    pub fn main_part(self) -> &'static str {
        match self {
            OfficeKind::Docx => "word/document.xml",
            OfficeKind::Xlsx => "xl/workbook.xml",
            OfficeKind::Pptx => "ppt/presentation.xml",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Severity {
    Info,
    Warning,
    Blocking,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CompatibilityWarning {
    pub code: String,
    pub message: String,
    pub severity: Severity,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub detail: Option<String>,
}

impl CompatibilityWarning {
    pub fn new(code: impl Into<String>, message: impl Into<String>, severity: Severity) -> Self {
        Self {
            code: code.into(),
            message: message.into(),
            severity,
            detail: None,
        }
    }

    pub fn detail(mut self, detail: impl Into<String>) -> Self {
        self.detail = Some(detail.into());
        self
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CapabilityFeature {
    pub code: String,
    pub label: String,
    pub supported: bool,
    pub severity: Severity,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note: Option<String>,
}

impl CapabilityFeature {
    pub fn new(
        code: impl Into<String>,
        label: impl Into<String>,
        supported: bool,
        severity: Severity,
    ) -> Self {
        Self {
            code: code.into(),
            label: label.into(),
            supported,
            severity,
            note: None,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CapabilityReport {
    /// L0 = cannot open, L1 = preview only, L2 = edit but save as copy, L3 = safe round trip.
    pub level: u8,
    pub warnings: Vec<CompatibilityWarning>,
    pub features: Vec<CapabilityFeature>,
    pub missing_fonts: Vec<String>,
    pub scanned_at: i64,
}

impl Default for CapabilityReport {
    fn default() -> Self {
        Self {
            level: 3,
            warnings: Vec::new(),
            features: Vec::new(),
            missing_fonts: Vec::new(),
            scanned_at: chrono::Utc::now().timestamp_millis(),
        }
    }
}

impl CapabilityReport {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn warn(&mut self, warning: CompatibilityWarning) {
        self.warnings.push(warning);
    }

    pub fn feature(&mut self, feature: CapabilityFeature) {
        self.features.push(feature);
    }

    pub fn blocking(&self) -> bool {
        self.warnings
            .iter()
            .any(|warning| warning.severity == Severity::Blocking)
    }

    /// 功能：根据已扫描能力计算轻量兼容性百分比，作为真实样本回归的统一指标；时间：2026-10-28；作者：lq。
    pub fn compatibility_score(&self) -> u8 {
        if self.features.is_empty() {
            return self.level.saturating_mul(25);
        }
        let supported = self.features.iter().filter(|feature| feature.supported).count();
        ((supported * 100) / self.features.len()).min(100) as u8
    }

    pub fn finalize(mut self, readonly_kind: bool) -> Self {
        if self.blocking() {
            self.level = 1;
        } else if self
            .warnings
            .iter()
            .any(|warning| warning.severity == Severity::Warning && warning.code != "fonts.missing")
        {
            self.level = 2;
        } else {
            self.level = 3;
        }
        if readonly_kind {
            self.level = 1;
        }
        self.scanned_at = chrono::Utc::now().timestamp_millis();
        self
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FileInfo {
    pub path: String,
    pub name: String,
    pub size: u64,
    pub modified_at: Option<i64>,
    pub sha256: String,
    pub kind: Option<OfficeKind>,
}

impl FileInfo {
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self, OfficeError> {
        let path = path.as_ref();
        let metadata = std::fs::metadata(path)?;
        let bytes = std::fs::read(path)?;
        let modified_at = metadata
            .modified()
            .ok()
            .and_then(|time| time.duration_since(std::time::UNIX_EPOCH).ok())
            .map(|duration| duration.as_millis() as i64);
        Ok(Self {
            path: path.to_string_lossy().to_string(),
            name: path
                .file_name()
                .and_then(|value| value.to_str())
                .unwrap_or("未命名文档")
                .to_string(),
            size: metadata.len(),
            modified_at,
            sha256: sha256_hex(&bytes),
            kind: OfficeKind::from_extension(path),
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum PlatformKind {
    Macos,
    Windows,
    Linux,
    Unknown,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PlatformInfo {
    pub os: PlatformKind,
    pub arch: String,
    pub family: String,
    pub app_data_dir: String,
    pub temp_dir: String,
}

impl PlatformInfo {
    pub fn current(app_data_dir: String, temp_dir: String) -> Self {
        let os = if cfg!(target_os = "macos") {
            PlatformKind::Macos
        } else if cfg!(target_os = "windows") {
            PlatformKind::Windows
        } else if cfg!(target_os = "linux") {
            PlatformKind::Linux
        } else {
            PlatformKind::Unknown
        };
        Self {
            os,
            arch: std::env::consts::ARCH.to_string(),
            family: std::env::consts::FAMILY.to_string(),
            app_data_dir,
            temp_dir,
        }
    }
}

#[derive(Debug, thiserror::Error)]
pub enum OfficeError {
    #[error("I/O 错误：{0}")]
    Io(#[from] std::io::Error),
    #[error("文件格式不受支持：{0}")]
    Unsupported(String),
    #[error("文件包无效或已损坏：{0}")]
    InvalidPackage(String),
    #[error("安全策略拒绝了该文件：{0}")]
    Security(String),
    #[error("保存失败：{0}")]
    Storage(String),
    #[error("PDF 导出失败：{0}")]
    Pdf(String),
    #[error("{0}")]
    Other(String),
}

pub fn sha256_hex(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    hex::encode(hasher.finalize())
}

pub fn now_millis() -> i64 {
    chrono::Utc::now().timestamp_millis()
}

pub const MAX_ENTRIES: usize = 6_000;
pub const MAX_UNCOMPRESSED_BYTES: u64 = 600 * 1024 * 1024;
pub const MAX_SINGLE_ENTRY_BYTES: u64 = 250 * 1024 * 1024;
pub const MAX_INPUT_BYTES: u64 = 250 * 1024 * 1024;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detects_office_kinds() {
        assert_eq!(OfficeKind::from_extension("a.docx"), Some(OfficeKind::Docx));
        assert_eq!(OfficeKind::from_extension("a.xlsm"), Some(OfficeKind::Xlsx));
        assert_eq!(OfficeKind::from_extension("a.PPTX"), Some(OfficeKind::Pptx));
        assert_eq!(OfficeKind::from_extension("a.pdf"), None);
    }

    #[test]
    fn capability_levels_follow_warnings() {
        let mut report = CapabilityReport::new();
        report.warn(CompatibilityWarning::new(
            "fonts.missing",
            "missing font",
            Severity::Warning,
        ));
        assert_eq!(report.clone().finalize(false).level, 3);
        report.warn(CompatibilityWarning::new(
            "docx.revisions",
            "tracked changes",
            Severity::Warning,
        ));
        assert_eq!(report.finalize(false).level, 2);

        let mut blocking = CapabilityReport::new();
        blocking.warn(CompatibilityWarning::new(
            "docx.macro",
            "macro",
            Severity::Blocking,
        ));
        assert_eq!(blocking.finalize(false).level, 1);
    }

    // 功能：验证兼容性评分按已支持能力比例计算；时间：2026-10-28；作者：lq。
    #[test]
    fn compatibility_score_uses_supported_features() {
        let mut report = CapabilityReport::new();
        report.feature(CapabilityFeature::new("a", "A", true, Severity::Info));
        report.feature(CapabilityFeature::new("b", "B", false, Severity::Warning));
        assert_eq!(report.compatibility_score(), 50);
    }

    #[test]
    fn sha256_is_stable() {
        assert_eq!(
            sha256_hex(b"mini-office"),
            "ac884570d08188c0104e9e0190fb7a43bd95d099b4c6a70c00f19db5496a68d1"
        );
    }
}

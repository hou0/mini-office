use office_core::{OfficeError, MAX_INPUT_BYTES};
use office_oxide::Document;
use std::path::{Path, PathBuf};

/// 功能：表示可转换的旧版 Office 文件类型；时间：2026-10-27；作者：lq。
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegacyKind {
    Doc,
    Xls,
    Ppt,
}

impl LegacyKind {
    pub fn from_path(path: impl AsRef<Path>) -> Option<Self> {
        match path.as_ref().extension()?.to_str()?.to_ascii_lowercase().as_str() {
            "doc" => Some(Self::Doc),
            "xls" => Some(Self::Xls),
            "ppt" => Some(Self::Ppt),
            _ => None,
        }
    }

    pub fn target_extension(self) -> &'static str {
        match self {
            Self::Doc => "docx",
            Self::Xls => "xlsx",
            Self::Ppt => "pptx",
        }
    }
}

/// 功能：为旧版 Office 文件生成默认 OOXML 输出路径；时间：2026-10-27；作者：lq。
pub fn default_output_path(input: impl AsRef<Path>) -> Option<PathBuf> {
    let kind = LegacyKind::from_path(&input)?;
    let path = input.as_ref();
    let stem = path.file_stem()?.to_string_lossy();
    Some(path.with_file_name(format!("{}_兼容转换.{}", stem, kind.target_extension())))
}

/// 功能：使用纯 Rust office_oxide 将 DOC/XLS/PPT 转换为 OOXML，保持主应用不依赖大型办公套件；时间：2026-10-27；作者：lq。
pub fn convert_legacy(input: impl AsRef<Path>, output: impl AsRef<Path>) -> Result<(), OfficeError> {
    let input = input.as_ref();
    let output = output.as_ref();
    if LegacyKind::from_path(input).is_none() {
        return Err(OfficeError::Unsupported("只支持 DOC、XLS、PPT 旧版文件转换。".to_string()));
    }
    let size = std::fs::metadata(input).map_err(OfficeError::Io)?.len();
    if size > MAX_INPUT_BYTES {
        return Err(OfficeError::Security("旧版 Office 文件超过 250 MB 安全上限。".to_string()));
    }
    if !matches!(output.extension().and_then(|value| value.to_str()).map(|value| value.to_ascii_lowercase()).as_deref(), Some("docx" | "xlsx" | "pptx")) {
        return Err(OfficeError::Unsupported("输出文件必须是 DOCX、XLSX 或 PPTX。".to_string()));
    }
    let document = Document::open(input).map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
    document.save_as(output).map_err(|error| OfficeError::Storage(error.to_string()))
}

/// 功能：在内存与临时目录之间完成旧版 Office 转换，供拖拽和非文件路径入口复用；时间：2026-10-28；作者：lq。
pub fn convert_legacy_bytes(bytes: &[u8], name: &str) -> Result<Vec<u8>, OfficeError> {
    let kind = LegacyKind::from_path(name)
        .ok_or_else(|| OfficeError::Unsupported("只支持 DOC、XLS、PPT 旧版文件转换。".to_string()))?;
    if bytes.len() as u64 > MAX_INPUT_BYTES {
        return Err(OfficeError::Security("旧版 Office 文件超过 250 MB 安全上限。".to_string()));
    }
    let directory = tempfile::tempdir().map_err(OfficeError::Io)?;
    let input = directory.path().join(name);
    let output = directory.path().join(format!("converted.{}", kind.target_extension()));
    std::fs::write(&input, bytes).map_err(OfficeError::Io)?;
    convert_legacy(&input, &output)?;
    std::fs::read(output).map_err(OfficeError::Io)
}

#[cfg(test)]
mod tests {
    use super::{convert_legacy_bytes, default_output_path, LegacyKind};

    #[test]
    fn detects_legacy_extensions_and_output_name() {
        assert_eq!(LegacyKind::from_path("报告.XLS"), Some(LegacyKind::Xls));
        assert_eq!(LegacyKind::from_path("演示.ppt"), Some(LegacyKind::Ppt));
        assert_eq!(default_output_path("合同.doc").unwrap().to_string_lossy(), "合同_兼容转换.docx");
    }

    // 功能：验证内存入口先校验格式并拒绝不支持的扩展名；时间：2026-10-28；作者：lq。
    #[test]
    fn rejects_unsupported_memory_input() {
        assert!(convert_legacy_bytes(b"invalid", "sample.txt").is_err());
    }
}

use office_core::OfficeError;
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
    if !matches!(output.extension().and_then(|value| value.to_str()).map(|value| value.to_ascii_lowercase()).as_deref(), Some("docx" | "xlsx" | "pptx")) {
        return Err(OfficeError::Unsupported("输出文件必须是 DOCX、XLSX 或 PPTX。".to_string()));
    }
    let document = Document::open(input).map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
    document.save_as(output).map_err(|error| OfficeError::Storage(error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::{default_output_path, LegacyKind};

    #[test]
    fn detects_legacy_extensions_and_output_name() {
        assert_eq!(LegacyKind::from_path("报告.XLS"), Some(LegacyKind::Xls));
        assert_eq!(LegacyKind::from_path("演示.ppt"), Some(LegacyKind::Ppt));
        assert_eq!(default_output_path("合同.doc").unwrap().to_string_lossy(), "合同_兼容转换.docx");
    }
}

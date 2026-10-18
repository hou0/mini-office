use office_core::{CapabilityReport, OfficeError, OfficeKind};
use office_ooxml::{analyze, read_text_entry, PackageInspection};

#[derive(Debug, Clone, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DocInspection {
    pub kind: OfficeKind,
    pub paragraph_count: usize,
    pub table_count: usize,
    pub capability: CapabilityReport,
}

pub fn inspect(bytes: &[u8], file_name: &str) -> Result<DocInspection, OfficeError> {
    let inspection = analyze(bytes, file_name)?;
    if inspection.kind != OfficeKind::Docx {
        return Err(OfficeError::Unsupported("该文件不是 DOCX。".to_string()));
    }
    let document = read_text_entry(bytes, "word/document.xml").unwrap_or_default();
    let paragraph_count = document.matches("<w:p ").count() + document.matches("<w:p>").count();
    let table_count = document.matches("<w:tbl>").count();
    Ok(DocInspection {
        kind: OfficeKind::Docx,
        paragraph_count,
        table_count,
        capability: inspection.capability,
    })
}

pub fn inspect_package(inspection: PackageInspection) -> Result<DocInspection, OfficeError> {
    if inspection.kind != OfficeKind::Docx {
        return Err(OfficeError::Unsupported("该文件不是 DOCX。".to_string()));
    }
    Ok(DocInspection {
        kind: OfficeKind::Docx,
        paragraph_count: 0,
        table_count: 0,
        capability: inspection.capability,
    })
}

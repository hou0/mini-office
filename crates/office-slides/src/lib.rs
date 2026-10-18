use office_core::{CapabilityReport, OfficeError, OfficeKind};
use office_ooxml::{analyze, entry_names, PackageInspection};

#[derive(Debug, Clone, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SlidesInspection {
    pub kind: OfficeKind,
    pub slide_count: usize,
    pub capability: CapabilityReport,
}

pub fn inspect(bytes: &[u8], file_name: &str) -> Result<SlidesInspection, OfficeError> {
    let inspection = analyze(bytes, file_name)?;
    if inspection.kind != OfficeKind::Pptx {
        return Err(OfficeError::Unsupported("该文件不是 PPTX。".to_string()));
    }
    let slide_count = entry_names(bytes)?
        .iter()
        .filter(|name| name.starts_with("ppt/slides/slide") && name.ends_with(".xml"))
        .count();
    Ok(SlidesInspection {
        kind: OfficeKind::Pptx,
        slide_count,
        capability: inspection.capability,
    })
}

pub fn from_package(inspection: PackageInspection) -> Result<SlidesInspection, OfficeError> {
    if inspection.kind != OfficeKind::Pptx {
        return Err(OfficeError::Unsupported("该文件不是 PPTX。".to_string()));
    }
    Ok(SlidesInspection {
        kind: OfficeKind::Pptx,
        slide_count: 0,
        capability: inspection.capability,
    })
}

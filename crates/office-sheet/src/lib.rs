use office_core::{CapabilityReport, OfficeError, OfficeKind};
use office_ooxml::{analyze, entry_names, PackageInspection};

#[derive(Debug, Clone, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetInspection {
    pub kind: OfficeKind,
    pub worksheet_count: usize,
    pub capability: CapabilityReport,
}

pub fn inspect(bytes: &[u8], file_name: &str) -> Result<SheetInspection, OfficeError> {
    let inspection = analyze(bytes, file_name)?;
    if inspection.kind != OfficeKind::Xlsx {
        return Err(OfficeError::Unsupported("该文件不是 XLSX。".to_string()));
    }
    let worksheet_count = entry_names(bytes)?
        .iter()
        .filter(|name| name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml"))
        .count();
    Ok(SheetInspection {
        kind: OfficeKind::Xlsx,
        worksheet_count,
        capability: inspection.capability,
    })
}

pub fn from_package(inspection: PackageInspection) -> Result<SheetInspection, OfficeError> {
    if inspection.kind != OfficeKind::Xlsx {
        return Err(OfficeError::Unsupported("该文件不是 XLSX。".to_string()));
    }
    Ok(SheetInspection {
        kind: OfficeKind::Xlsx,
        worksheet_count: 0,
        capability: inspection.capability,
    })
}

use office_core::{CapabilityReport, OfficeError, OfficeKind};
use office_ooxml::{analyze, entry_names, OoxmlPackage, PackageInspection};

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

/// 功能：只替换 XLSX 工作簿主部件并保留其他原始部件；时间：2026-10-28；作者：lq。
pub fn replace_workbook_part(source: &[u8], workbook_xml: impl Into<Vec<u8>>) -> Result<Vec<u8>, OfficeError> {
    let mut package = OoxmlPackage::open(source.to_vec())?;
    package.replace("xl/workbook.xml", workbook_xml);
    package.finish()
}

#[cfg(test)]
mod tests {
    use super::replace_workbook_part;
    use std::io::{Cursor, Read, Write};
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    // 功能：验证 XLSX 主部件增量替换不破坏主题部件；时间：2026-10-28；作者：lq。
    #[test]
    fn replaces_workbook_and_preserves_theme() {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        writer.start_file("xl/workbook.xml", SimpleFileOptions::default()).unwrap();
        writer.write_all(b"old").unwrap();
        writer.start_file("xl/theme/theme1.xml", SimpleFileOptions::default()).unwrap();
        writer.write_all(b"theme").unwrap();
        let source = writer.finish().unwrap().into_inner();
        let output = replace_workbook_part(&source, b"new".to_vec()).unwrap();
        let mut archive = zip::ZipArchive::new(Cursor::new(output)).unwrap();
        let mut workbook = String::new();
        archive.by_name("xl/workbook.xml").unwrap().read_to_string(&mut workbook).unwrap();
        assert_eq!(workbook, "new");
        let mut theme = String::new();
        archive.by_name("xl/theme/theme1.xml").unwrap().read_to_string(&mut theme).unwrap();
        assert_eq!(theme, "theme");
    }
}

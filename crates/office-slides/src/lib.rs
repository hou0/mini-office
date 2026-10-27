use office_core::{CapabilityReport, OfficeError, OfficeKind};
use office_ooxml::{analyze, entry_names, OoxmlPackage, PackageInspection};

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

/// 功能：只替换 PPTX 演示文稿主部件并保留其他原始部件；时间：2026-10-28；作者：lq。
pub fn replace_presentation_part(source: &[u8], presentation_xml: impl Into<Vec<u8>>) -> Result<Vec<u8>, OfficeError> {
    let mut package = OoxmlPackage::open(source.to_vec())?;
    package.replace("ppt/presentation.xml", presentation_xml);
    package.finish()
}

#[cfg(test)]
mod tests {
    use super::replace_presentation_part;
    use std::io::{Cursor, Read, Write};
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    // 功能：验证 PPTX 主部件增量替换不破坏主题部件；时间：2026-10-28；作者：lq。
    #[test]
    fn replaces_presentation_and_preserves_theme() {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        writer.start_file("ppt/presentation.xml", SimpleFileOptions::default()).unwrap();
        writer.write_all(b"old").unwrap();
        writer.start_file("ppt/theme/theme1.xml", SimpleFileOptions::default()).unwrap();
        writer.write_all(b"theme").unwrap();
        let source = writer.finish().unwrap().into_inner();
        let output = replace_presentation_part(&source, b"new".to_vec()).unwrap();
        let mut archive = zip::ZipArchive::new(Cursor::new(output)).unwrap();
        let mut presentation = String::new();
        archive.by_name("ppt/presentation.xml").unwrap().read_to_string(&mut presentation).unwrap();
        assert_eq!(presentation, "new");
        let mut theme = String::new();
        archive.by_name("ppt/theme/theme1.xml").unwrap().read_to_string(&mut theme).unwrap();
        assert_eq!(theme, "theme");
    }
}

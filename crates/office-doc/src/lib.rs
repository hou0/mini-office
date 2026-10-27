use office_core::{CapabilityReport, OfficeError, OfficeKind};
use office_ooxml::{analyze, read_text_entry, OoxmlPackage, PackageInspection};

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

/// 功能：只替换 DOCX 主文档部件并保留其他原始部件，供后续 Rust 保存链路复用；时间：2026-10-28；作者：lq。
pub fn replace_document_part(source: &[u8], document_xml: impl Into<Vec<u8>>) -> Result<Vec<u8>, OfficeError> {
    let mut package = OoxmlPackage::open(source.to_vec())?;
    package.replace("word/document.xml", document_xml);
    package.finish()
}

#[cfg(test)]
mod tests {
    use super::replace_document_part;
    use std::io::{Cursor, Read, Write};
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn sample() -> Vec<u8> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        writer.start_file("word/document.xml", SimpleFileOptions::default()).unwrap();
        writer.write_all(b"old").unwrap();
        writer.start_file("word/theme/theme1.xml", SimpleFileOptions::default()).unwrap();
        writer.write_all(b"theme").unwrap();
        writer.finish().unwrap().into_inner()
    }

    #[test]
    fn replaces_only_document_part() {
        let output = replace_document_part(&sample(), b"new".to_vec()).unwrap();
        let mut archive = zip::ZipArchive::new(Cursor::new(output)).unwrap();
        let mut document = String::new();
        archive.by_name("word/document.xml").unwrap().read_to_string(&mut document).unwrap();
        assert_eq!(document, "new");
        let mut theme = String::new();
        archive.by_name("word/theme/theme1.xml").unwrap().read_to_string(&mut theme).unwrap();
        assert_eq!(theme, "theme");
    }
}

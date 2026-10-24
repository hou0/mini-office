use office_core::{
    CapabilityFeature, CapabilityReport, CompatibilityWarning, OfficeError, OfficeKind, Severity,
};
use office_security::{inspect_zip, validate_zip, SecurityReport};
use quick_xml::events::Event;
use quick_xml::Reader;
use serde::{Deserialize, Serialize};
use std::io::{Cursor, Read};
use zip::ZipArchive;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PackageEntry {
    pub name: String,
    pub size: u64,
    pub compressed_size: u64,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PackageInspection {
    pub kind: OfficeKind,
    pub main_part: String,
    pub entry_count: usize,
    pub entries: Vec<PackageEntry>,
    pub security: SecurityReport,
    pub capability: CapabilityReport,
}

pub fn open_archive(bytes: &[u8]) -> Result<ZipArchive<Cursor<&[u8]>>, OfficeError> {
    ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| OfficeError::InvalidPackage(format!("无法解压 OOXML 包：{error}")))
}

pub fn entry_names(bytes: &[u8]) -> Result<Vec<String>, OfficeError> {
    let mut archive = open_archive(bytes)?;
    let mut names = Vec::with_capacity(archive.len());
    for index in 0..archive.len() {
        let entry = archive
            .by_index(index)
            .map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
        if !entry.is_dir() {
            names.push(entry.name().to_string());
        }
    }
    Ok(names)
}

pub fn read_entry(bytes: &[u8], name: &str) -> Result<Vec<u8>, OfficeError> {
    let mut archive = open_archive(bytes)?;
    let mut entry = archive
        .by_name(name)
        .map_err(|_| OfficeError::InvalidPackage(format!("缺少 OOXML 部件：{name}")))?;
    let mut buffer = Vec::with_capacity(entry.size() as usize);
    entry
        .read_to_end(&mut buffer)
        .map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
    Ok(buffer)
}

pub fn read_text_entry(bytes: &[u8], name: &str) -> Result<String, OfficeError> {
    let data = read_entry(bytes, name)?;
    String::from_utf8(data)
        .map_err(|error| OfficeError::InvalidPackage(format!("部件 {name} 不是 UTF-8：{error}")))
}

pub fn relationship_target(xml: &str, suffix: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                if element.name().as_ref() != b"Relationship" {
                    continue;
                }
                let mut is_match = false;
                let mut target = None;
                for attribute in element.attributes().flatten() {
                    let key = attribute.key.as_ref();
                    if key == b"Type" {
                        let value = String::from_utf8_lossy(&attribute.value);
                        is_match = value.ends_with(suffix);
                    } else if key == b"Target" {
                        target = Some(String::from_utf8_lossy(&attribute.value).to_string());
                    }
                }
                if is_match {
                    return target.map(|value| value.trim_start_matches('/').to_string());
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    None
}

pub fn main_part(bytes: &[u8]) -> Result<Option<String>, OfficeError> {
    if let Ok(rels) = read_text_entry(bytes, "_rels/.rels") {
        if let Some(target) = relationship_target(&rels, "/officeDocument") {
            return Ok(Some(target));
        }
    }
    let names = entry_names(bytes)?;
    if names.iter().any(|name| name == "word/document.xml") {
        return Ok(Some("word/document.xml".to_string()));
    }
    if names.iter().any(|name| name == "xl/workbook.xml") {
        return Ok(Some("xl/workbook.xml".to_string()));
    }
    if names.iter().any(|name| name == "ppt/presentation.xml") {
        return Ok(Some("ppt/presentation.xml".to_string()));
    }
    Ok(None)
}

pub fn detect_kind(bytes: &[u8], file_name: &str) -> Result<OfficeKind, OfficeError> {
    if office_security::is_legacy_ole(bytes) {
        return Err(OfficeError::Unsupported(
            "这是旧版二进制 Office 格式（.doc/.xls/.ppt），V1 暂不支持。".to_string(),
        ));
    }
    let main = main_part(bytes)?;
    if let Some(main) = main {
        if main.starts_with("word/") {
            return Ok(OfficeKind::Docx);
        }
        if main.starts_with("xl/") {
            return Ok(OfficeKind::Xlsx);
        }
        if main.starts_with("ppt/") {
            return Ok(OfficeKind::Pptx);
        }
    }
    if let Some(kind) = OfficeKind::from_extension(file_name) {
        return Ok(kind);
    }
    Err(OfficeError::Unsupported(
        "无法识别该 OOXML 文件类型。".to_string(),
    ))
}

fn names_contain(names: &[String], needle: &str) -> bool {
    names.iter().any(|name| name.eq_ignore_ascii_case(needle))
}

fn any_name(names: &[String], predicate: impl Fn(&str) -> bool) -> bool {
    names.iter().any(|name| predicate(name))
}

pub fn analyze(bytes: &[u8], file_name: &str) -> Result<PackageInspection, OfficeError> {
    let security = inspect_zip(bytes, file_name)?;
    let kind = detect_kind(bytes, file_name)?;
    let entries = entry_names(bytes)?;
    let main_part = main_part(bytes)?.unwrap_or_else(|| kind.main_part().to_string());
    if !names_contain(&entries, &main_part) && !any_name(&entries, |name| name.ends_with(".xml")) {
        return Err(OfficeError::InvalidPackage(
            "文件包中缺少主文档部件。".to_string(),
        ));
    }

    let mut report = CapabilityReport::new();
    for warning in &security.warnings {
        report.warn(warning.clone());
    }

    match kind {
        OfficeKind::Docx => {
            let document = read_text_entry(bytes, "word/document.xml").unwrap_or_default();
            report.feature(CapabilityFeature::new(
                "docx.paragraphs",
                "段落与文字样式",
                true,
                Severity::Info,
            ));
            report.feature(CapabilityFeature::new(
                "docx.tables",
                "表格",
                true,
                Severity::Info,
            ));
            if document.contains("<w:drawing") || document.contains("<w:pict") {
                report.feature(CapabilityFeature::new(
                    "docx.images",
                    "图片与图形",
                    true,
                    Severity::Info,
                ));
            }
            if document.contains("<w:ins ") || document.contains("<w:del ") {
                report.feature(CapabilityFeature::new(
                    "docx.revisions",
                    "修订记录",
                    false,
                    Severity::Warning,
                ));
                report.warn(CompatibilityWarning::new(
                    "docx.revisions",
                    "文档包含修订。V1 可查看但不会编辑修订，建议另存副本。",
                    Severity::Warning,
                ));
            }
            if document.contains("<w:sdt") {
                report.feature(CapabilityFeature::new(
                    "docx.contentControls",
                    "内容控件",
                    false,
                    Severity::Warning,
                ));
                report.warn(CompatibilityWarning::new(
                    "docx.contentControls",
                    "文档包含内容控件，V1 以普通文本方式显示。",
                    Severity::Warning,
                ));
            }
            if any_name(&entries, |name| name.starts_with("word/embeddings/"))
                || document.contains("<w:object")
            {
                report.warn(CompatibilityWarning::new(
                    "docx.embedded",
                    "文档包含嵌入对象，对象会保留但无法编辑。",
                    Severity::Warning,
                ));
            }
            if read_text_entry(bytes, "word/header1.xml")
                .map(|xml| !xml.is_empty())
                .unwrap_or(false)
            {
                report.feature(CapabilityFeature::new(
                    "docx.headers",
                    "页眉页脚",
                    true,
                    Severity::Info,
                ));
            }
            report.feature(CapabilityFeature::new(
                "docx.fields",
                "页码与域",
                true,
                Severity::Info,
            ));
        }
        OfficeKind::Xlsx => {
            if any_name(&entries, |name| name.starts_with("xl/pivotTables/")) {
                report.warn(CompatibilityWarning::new(
                    "xlsx.pivot",
                    "工作簿包含数据透视表，V1 不支持编辑或刷新。",
                    Severity::Warning,
                ));
            }
            if any_name(&entries, |name| name.starts_with("xl/charts/")) {
                report.warn(CompatibilityWarning::new(
                    "xlsx.charts",
                    "工作簿包含图表，保存副本时需要重新核对显示。",
                    Severity::Warning,
                ));
            }
            if any_name(&entries, |name| name.starts_with("xl/externalLinks/")) {
                report.warn(CompatibilityWarning::new(
                    "xlsx.externalLinks",
                    "工作簿包含外部链接，V1 不会刷新外部数据。",
                    Severity::Warning,
                ));
            }
            report.feature(CapabilityFeature::new(
                "xlsx.formulas",
                "公式与跨表引用",
                true,
                Severity::Info,
            ));
            report.feature(CapabilityFeature::new(
                "xlsx.styles",
                "单元格格式",
                true,
                Severity::Info,
            ));
        }
        OfficeKind::Pptx => {
            let slide_names: Vec<String> = entries
                .iter()
                .filter(|name| {
                    name.starts_with("ppt/slides/slide") && name.ends_with(".xml")
                })
                .cloned()
                .collect();
            let mut animations = false;
            let mut media = false;
            for name in &slide_names {
                if let Ok(xml) = read_text_entry(bytes, name) {
                    animations |= xml.contains("<p:timing");
                    media |= xml.contains("<p:video") || xml.contains("<p:audio");
                }
            }
            if animations {
                report.warn(CompatibilityWarning::new(
                    "pptx.animations",
                    "演示文稿包含动画，V1 仅显示静态页面。",
                    Severity::Warning,
                ));
            }
            if media {
                report.warn(CompatibilityWarning::new(
                    "pptx.media",
                    "演示文稿包含视频或音频，V1 不会播放媒体。",
                    Severity::Warning,
                ));
            }
            if any_name(&entries, |name| name.starts_with("ppt/notesSlides/")) {
                report.feature(CapabilityFeature::new(
                    "pptx.notes",
                    "演讲者备注",
                    true,
                    Severity::Info,
                ));
            }
            report.feature(CapabilityFeature::new(
                "pptx.preview",
                "静态高保真预览",
                true,
                Severity::Info,
            ));
            report.feature(CapabilityFeature::new(
                "pptx.editing",
                "基础内容编辑与安全另存",
                true,
                Severity::Info,
            ));
            report.warn(CompatibilityWarning::new(
                "pptx.editing",
                "支持基础内容编辑并另存副本；复杂动画、媒体和图表保留原始部件。",
                Severity::Info,
            ));
        }
    }

    let capability = report.finalize(kind == OfficeKind::Pptx);
    Ok(PackageInspection {
        kind,
        main_part,
        entry_count: entries.len(),
        entries: entries
            .into_iter()
            .map(|name| PackageEntry {
                name,
                size: 0,
                compressed_size: 0,
            })
            .collect(),
        security,
        capability,
    })
}

pub fn validate_package(bytes: &[u8]) -> Result<OfficeKind, OfficeError> {
    validate_zip(bytes)?;
    let kind = detect_kind(bytes, "")?;
    let main = main_part(bytes)?
        .ok_or_else(|| OfficeError::InvalidPackage("缺少主文档部件。".to_string()))?;
    if !names_contain(&entry_names(bytes)?, &main) {
        return Err(OfficeError::InvalidPackage(
            "主文档部件不存在。".to_string(),
        ));
    }
    Ok(kind)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    fn fixture(path: &str) -> Vec<u8> {
        let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
        std::fs::read(root.join(path)).expect("fixture should exist")
    }

    #[test]
    fn detects_and_scans_docx() {
        let bytes = fixture("tests/fixtures/docx/contract.docx");
        let inspection = analyze(&bytes, "contract.docx").unwrap();
        assert_eq!(inspection.kind, OfficeKind::Docx);
        assert!(inspection.capability.level >= 2);
        assert!(inspection.entry_count > 5);
        assert!(read_text_entry(&bytes, "word/document.xml")
            .unwrap()
            .contains("房屋租赁合同"));
    }

    #[test]
    fn detects_and_scans_xlsx() {
        let bytes = fixture("tests/fixtures/xlsx/workbook.xlsx");
        let inspection = analyze(&bytes, "workbook.xlsx").unwrap();
        assert_eq!(inspection.kind, OfficeKind::Xlsx);
        assert!(inspection.entry_count > 5);
    }

    #[test]
    fn pptx_supports_basic_editing() {
        let bytes = fixture("tests/fixtures/pptx/deck.pptx");
        let inspection = analyze(&bytes, "deck.pptx").unwrap();
        assert_eq!(inspection.kind, OfficeKind::Pptx);
        assert_eq!(inspection.capability.level, 1);
        assert!(inspection.capability.features.iter().any(|feature| feature.code == "pptx.editing"));
    }

    #[test]
    fn validation_round_trips_our_output() {
        let bytes = fixture("tests/fixtures/docx/contract.docx");
        assert_eq!(validate_package(&bytes).unwrap(), OfficeKind::Docx);
    }
}

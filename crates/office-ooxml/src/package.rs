use crate::open_archive;
use office_core::OfficeError;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

/// 功能：保存 OOXML 原始部件和精确修改，未修改部件不经过 XML 重新序列化；时间：2026-10-28；作者：lq。
#[derive(Debug, Clone)]
pub struct OoxmlPackage {
    source: Vec<u8>,
    replacements: HashMap<String, Vec<u8>>,
    deleted: HashSet<String>,
}

impl OoxmlPackage {
    /// 功能：从受安全限制的 OOXML 字节创建可增量保存的包；时间：2026-10-28；作者：lq。
    pub fn open(source: impl Into<Vec<u8>>) -> Result<Self, OfficeError> {
        let source = source.into();
        let _ = open_archive(&source)?;
        Ok(Self { source, replacements: HashMap::new(), deleted: HashSet::new() })
    }

    /// 功能：读取部件当前内容，优先返回尚未写回的修改；时间：2026-10-28；作者：lq。
    pub fn read(&self, name: &str) -> Result<Vec<u8>, OfficeError> {
        if self.deleted.contains(name) {
            return Err(OfficeError::InvalidPackage(format!("部件已删除：{name}")));
        }
        if let Some(value) = self.replacements.get(name) {
            return Ok(value.clone());
        }
        crate::read_entry(&self.source, name)
    }

    /// 功能：替换或新增一个 OOXML 部件，不触碰其他原始部件；时间：2026-10-28；作者：lq。
    pub fn replace(&mut self, name: impl Into<String>, bytes: impl Into<Vec<u8>>) {
        let name = name.into();
        self.deleted.remove(&name);
        self.replacements.insert(name, bytes.into());
    }

    /// 功能：标记部件删除，保存时才从输出包移除；时间：2026-10-28；作者：lq。
    pub fn delete(&mut self, name: impl Into<String>) {
        let name = name.into();
        self.replacements.remove(&name);
        self.deleted.insert(name);
    }

    /// 功能：列出原始包中的部件名称，便于能力扫描和未知部件保留；时间：2026-10-28；作者：lq。
    pub fn names(&self) -> Result<Vec<String>, OfficeError> {
        crate::entry_names(&self.source)
    }

    /// 功能：仅在最终保存时重建 ZIP，保留未修改部件的字节内容；时间：2026-10-28；作者：lq。
    pub fn finish(self) -> Result<Vec<u8>, OfficeError> {
        let mut archive = open_archive(&self.source)?;
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        let original_names = crate::entry_names(&self.source)?;
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
            let name = entry.name().to_string();
            if entry.is_dir() || self.deleted.contains(&name) { continue; }
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).map_err(|error| OfficeError::InvalidPackage(error.to_string()))?;
            let bytes = self.replacements.get(&name).map(Vec::as_slice).unwrap_or(&bytes);
            writer.start_file(name, SimpleFileOptions::default().compression_method(CompressionMethod::Deflated)).map_err(|error| OfficeError::Storage(error.to_string()))?;
            writer.write_all(bytes).map_err(|error| OfficeError::Storage(error.to_string()))?;
        }
        for (name, bytes) in self.replacements.iter().filter(|(name, _)| !original_names.iter().any(|item| item == *name)) {
            writer.start_file(name, SimpleFileOptions::default().compression_method(CompressionMethod::Deflated)).map_err(|error| OfficeError::Storage(error.to_string()))?;
            writer.write_all(bytes).map_err(|error| OfficeError::Storage(error.to_string()))?;
        }
        writer.finish().map(|cursor| cursor.into_inner()).map_err(|error| OfficeError::Storage(error.to_string()))
    }
}

/// 功能：把源包中输出包缺失的部件补回，作为 native 保存链路的兼容性兜底；时间：2026-10-28；作者：lq。
pub fn preserve_missing_parts(source: &[u8], output: &[u8]) -> Result<Vec<u8>, OfficeError> {
    let source_package = OoxmlPackage::open(source.to_vec())?;
    let output_names = crate::entry_names(output)?;
    let mut package = OoxmlPackage::open(output.to_vec())?;
    for name in source_package.names()? {
        if !output_names.iter().any(|item| item == &name) {
            package.replace(name.clone(), source_package.read(&name)?);
        }
    }
    // 功能：同步恢复未知部件对应的内容类型声明，避免 Office 将已保留部件视为无效扩展；时间：2026-10-28；作者：lq。
    if let (Ok(source_types), Ok(output_types)) = (
        source_package.read("[Content_Types].xml"),
        package.read("[Content_Types].xml"),
    ) {
        let merged = merge_content_types(&source_types, &output_types)?;
        package.replace("[Content_Types].xml", merged);
    }
    // 功能：合并同名关系部件中源包缺失的 Relationship 记录，恢复未知对象的可达性；时间：2026-10-28；作者：lq。
    for name in source_package.names()? {
        if !name.ends_with(".rels") || !output_names.iter().any(|item| item == &name) {
            continue;
        }
        let Ok(source_relationships) = source_package.read(&name) else { continue; };
        let Ok(output_relationships) = package.read(&name) else { continue; };
        let merged = merge_relationships(&source_relationships, &output_relationships)?;
        package.replace(name, merged);
    }
    package.finish()
}

/// 功能：把源包中缺失的 Default/Override 声明追加到输出包，保持未知格式部件可被宿主识别；时间：2026-10-28；作者：lq。
fn merge_content_types(source: &[u8], output: &[u8]) -> Result<Vec<u8>, OfficeError> {
    let mut source_entries: Vec<(String, Vec<u8>)> = Vec::new();
    let mut source_reader = Reader::from_reader(source);
    source_reader.config_mut().trim_text(true);
    loop {
        match source_reader.read_event() {
            Ok(Event::Empty(element)) | Ok(Event::Start(element))
                if matches!(element.name().local_name().as_ref(), b"Default" | b"Override") =>
            {
                let key: &[u8] = if element.name().local_name().as_ref() == b"Default" {
                    b"Extension"
                } else {
                    b"PartName"
                };
                if let Some(value) = element
                    .attributes()
                    .flatten()
                    .find(|attribute| attribute.key.local_name().as_ref() == key)
                    .map(|attribute| String::from_utf8_lossy(&attribute.value).into_owned())
                {
                    source_entries.push((value, element.to_vec()));
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(OfficeError::InvalidPackage(format!("内容类型解析失败：{error}"))),
            _ => {}
        }
    }

    let mut output_reader = Reader::from_reader(output);
    output_reader.config_mut().trim_text(true);
    let mut existing = std::collections::HashSet::new();
    loop {
        match output_reader.read_event() {
            Ok(Event::Empty(element)) | Ok(Event::Start(element))
                if matches!(element.name().local_name().as_ref(), b"Default" | b"Override") =>
            {
                let key: &[u8] = if element.name().local_name().as_ref() == b"Default" {
                    b"Extension"
                } else {
                    b"PartName"
                };
                if let Some(value) = element
                    .attributes()
                    .flatten()
                    .find(|attribute| attribute.key.local_name().as_ref() == key)
                    .map(|attribute| String::from_utf8_lossy(&attribute.value).into_owned())
                {
                    existing.insert(value);
                }
            }
            Ok(Event::End(element)) if element.name().local_name().as_ref() == b"Types" => break,
            Ok(Event::Eof) => break,
            Err(error) => return Err(OfficeError::InvalidPackage(format!("输出内容类型解析失败：{error}"))),
            _ => {}
        }
    }
    // 功能：避开 quick-xml buffer_position 在不同输入编码下的偏差，稳定定位内容类型根元素闭合标签；时间：2026-10-28；作者：lq。
    let close_start = output
        .iter()
        .rposition(|byte| *byte == b'<')
        .filter(|position| output.get(*position..).is_some_and(|tail| tail.starts_with(b"</")))
        .unwrap_or(output.len());
    if close_start == output.len() {
        return Ok(output.to_vec());
    }
    let mut additions = Vec::new();
    for (key, raw) in source_entries {
        if existing.insert(key) {
            additions.extend_from_slice(&raw);
        }
    }
    if additions.is_empty() {
        return Ok(output.to_vec());
    }
    let mut merged = Vec::with_capacity(output.len() + additions.len());
    merged.extend_from_slice(&output[..close_start]);
    merged.extend_from_slice(&additions);
    merged.extend_from_slice(&output[close_start..]);
    Ok(merged)
}

/// 功能：按 Relationship Id 合并关系记录，保留输出包已有目标并补回源包新增目标；时间：2026-10-28；作者：lq。
fn merge_relationships(source: &[u8], output: &[u8]) -> Result<Vec<u8>, OfficeError> {
    let mut source_entries: Vec<(String, Vec<u8>)> = Vec::new();
    let mut source_reader = Reader::from_reader(source);
    source_reader.config_mut().trim_text(true);
    loop {
        match source_reader.read_event() {
            Ok(Event::Empty(element)) | Ok(Event::Start(element))
                if element.name().local_name().as_ref() == b"Relationship" =>
            {
                if let Some(id) = element
                    .attributes()
                    .flatten()
                    .find(|attribute| attribute.key.local_name().as_ref() == b"Id")
                    .map(|attribute| String::from_utf8_lossy(&attribute.value).into_owned())
                {
                    source_entries.push((id, element.to_vec()));
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(OfficeError::InvalidPackage(format!("关系解析失败：{error}"))),
            _ => {}
        }
    }
    let mut output_reader = Reader::from_reader(output);
    output_reader.config_mut().trim_text(true);
    let mut existing = std::collections::HashSet::new();
    loop {
        match output_reader.read_event() {
            Ok(Event::Empty(element)) | Ok(Event::Start(element))
                if element.name().local_name().as_ref() == b"Relationship" =>
            {
                if let Some(id) = element
                    .attributes()
                    .flatten()
                    .find(|attribute| attribute.key.local_name().as_ref() == b"Id")
                    .map(|attribute| String::from_utf8_lossy(&attribute.value).into_owned())
                {
                    existing.insert(id);
                }
            }
            Ok(Event::End(element)) if element.name().local_name().as_ref() == b"Relationships" => break,
            Ok(Event::Eof) => break,
            Err(error) => return Err(OfficeError::InvalidPackage(format!("输出关系解析失败：{error}"))),
            _ => {}
        }
    }
    let close_start = output
        .iter()
        .rposition(|byte| *byte == b'<')
        .filter(|position| output.get(*position..).is_some_and(|tail| tail.starts_with(b"</")))
        .unwrap_or(output.len());
    if close_start == output.len() {
        return Ok(output.to_vec());
    }
    let mut additions = Vec::new();
    for (id, raw) in source_entries {
        if existing.insert(id) {
            additions.extend_from_slice(&raw);
        }
    }
    if additions.is_empty() {
        return Ok(output.to_vec());
    }
    let mut merged = Vec::with_capacity(output.len() + additions.len());
    merged.extend_from_slice(&output[..close_start]);
    merged.extend_from_slice(&additions);
    merged.extend_from_slice(&output[close_start..]);
    Ok(merged)
}

#[cfg(test)]
mod tests {
    use super::{preserve_missing_parts, OoxmlPackage};
    use std::io::Cursor;
    use std::io::Read;
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn sample() -> Vec<u8> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        writer.start_file("word/document.xml", SimpleFileOptions::default()).unwrap();
        std::io::Write::write_all(&mut writer, b"original").unwrap();
        writer.start_file("custom/data.bin", SimpleFileOptions::default()).unwrap();
        std::io::Write::write_all(&mut writer, b"keep-me").unwrap();
        writer.finish().unwrap().into_inner()
    }

    #[test]
    fn preserves_unknown_parts_when_replacing_one_part() {
        let mut package = OoxmlPackage::open(sample()).unwrap();
        package.replace("word/document.xml", b"changed".to_vec());
        let bytes = package.finish().unwrap();
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut custom = String::new();
        archive.by_name("custom/data.bin").unwrap().read_to_string(&mut custom).unwrap();
        assert_eq!(custom, "keep-me");
    }

    // 功能：验证源包缺失部件可在 native 保存兜底中恢复；时间：2026-10-28；作者：lq。
    #[test]
    fn restores_missing_parts_from_source() {
        let source = sample();
        let mut output = OoxmlPackage::open(source.clone()).unwrap();
        output.delete("custom/data.bin");
        let output = output.finish().unwrap();
        let restored = preserve_missing_parts(&source, &output).unwrap();
        assert!(crate::entry_names(&restored).unwrap().contains(&"custom/data.bin".to_string()));
    }

    // 功能：验证补回未知部件时同步恢复内容类型声明，覆盖真实 OOXML 宿主识别链路；时间：2026-10-28；作者：lq。
    #[test]
    fn restores_missing_content_type_override() {
        let mut source_writer = ZipWriter::new(Cursor::new(Vec::new()));
        source_writer.start_file("[Content_Types].xml", SimpleFileOptions::default()).unwrap();
        std::io::Write::write_all(&mut source_writer, br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/custom/data.bin" ContentType="application/x-custom"/></Types>"#).unwrap();
        source_writer.start_file("custom/data.bin", SimpleFileOptions::default()).unwrap();
        std::io::Write::write_all(&mut source_writer, b"keep").unwrap();
        let source = source_writer.finish().unwrap().into_inner();

        let mut output_writer = ZipWriter::new(Cursor::new(Vec::new()));
        output_writer.start_file("[Content_Types].xml", SimpleFileOptions::default()).unwrap();
        std::io::Write::write_all(&mut output_writer, br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>"#).unwrap();
        let output = output_writer.finish().unwrap().into_inner();
        let restored = preserve_missing_parts(&source, &output).unwrap();
        let types = crate::read_text_entry(&restored, "[Content_Types].xml").unwrap();
        assert!(types.contains("/custom/data.bin"));
        assert!(crate::entry_names(&restored).unwrap().contains(&"custom/data.bin".to_string()));
    }

    // 功能：验证同名关系部件会恢复源包缺失的目标记录，确保未知图片或对象仍可被主文档访问；时间：2026-10-28；作者：lq。
    #[test]
    fn restores_missing_relationship_record() {
        let mut source_writer = ZipWriter::new(Cursor::new(Vec::new()));
        source_writer.start_file("word/_rels/document.xml.rels", SimpleFileOptions::default()).unwrap();
        std::io::Write::write_all(&mut source_writer, br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:custom" Target="custom/data.bin"/></Relationships>"#).unwrap();
        source_writer.start_file("custom/data.bin", SimpleFileOptions::default()).unwrap();
        std::io::Write::write_all(&mut source_writer, b"keep").unwrap();
        let source = source_writer.finish().unwrap().into_inner();

        let mut output_writer = ZipWriter::new(Cursor::new(Vec::new()));
        output_writer.start_file("word/_rels/document.xml.rels", SimpleFileOptions::default()).unwrap();
        std::io::Write::write_all(&mut output_writer, br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#).unwrap();
        let output = output_writer.finish().unwrap().into_inner();
        let restored = preserve_missing_parts(&source, &output).unwrap();
        let relationships = crate::read_text_entry(&restored, "word/_rels/document.xml.rels").unwrap();
        assert!(relationships.contains("rId1"));
        assert!(relationships.contains("custom/data.bin"));
    }
}

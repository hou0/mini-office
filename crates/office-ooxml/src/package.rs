use crate::open_archive;
use office_core::OfficeError;
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

#[cfg(test)]
mod tests {
    use super::OoxmlPackage;
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
}

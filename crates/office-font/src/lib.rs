use office_core::OfficeError;
use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FontInfo {
    pub family: String,
    pub style: String,
    pub path: String,
    pub has_chinese: bool,
}

const CJK_FONT_HINTS: &[&str] = &[
    "songti",
    "heiti",
    "kaiti",
    "fangsong",
    "pingfang",
    "hiragino",
    "simsun",
    "simhei",
    "kaiti",
    "yahei",
    "dengxian",
    "source han",
    "noto sans cjk",
    "noto serif cjk",
    "wenquanyi",
    "微软雅黑",
    "宋体",
    "黑体",
    "楷体",
    "仿宋",
    "等线",
    "思源",
];

pub fn font_directories() -> Vec<PathBuf> {
    let mut dirs: Vec<PathBuf> = Vec::new();
    #[cfg(target_os = "macos")]
    {
        dirs.push(PathBuf::from("/System/Library/Fonts"));
        dirs.push(PathBuf::from("/System/Library/Fonts/Supplemental"));
        dirs.push(PathBuf::from("/Library/Fonts"));
        if let Some(home) = dirs::home_dir() {
            dirs.push(home.join("Library/Fonts"));
        }
    }
    #[cfg(target_os = "windows")]
    {
        if let Some(windows) = std::env::var_os("WINDIR") {
            dirs.push(PathBuf::from(windows).join("Fonts"));
        }
        if let Some(local) = dirs::data_local_dir() {
            dirs.push(local.join("Microsoft/Windows/Fonts"));
        }
    }
    #[cfg(target_os = "linux")]
    {
        dirs.push(PathBuf::from("/usr/share/fonts"));
        dirs.push(PathBuf::from("/usr/local/share/fonts"));
        if let Some(home) = dirs::home_dir() {
            dirs.push(home.join(".fonts"));
            dirs.push(home.join(".local/share/fonts"));
        }
    }
    dirs
}

fn has_chinese_name(family: &str) -> bool {
    family.chars().any(|character| {
        matches!(character, '\u{3400}'..='\u{9fff}' | '\u{f900}'..='\u{faff}')
    }) || {
        let lower = family.to_ascii_lowercase();
        CJK_FONT_HINTS
            .iter()
            .any(|hint| lower.contains(&hint.to_ascii_lowercase()))
    }
}

fn face_names(data: &[u8], index: u32) -> Option<(String, String)> {
    let face = ttf_parser::Face::parse(data, index).ok()?;
    let mut family: Option<String> = None;
    let mut style: Option<String> = None;
    for name in face.names() {
        if name.name_id == ttf_parser::name_id::FAMILY && family.is_none() {
            family = name.to_string();
        }
        if name.name_id == ttf_parser::name_id::SUBFAMILY && style.is_none() {
            style = name.to_string();
        }
    }
    Some((
        family.unwrap_or_else(|| "Unknown".to_string()),
        style.unwrap_or_else(|| "Regular".to_string()),
    ))
}

fn parse_font_file(path: &Path, output: &mut Vec<FontInfo>, seen: &mut HashSet<String>) {
    let data = match fs::read(path) {
        Ok(data) => data,
        Err(_) => return,
    };
    let count = ttf_parser::fonts_in_collection(&data).unwrap_or(1);
    for index in 0..count {
        if let Some((family, style)) = face_names(&data, index) {
            let key = format!("{}::{}", family.to_ascii_lowercase(), style.to_ascii_lowercase());
            if !seen.insert(key) {
                continue;
            }
            output.push(FontInfo {
                has_chinese: has_chinese_name(&family),
                family,
                style,
                path: path.to_string_lossy().to_string(),
            });
        }
    }
}

fn walk(directory: &Path, depth: usize, output: &mut Vec<FontInfo>, seen: &mut HashSet<String>) {
    if depth > 2 || output.len() > 4_000 {
        return;
    }
    let entries = match fs::read_dir(directory) {
        Ok(entries) => entries,
        Err(_) => return,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            walk(&path, depth + 1, output, seen);
            continue;
        }
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        if matches!(extension.as_str(), "ttf" | "otf" | "ttc" | "otc") {
            parse_font_file(&path, output, seen);
        }
    }
}

pub fn list_fonts() -> Vec<FontInfo> {
    let mut output = Vec::new();
    let mut seen = HashSet::new();
    for directory in font_directories() {
        walk(&directory, 0, &mut output, &mut seen);
    }
    output.sort_by(|a, b| {
        b.has_chinese
            .cmp(&a.has_chinese)
            .then_with(|| a.family.to_lowercase().cmp(&b.family.to_lowercase()))
    });
    output
}

pub fn cached_fonts() -> &'static [FontInfo] {
    static FONTS: OnceLock<Vec<FontInfo>> = OnceLock::new();
    FONTS.get_or_init(list_fonts)
}

pub fn font_families() -> Vec<String> {
    let mut families = Vec::new();
    let mut seen = HashSet::new();
    for font in list_fonts() {
        let key = font.family.to_ascii_lowercase();
        if seen.insert(key) {
            families.push(font.family);
        }
    }
    families
}

pub fn require_font(family: &str) -> Result<FontInfo, OfficeError> {
    let lower = family.to_ascii_lowercase();
    list_fonts()
        .into_iter()
        .find(|font| font.family.to_ascii_lowercase() == lower)
        .ok_or_else(|| OfficeError::Other(format!("缺少字体：{family}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scans_system_font_directories() {
        let fonts = list_fonts();
        assert!(!fonts.is_empty(), "system should provide at least one font");
        // Chinese coverage depends on the host's installed fonts.
    }

    #[test]
    fn directories_are_platform_specific() {
        assert!(!font_directories().is_empty());
    }
}

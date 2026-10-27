use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::BTreeMap;

/// 功能：保存 DrawingML 主题色和主题字体，供 DOCX/PPTX 共用；时间：2026-10-28；作者：lq。
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ThemeInfo {
    pub colors: BTreeMap<String, String>,
    pub major_font: Option<String>,
    pub minor_font: Option<String>,
}

fn attr<'a>(event: &'a quick_xml::events::BytesStart<'a>, name: &[u8]) -> Option<String> {
    event
        .attributes()
        .flatten()
        .find(|item| item.key.local_name().as_ref() == name)
        .map(|item| String::from_utf8_lossy(&item.value).into_owned())
}

/// 功能：读取主题色、主题字体和常用别名，避免不同格式各自实现主题继承；时间：2026-10-28；作者：lq。
pub fn parse_theme(xml: &str) -> ThemeInfo {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut info = ThemeInfo::default();
    let mut color_name: Option<String> = None;
    let mut font_group: Option<String> = None;
    let mut font_kind: Option<String> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) | Ok(Event::Empty(event)) => {
                let name = event.local_name();
                match name.as_ref() {
                    b"dk1" | b"lt1" | b"dk2" | b"lt2" | b"accent1" | b"accent2" | b"accent3" | b"accent4" | b"accent5" | b"accent6" | b"hlink" | b"folHlink" => {
                        color_name = Some(String::from_utf8_lossy(name.as_ref()).into_owned());
                    }
                    b"srgbClr" | b"sysClr" if color_name.is_some() => {
                        if let Some(value) = attr(&event, if name.as_ref() == b"sysClr" { b"lastClr" } else { b"val" }) {
                            info.colors.insert(color_name.take().unwrap(), format!("#{}", value.to_uppercase()));
                        }
                    }
                    b"majorFont" => font_group = Some("major".to_string()),
                    b"minorFont" => font_group = Some("minor".to_string()),
                    b"latin" | b"ea" if font_group.is_some() => font_kind = Some(String::from_utf8_lossy(name.as_ref()).into_owned()),
                    _ => {}
                }
                if let Some(group) = font_group.as_deref() {
                    if font_kind.is_some() {
                        if let Some(value) = attr(&event, b"typeface") {
                            if !value.is_empty() && !value.starts_with('+') {
                                if group == "major" { info.major_font = Some(value); } else { info.minor_font = Some(value); }
                                font_kind = None;
                            }
                        }
                    }
                }
            }
            Ok(Event::End(event)) => match event.local_name().as_ref() {
                b"majorFont" | b"minorFont" => { font_group = None; font_kind = None; }
                b"latin" | b"ea" => font_kind = None,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    if let Some(value) = info.colors.get("dk1").cloned() { info.colors.insert("tx1".to_string(), value); }
    if let Some(value) = info.colors.get("lt1").cloned() { info.colors.insert("bg1".to_string(), value); }
    if let Some(value) = info.colors.get("dk2").cloned() { info.colors.insert("tx2".to_string(), value); }
    if let Some(value) = info.colors.get("lt2").cloned() { info.colors.insert("bg2".to_string(), value); }
    info
}

#[cfg(test)]
mod tests {
    use super::parse_theme;

    #[test]
    fn parses_theme_colors_fonts_and_aliases() {
        let theme = parse_theme(r#"<a:theme xmlns:a="x"><a:themeElements><a:clrScheme><a:dk1><a:srgbClr val="112233"/></a:dk1><a:accent1><a:srgbClr val="AaBbCc"/></a:accent1></a:clrScheme><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#);
        assert_eq!(theme.colors.get("tx1"), Some(&"#112233".to_string()));
        assert_eq!(theme.colors.get("accent1"), Some(&"#AABBCC".to_string()));
        assert_eq!(theme.major_font.as_deref(), Some("Aptos Display"));
        assert_eq!(theme.minor_font.as_deref(), Some("Aptos"));
    }
}

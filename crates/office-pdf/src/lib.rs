use office_core::OfficeError;
use serde::{Deserialize, Serialize};
use std::path::{Path, PathBuf};
use std::process::Command;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PdfExportResult {
    pub input: String,
    pub output: String,
    pub converter: String,
}

#[derive(Debug, Clone)]
pub struct LibreOfficeConverter {
    executable: PathBuf,
}

impl LibreOfficeConverter {
    pub fn discover() -> Option<Self> {
        for candidate in candidates() {
            if candidate.is_file() {
                return Some(Self {
                    executable: candidate,
                });
            }
        }
        for name in ["soffice", "libreoffice"] {
            if let Ok(path) = which(name) {
                return Some(Self { executable: path });
            }
        }
        None
    }

    pub fn executable(&self) -> &Path {
        &self.executable
    }

    /// Convert a document with a headless LibreOffice process. This is used as
    /// a PoC comparison backend; the shipping frontend renders PDF itself.
    pub fn convert_to_pdf(
        &self,
        input: impl AsRef<Path>,
        output_directory: impl AsRef<Path>,
    ) -> Result<PdfExportResult, OfficeError> {
        let input = input.as_ref();
        let output_directory = output_directory.as_ref();
        if !input.is_file() {
            return Err(OfficeError::Pdf(format!(
                "输入文件不存在：{}",
                input.display()
            )));
        }
        std::fs::create_dir_all(output_directory)?;
        let output = Command::new(&self.executable)
            .arg("--headless")
            .arg("--norestore")
            .arg("--convert-to")
            .arg("pdf")
            .arg("--outdir")
            .arg(output_directory)
            .arg(input)
            .output()
            .map_err(|error| OfficeError::Pdf(format!("无法启动 LibreOffice：{error}")))?;
        if !output.status.success() {
            return Err(OfficeError::Pdf(format!(
                "LibreOffice 转换失败：{}",
                String::from_utf8_lossy(&output.stderr)
            )));
        }
        let file_name = input
            .file_stem()
            .and_then(|value| value.to_str())
            .unwrap_or("document");
        let pdf = output_directory.join(format!("{file_name}.pdf"));
        if !pdf.is_file() {
            return Err(OfficeError::Pdf(
                "LibreOffice 未生成 PDF 文件。".to_string(),
            ));
        }
        Ok(PdfExportResult {
            input: input.to_string_lossy().to_string(),
            output: pdf.to_string_lossy().to_string(),
            converter: self.executable.to_string_lossy().to_string(),
        })
    }
}

fn candidates() -> Vec<PathBuf> {
    let mut result = Vec::new();
    #[cfg(target_os = "macos")]
    {
        result.push(PathBuf::from(
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ));
        result.push(PathBuf::from(
            "/Applications/LibreOffice.app/Contents/MacOS/libreoffice",
        ));
    }
    #[cfg(target_os = "windows")]
    {
        if let Some(program_files) = std::env::var_os("ProgramFiles") {
            result.push(
                PathBuf::from(program_files)
                    .join("LibreOffice")
                    .join("program")
                    .join("soffice.exe"),
            );
        }
    }
    #[cfg(target_os = "linux")]
    {
        result.push(PathBuf::from("/usr/bin/soffice"));
        result.push(PathBuf::from("/usr/bin/libreoffice"));
        result.push(PathBuf::from("/snap/bin/libreoffice"));
    }
    result
}

fn which(name: &str) -> Result<PathBuf, OfficeError> {
    let output = Command::new("which")
        .arg(name)
        .output()
        .map_err(|error| OfficeError::Pdf(error.to_string()))?;
    if !output.status.success() {
        return Err(OfficeError::Pdf(format!("找不到可执行文件：{name}")));
    }
    let path = String::from_utf8_lossy(&output.stdout).trim().to_string();
    if path.is_empty() {
        return Err(OfficeError::Pdf(format!("找不到可执行文件：{name}")));
    }
    Ok(PathBuf::from(path))
}

pub fn is_available() -> bool {
    LibreOfficeConverter::discover().is_some()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn discovery_does_not_panic() {
        let converter = LibreOfficeConverter::discover();
        if let Some(converter) = converter {
            assert!(converter.executable().exists());
        }
    }

    #[test]
    fn missing_input_is_reported() {
        if let Some(converter) = LibreOfficeConverter::discover() {
            let directory = tempfile::tempdir().unwrap();
            let result = converter.convert_to_pdf(
                directory.path().join("missing.docx"),
                directory.path(),
            );
            assert!(result.is_err());
        }
    }
}

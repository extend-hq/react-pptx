mod legacy;
mod model;
mod pptx;

pub use model::*;

use thiserror::Error;

#[derive(Debug, Error)]
pub enum ParseError {
    #[error("input is empty")]
    Empty,
    #[error("input exceeds the configured {0} byte limit")]
    InputLimit(usize),
    #[error("unsupported presentation format")]
    UnsupportedFormat,
    #[error("encrypted presentations are not supported")]
    Encrypted,
    #[error("corrupt presentation: {0}")]
    Corrupt(String),
    #[error("resource limit exceeded: {0}")]
    ResourceLimit(String),
}

#[derive(Debug, Clone)]
pub struct ParseLimits {
    pub max_input_bytes: usize,
    pub max_entries: usize,
    pub max_entry_bytes: usize,
    pub max_total_uncompressed_bytes: usize,
    pub max_xml_depth: usize,
    pub max_xml_nodes: usize,
    pub max_xml_attributes: usize,
    pub max_xml_text_bytes: usize,
    pub max_xml_attribute_bytes: usize,
}

impl Default for ParseLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: 100 * 1024 * 1024,
            max_entries: 20_000,
            max_entry_bytes: 64 * 1024 * 1024,
            max_total_uncompressed_bytes: 512 * 1024 * 1024,
            max_xml_depth: 256,
            max_xml_nodes: 250_000,
            max_xml_attributes: 1_000_000,
            max_xml_text_bytes: 64 * 1024 * 1024,
            max_xml_attribute_bytes: 32 * 1024 * 1024,
        }
    }
}

pub fn detect_format(bytes: &[u8]) -> Result<PresentationFormat, ParseError> {
    if bytes.is_empty() {
        return Err(ParseError::Empty);
    }
    if bytes.starts_with(b"PK\x03\x04")
        || bytes.starts_with(b"PK\x05\x06")
        || bytes.starts_with(b"PK\x07\x08")
    {
        return Ok(PresentationFormat::Pptx);
    }
    const OLE_MAGIC: &[u8; 8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1";
    if bytes.starts_with(OLE_MAGIC) {
        return Ok(PresentationFormat::Ppt);
    }
    Err(ParseError::UnsupportedFormat)
}

pub fn parse_presentation(
    bytes: &[u8],
    limits: &ParseLimits,
) -> Result<PresentationDocument, ParseError> {
    if bytes.len() > limits.max_input_bytes {
        return Err(ParseError::InputLimit(limits.max_input_bytes));
    }
    match detect_format(bytes)? {
        PresentationFormat::Pptx => pptx::parse(bytes, limits),
        PresentationFormat::Ppt => legacy::parse(bytes, limits),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detects_ooxml_and_ole_magic() {
        assert_eq!(
            detect_format(b"PK\x03\x04rest").unwrap(),
            PresentationFormat::Pptx
        );
        assert_eq!(
            detect_format(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1rest").unwrap(),
            PresentationFormat::Ppt
        );
    }

    #[test]
    fn rejects_unknown_input() {
        assert!(matches!(
            detect_format(b"not powerpoint"),
            Err(ParseError::UnsupportedFormat)
        ));
    }
}

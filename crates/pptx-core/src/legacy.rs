use std::{
    collections::{BTreeMap, BTreeSet},
    io::{Cursor, Read},
};

use flate2::read::ZlibDecoder;

use crate::{
    ColorValue, FillStyle, LineStyle, ParseError, ParseLimits, PresentationAsset,
    PresentationDocument, PresentationFormat, PresentationSize, PresentationSlide,
    PresentationWarning, ShapeGeometry, SlideNode, TextParagraph, TextRun, Transform,
    VerticalAlignment,
};

const DEFAULT_WIDTH: u64 = 9_144_000;
const DEFAULT_HEIGHT: u64 = 6_858_000;
const MASTER_UNIT_EMU_NUMERATOR: i64 = 3_175;
const MASTER_UNIT_EMU_DENOMINATOR: i64 = 2;

const RT_DOCUMENT: u16 = 1000;
const RT_DOCUMENT_ATOM: u16 = 1001;
const RT_SLIDE: u16 = 1006;
const RT_SLIDE_ATOM: u16 = 1007;
const RT_MAIN_MASTER: u16 = 1016;
const RT_SLIDE_PERSIST_ATOM: u16 = 1011;
const RT_ENVIRONMENT: u16 = 1010;
const RT_PLACEHOLDER_ATOM: u16 = 3011;
const RT_OUTLINE_TEXT_REF_ATOM: u16 = 3998;
const RT_TEXT_HEADER_ATOM: u16 = 3999;
const RT_TEXT_CHARS_ATOM: u16 = 4000;
const RT_TEXT_BYTES_ATOM: u16 = 4008;
const RT_FONT_ENTITY_ATOM: u16 = 4023;
const RT_SLIDE_LIST_WITH_TEXT: u16 = 4080;
const RT_USER_EDIT_ATOM: u16 = 4085;
const RT_CURRENT_USER_ATOM: u16 = 4086;
const RT_PERSIST_PTR_FULL_BLOCK: u16 = 6001;
const RT_PERSIST_PTR_INCREMENTAL_BLOCK: u16 = 6002;

const OFFICE_ART_BSTORE_CONTAINER: u16 = 0xf001;
const OFFICE_ART_SPGR_CONTAINER: u16 = 0xf003;
const OFFICE_ART_SP_CONTAINER: u16 = 0xf004;
const OFFICE_ART_FBSE: u16 = 0xf007;
const OFFICE_ART_FSPGR: u16 = 0xf009;
const OFFICE_ART_FSP: u16 = 0xf00a;
const OFFICE_ART_FOPT: u16 = 0xf00b;
const OFFICE_ART_CLIENT_TEXTBOX: u16 = 0xf00d;
const OFFICE_ART_CHILD_ANCHOR: u16 = 0xf00f;
const OFFICE_ART_CLIENT_ANCHOR: u16 = 0xf010;
const OFFICE_ART_CLIENT_DATA: u16 = 0xf011;
const OFFICE_ART_BLIP_EMF: u16 = 0xf01a;
const OFFICE_ART_BLIP_WMF: u16 = 0xf01b;
const OFFICE_ART_BLIP_PICT: u16 = 0xf01c;
const OFFICE_ART_BLIP_JPEG: u16 = 0xf01d;
const OFFICE_ART_BLIP_PNG: u16 = 0xf01e;
const OFFICE_ART_BLIP_DIB: u16 = 0xf01f;
const OFFICE_ART_BLIP_TIFF: u16 = 0xf029;

const FOPT_PIB: u16 = 0x0104;
const FOPT_ANCHOR_TEXT: u16 = 0x0087;
const FOPT_FILL_TYPE: u16 = 0x0180;
const FOPT_FILL_COLOR: u16 = 0x0181;
const FOPT_FILL_BACK_COLOR: u16 = 0x0183;
const FOPT_FILL_STYLE_BOOLEAN: u16 = 0x01bf;
const FOPT_LINE_COLOR: u16 = 0x01c0;
const FOPT_LINE_WIDTH: u16 = 0x01cb;
const FOPT_LINE_STYLE_BOOLEAN: u16 = 0x01ff;

#[derive(Clone, Copy, Debug)]
struct Record<'a> {
    version: u8,
    instance: u16,
    record_type: u16,
    payload: &'a [u8],
}

fn read_u16(data: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        data.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn read_i16(data: &[u8], offset: usize) -> Option<i16> {
    Some(i16::from_le_bytes(
        data.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn read_u32(data: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        data.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

fn read_i32(data: &[u8], offset: usize) -> Option<i32> {
    Some(i32::from_le_bytes(
        data.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

fn record_at(data: &[u8], offset: usize) -> Option<Record<'_>> {
    let header = data.get(offset..offset.checked_add(8)?)?;
    let version_instance = u16::from_le_bytes([header[0], header[1]]);
    let length = u32::from_le_bytes([header[4], header[5], header[6], header[7]]) as usize;
    let payload_start = offset.checked_add(8)?;
    let payload_end = payload_start.checked_add(length)?;
    Some(Record {
        version: (version_instance & 0x000f) as u8,
        instance: version_instance >> 4,
        record_type: u16::from_le_bytes([header[2], header[3]]),
        payload: data.get(payload_start..payload_end)?,
    })
}

fn child_records(data: &[u8]) -> Vec<Record<'_>> {
    let mut result = Vec::new();
    let mut offset = 0usize;
    while offset + 8 <= data.len() {
        let Some(record) = record_at(data, offset) else {
            break;
        };
        let size = 8 + record.payload.len();
        result.push(record);
        let Some(next) = offset.checked_add(size) else {
            break;
        };
        offset = next;
    }
    result
}

fn walk_records<'a>(data: &'a [u8], depth: usize, max_depth: usize, output: &mut Vec<Record<'a>>) {
    if depth > max_depth {
        return;
    }
    for record in child_records(data) {
        output.push(record);
        if record.version == 0x0f {
            walk_records(record.payload, depth + 1, max_depth, output);
        }
    }
}

fn read_stream<F: std::io::Read + std::io::Seek>(
    compound: &mut cfb::CompoundFile<F>,
    name: &str,
    limit: usize,
) -> Result<Vec<u8>, ParseError> {
    let mut stream = compound
        .open_stream(name)
        .map_err(|_| ParseError::Corrupt(format!("missing {name} stream")))?;
    let declared_len = stream.len() as usize;
    if declared_len > limit {
        return Err(ParseError::ResourceLimit(format!("{name} stream size")));
    }
    let mut data = Vec::with_capacity(declared_len);
    stream
        .read_to_end(&mut data)
        .map_err(|error| ParseError::Corrupt(error.to_string()))?;
    Ok(data)
}

#[derive(Debug)]
struct UserEdit {
    previous_offset: u32,
    persist_directory_offset: u32,
    document_persist_id: u32,
}

fn current_edit_offset(current_user: &[u8]) -> Result<u32, ParseError> {
    let record = record_at(current_user, 0)
        .filter(|record| record.record_type == RT_CURRENT_USER_ATOM)
        .ok_or_else(|| ParseError::Corrupt("invalid Current User stream".into()))?;
    if record.payload.len() < 20 {
        return Err(ParseError::Corrupt("truncated CurrentUserAtom".into()));
    }
    match read_u32(record.payload, 4) {
        Some(0xf3d1_c4df) => return Err(ParseError::Encrypted),
        Some(0xe391_c05f) => {}
        Some(_) | None => {
            return Err(ParseError::Corrupt(
                "invalid CurrentUserAtom header token".into(),
            ));
        }
    }
    read_u32(record.payload, 8)
        .ok_or_else(|| ParseError::Corrupt("missing current edit offset".into()))
}

fn user_edit_at(data: &[u8], offset: u32) -> Result<UserEdit, ParseError> {
    let record = record_at(data, offset as usize)
        .filter(|record| record.record_type == RT_USER_EDIT_ATOM)
        .ok_or_else(|| ParseError::Corrupt(format!("invalid UserEditAtom at {offset:#x}")))?;
    if record.payload.len() != 28 && record.payload.len() != 32 {
        return Err(ParseError::Corrupt(format!(
            "invalid UserEditAtom length {}",
            record.payload.len()
        )));
    }
    if record.payload.get(7).copied() != Some(3) {
        return Err(ParseError::Corrupt(
            "unsupported UserEditAtom major version".into(),
        ));
    }
    if record.payload.len() == 32 && read_u32(record.payload, 28).unwrap_or(0) != 0 {
        return Err(ParseError::Encrypted);
    }
    Ok(UserEdit {
        previous_offset: read_u32(record.payload, 8).unwrap_or(0),
        persist_directory_offset: read_u32(record.payload, 12)
            .ok_or_else(|| ParseError::Corrupt("missing persist directory offset".into()))?,
        document_persist_id: read_u32(record.payload, 16)
            .ok_or_else(|| ParseError::Corrupt("missing document persist id".into()))?,
    })
}

fn parse_persist_directory(data: &[u8], offset: u32) -> Result<Vec<(u32, u32)>, ParseError> {
    let record = record_at(data, offset as usize)
        .filter(|record| {
            matches!(
                record.record_type,
                RT_PERSIST_PTR_FULL_BLOCK | RT_PERSIST_PTR_INCREMENTAL_BLOCK
            )
        })
        .ok_or_else(|| ParseError::Corrupt(format!("invalid persist directory at {offset:#x}")))?;
    let mut entries = Vec::new();
    let mut position = 0usize;
    while position < record.payload.len() {
        let descriptor = read_u32(record.payload, position)
            .ok_or_else(|| ParseError::Corrupt("truncated persist directory entry".into()))?;
        position += 4;
        let first_id = descriptor & 0x000f_ffff;
        let count = descriptor >> 20;
        if first_id == 0 || count == 0 {
            return Err(ParseError::Corrupt(
                "invalid persist directory entry range".into(),
            ));
        }
        for index in 0..count {
            let object_offset = read_u32(record.payload, position)
                .ok_or_else(|| ParseError::Corrupt("truncated persist offset array".into()))?;
            position += 4;
            if object_offset as usize >= data.len() {
                return Err(ParseError::Corrupt(format!(
                    "persist object offset {object_offset:#x} is out of bounds"
                )));
            }
            entries.push((first_id + index, object_offset));
        }
    }
    Ok(entries)
}

fn live_persist_directory(
    data: &[u8],
    current_user: &[u8],
    limits: &ParseLimits,
) -> Result<(BTreeMap<u32, u32>, u32), ParseError> {
    let mut edit_offset = current_edit_offset(current_user)?;
    let mut visited = BTreeSet::new();
    let mut edits = Vec::new();
    while edit_offset != 0 {
        if edits.len() >= limits.max_xml_depth.max(1) {
            return Err(ParseError::ResourceLimit("legacy edit-chain depth".into()));
        }
        if !visited.insert(edit_offset) {
            return Err(ParseError::Corrupt("cyclic legacy edit chain".into()));
        }
        let edit = user_edit_at(data, edit_offset)?;
        if edit.persist_directory_offset >= edit_offset {
            return Err(ParseError::Corrupt(
                "persist directory does not precede its user edit".into(),
            ));
        }
        edit_offset = edit.previous_offset;
        edits.push(edit);
    }
    let document_persist_id = edits
        .first()
        .map(|edit| edit.document_persist_id)
        .ok_or_else(|| ParseError::Corrupt("empty legacy edit chain".into()))?;
    let mut directory = BTreeMap::new();
    for edit in edits.iter().rev() {
        for (persist_id, offset) in parse_persist_directory(data, edit.persist_directory_offset)? {
            directory.insert(persist_id, offset);
        }
    }
    Ok((directory, document_persist_id))
}

fn master_units_to_emu(value: i64) -> i64 {
    value
        .saturating_mul(MASTER_UNIT_EMU_NUMERATOR)
        .div_euclid(MASTER_UNIT_EMU_DENOMINATOR)
}

fn document_size(document: Record<'_>) -> PresentationSize {
    let atom = child_records(document.payload)
        .into_iter()
        .find(|record| record.record_type == RT_DOCUMENT_ATOM);
    let (width, height) = atom
        .and_then(|record| Some((read_i32(record.payload, 0)?, read_i32(record.payload, 4)?)))
        .filter(|(width, height)| *width > 0 && *height > 0)
        .map(|(width, height)| {
            (
                master_units_to_emu(width as i64) as u64,
                master_units_to_emu(height as i64) as u64,
            )
        })
        .unwrap_or((DEFAULT_WIDTH, DEFAULT_HEIGHT));
    PresentationSize {
        width_emu: width,
        height_emu: height,
    }
}

fn decode_utf16(payload: &[u8]) -> String {
    let units = payload
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect::<Vec<_>>();
    String::from_utf16_lossy(&units)
        .trim_matches(char::from(0))
        .to_owned()
}

fn cp1252_character(byte: u8) -> char {
    const REPLACEMENTS: [char; 32] = [
        '\u{20ac}', '\u{0081}', '\u{201a}', '\u{0192}', '\u{201e}', '\u{2026}', '\u{2020}',
        '\u{2021}', '\u{02c6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008d}',
        '\u{017d}', '\u{008f}', '\u{0090}', '\u{2018}', '\u{2019}', '\u{201c}', '\u{201d}',
        '\u{2022}', '\u{2013}', '\u{2014}', '\u{02dc}', '\u{2122}', '\u{0161}', '\u{203a}',
        '\u{0153}', '\u{009d}', '\u{017e}', '\u{0178}',
    ];
    if (0x80..=0x9f).contains(&byte) {
        REPLACEMENTS[(byte - 0x80) as usize]
    } else {
        byte as char
    }
}

fn decode_text(record: Record<'_>) -> Option<String> {
    let text = match record.record_type {
        RT_TEXT_CHARS_ATOM => decode_utf16(record.payload),
        RT_TEXT_BYTES_ATOM => record
            .payload
            .iter()
            .copied()
            .map(cp1252_character)
            .collect::<String>()
            .trim_matches(char::from(0))
            .to_owned(),
        _ => return None,
    };
    (!text.trim().is_empty()).then_some(text)
}

#[derive(Clone, Debug)]
struct TextBody {
    text: String,
    text_type: Option<u32>,
}

fn text_bodies(data: &[u8]) -> Vec<TextBody> {
    let mut bodies = Vec::new();
    let mut text_type = None;
    for record in child_records(data) {
        match record.record_type {
            RT_TEXT_HEADER_ATOM => text_type = read_u32(record.payload, 0),
            RT_TEXT_CHARS_ATOM | RT_TEXT_BYTES_ATOM => {
                if let Some(text) = decode_text(record) {
                    bodies.push(TextBody { text, text_type });
                }
            }
            _ if record.version == 0x0f => bodies.extend(text_bodies(record.payload)),
            _ => {}
        }
    }
    bodies
}

fn default_font_family(document: Record<'_>, max_depth: usize) -> Option<String> {
    let environment = child_records(document.payload)
        .into_iter()
        .find(|record| record.record_type == RT_ENVIRONMENT)?;
    let mut records = Vec::new();
    walk_records(environment.payload, 0, max_depth, &mut records);
    let fonts = records
        .into_iter()
        .filter(|record| record.record_type == RT_FONT_ENTITY_ATOM)
        .map(|record| decode_utf16(&record.payload[..record.payload.len().min(64)]))
        .filter(|font| !font.is_empty())
        .collect::<Vec<_>>();
    fonts
        .iter()
        .find(|font| font.eq_ignore_ascii_case("Arial"))
        .cloned()
        .or_else(|| fonts.into_iter().next())
}

fn text_size_for_type(text_type: Option<u32>) -> Option<f64> {
    match text_type {
        Some(0 | 6) => Some(36.0),
        Some(1 | 5 | 7 | 8) => Some(15.0),
        Some(_) => Some(18.0),
        None => None,
    }
}

fn paragraphs(body: &TextBody, font_family: Option<&str>) -> Vec<TextParagraph> {
    body.text
        .split(['\r', '\n', '\u{000b}'])
        .map(|text| TextParagraph {
            runs: vec![TextRun {
                text: text.to_owned(),
                font_family: font_family.map(str::to_owned),
                font_size_pt: text_size_for_type(body.text_type),
                bold: None,
                italic: None,
                ..Default::default()
            }],
            ..Default::default()
        })
        .collect()
}

#[derive(Clone, Copy, Debug)]
struct Rect {
    left: f64,
    top: f64,
    right: f64,
    bottom: f64,
}

impl Rect {
    fn width(self) -> f64 {
        self.right - self.left
    }

    fn height(self) -> f64 {
        self.bottom - self.top
    }
}

#[derive(Clone, Copy)]
struct GroupTransform {
    source: Rect,
    target: Rect,
}

impl GroupTransform {
    fn map(self, rect: Rect) -> Rect {
        let scale_x = if self.source.width().abs() < f64::EPSILON {
            1.0
        } else {
            self.target.width() / self.source.width()
        };
        let scale_y = if self.source.height().abs() < f64::EPSILON {
            1.0
        } else {
            self.target.height() / self.source.height()
        };
        Rect {
            left: self.target.left + (rect.left - self.source.left) * scale_x,
            top: self.target.top + (rect.top - self.source.top) * scale_y,
            right: self.target.left + (rect.right - self.source.left) * scale_x,
            bottom: self.target.top + (rect.bottom - self.source.top) * scale_y,
        }
    }
}

fn client_anchor(payload: &[u8]) -> Option<Rect> {
    match payload.len() {
        8 => Some(Rect {
            top: read_i16(payload, 0)? as f64,
            left: read_i16(payload, 2)? as f64,
            right: read_i16(payload, 4)? as f64,
            bottom: read_i16(payload, 6)? as f64,
        }),
        16 => Some(Rect {
            top: read_i32(payload, 0)? as f64,
            left: read_i32(payload, 4)? as f64,
            right: read_i32(payload, 8)? as f64,
            bottom: read_i32(payload, 12)? as f64,
        }),
        _ => None,
    }
}

fn coordinate_rect(payload: &[u8]) -> Option<Rect> {
    Some(Rect {
        left: read_i32(payload, 0)? as f64,
        top: read_i32(payload, 4)? as f64,
        right: read_i32(payload, 8)? as f64,
        bottom: read_i32(payload, 12)? as f64,
    })
}

fn rect_transform(rect: Rect) -> Transform {
    let left = master_units_to_emu(rect.left.round() as i64);
    let top = master_units_to_emu(rect.top.round() as i64);
    let right = master_units_to_emu(rect.right.round() as i64);
    let bottom = master_units_to_emu(rect.bottom.round() as i64);
    Transform {
        x: left,
        y: top,
        width: right.saturating_sub(left).max(0),
        height: bottom.saturating_sub(top).max(0),
        ..Default::default()
    }
}

fn fopt_property(record: Record<'_>, property_id: u16) -> Option<u32> {
    if record.record_type != OFFICE_ART_FOPT {
        return None;
    }
    for index in 0..record.instance as usize {
        let offset = index.checked_mul(6)?;
        let opid = read_u16(record.payload, offset)?;
        if opid & 0x3fff == property_id {
            return read_u32(record.payload, offset + 2);
        }
    }
    None
}

fn direct_color(value: u32) -> Option<ColorValue> {
    if value >> 24 != 0 {
        return None;
    }
    Some(ColorValue {
        value: format!(
            "#{:02X}{:02X}{:02X}",
            value & 0xff,
            (value >> 8) & 0xff,
            (value >> 16) & 0xff
        ),
        alpha: None,
    })
}

fn fopt_style(
    record: Record<'_>,
) -> (
    Option<FillStyle>,
    Option<LineStyle>,
    Option<VerticalAlignment>,
) {
    let fill_boolean = fopt_property(record, FOPT_FILL_STYLE_BOOLEAN);
    let explicitly_filled = fill_boolean
        .filter(|value| value & 0x0010_0000 != 0)
        .map(|value| value & 0x10 != 0);
    let fill = if explicitly_filled == Some(false) {
        Some(FillStyle::None)
    } else {
        let fill_type = fopt_property(record, FOPT_FILL_TYPE);
        let foreground = fopt_property(record, FOPT_FILL_COLOR).and_then(direct_color);
        let background = fopt_property(record, FOPT_FILL_BACK_COLOR).and_then(direct_color);
        match fill_type {
            Some(1) => Some(FillStyle::Pattern {
                preset: "legacy-pattern".into(),
                foreground: foreground.unwrap_or_else(|| ColorValue {
                    value: "#FFFFFF".into(),
                    alpha: None,
                }),
                background: background.unwrap_or_else(|| ColorValue {
                    value: "#FFFFFF".into(),
                    alpha: None,
                }),
            }),
            Some(0) | None if foreground.is_some() || explicitly_filled == Some(true) => {
                Some(FillStyle::Solid {
                    color: foreground.unwrap_or_else(|| ColorValue {
                        value: "#FFFFFF".into(),
                        alpha: None,
                    }),
                })
            }
            _ => None,
        }
    };

    let line_boolean = fopt_property(record, FOPT_LINE_STYLE_BOOLEAN);
    let explicitly_lined = line_boolean
        .filter(|value| value & 0x0008_0000 != 0)
        .map(|value| value & 0x8 != 0);
    let line_color = fopt_property(record, FOPT_LINE_COLOR).and_then(direct_color);
    let line_width = fopt_property(record, FOPT_LINE_WIDTH).map(|width| width as f64 / 9_525.0);
    let line = if explicitly_lined == Some(false) {
        None
    } else if line_color.is_some() || line_width.is_some() || explicitly_lined == Some(true) {
        Some(LineStyle {
            color: line_color,
            width: line_width.or(Some(1.0)),
            ..Default::default()
        })
    } else {
        None
    };

    let vertical_alignment = match fopt_property(record, FOPT_ANCHOR_TEXT) {
        Some(0) => Some(VerticalAlignment::Top),
        Some(1) => Some(VerticalAlignment::Middle),
        Some(2) => Some(VerticalAlignment::Bottom),
        _ => None,
    };
    (fill, line, vertical_alignment)
}

fn placeholder(record: Record<'_>) -> bool {
    if record.record_type != OFFICE_ART_CLIENT_DATA {
        return false;
    }
    let mut records = Vec::new();
    walk_records(record.payload, 0, 8, &mut records);
    records.into_iter().any(|record| {
        record.record_type == RT_PLACEHOLDER_ATOM
            && read_u32(record.payload, 0).is_some_and(|position| position != u32::MAX)
    })
}

fn textbox_body(record: Record<'_>, outline: &[TextBody]) -> Option<TextBody> {
    if record.record_type != OFFICE_ART_CLIENT_TEXTBOX {
        return None;
    }
    let records = child_records(record.payload);
    if let Some(reference) = records
        .iter()
        .find(|record| record.record_type == RT_OUTLINE_TEXT_REF_ATOM)
        .and_then(|record| read_u32(record.payload, 0))
    {
        return outline.get(reference as usize).cloned();
    }
    text_bodies(record.payload).into_iter().next()
}

#[derive(Default)]
struct ShapeParts {
    spid: u32,
    shape_type: u16,
    anchor: Option<Rect>,
    child_anchor: bool,
    group_coordinates: Option<Rect>,
    text: Option<TextBody>,
    picture_index: Option<u32>,
    placeholder: bool,
    fill: Option<FillStyle>,
    line: Option<LineStyle>,
    vertical_alignment: Option<VerticalAlignment>,
}

fn shape_parts(record: Record<'_>, outline: &[TextBody]) -> ShapeParts {
    let mut result = ShapeParts::default();
    for child in child_records(record.payload) {
        match child.record_type {
            OFFICE_ART_FSPGR => result.group_coordinates = coordinate_rect(child.payload),
            OFFICE_ART_FSP => {
                result.spid = read_u32(child.payload, 0).unwrap_or(0);
                result.shape_type = child.instance;
            }
            OFFICE_ART_FOPT => {
                result.picture_index = result
                    .picture_index
                    .or_else(|| fopt_property(child, FOPT_PIB));
                let (fill, line, vertical_alignment) = fopt_style(child);
                if fill.is_some() {
                    result.fill = fill;
                }
                if line.is_some() {
                    result.line = line;
                }
                if vertical_alignment.is_some() {
                    result.vertical_alignment = vertical_alignment;
                }
            }
            OFFICE_ART_CLIENT_ANCHOR => result.anchor = client_anchor(child.payload),
            OFFICE_ART_CHILD_ANCHOR => {
                result.anchor = coordinate_rect(child.payload);
                result.child_anchor = true;
            }
            OFFICE_ART_CLIENT_DATA => result.placeholder |= placeholder(child),
            OFFICE_ART_CLIENT_TEXTBOX => result.text = textbox_body(child, outline),
            _ => {}
        }
    }
    result
}

fn shape_preset(shape_type: u16) -> &'static str {
    match shape_type {
        3 => "ellipse",
        20 | 32..=40 => "line",
        _ => "rect",
    }
}

struct ShapeCollector<'a> {
    prefix: &'a str,
    font_family: Option<&'a str>,
    outline: &'a [TextBody],
    picture_assets: &'a BTreeMap<u32, String>,
    include_placeholders: bool,
    slide_size: &'a PresentationSize,
    nodes: Vec<SlideNode>,
}

impl ShapeCollector<'_> {
    fn push_shape(&mut self, parts: ShapeParts, group: Option<GroupTransform>) {
        if parts.placeholder && !self.include_placeholders {
            return;
        }
        let Some(mut anchor) = parts.anchor else {
            return;
        };
        if parts.child_anchor {
            let Some(group) = group else {
                return;
            };
            anchor = group.map(anchor);
        }
        let id = format!("{}-{}", self.prefix, parts.spid);
        let transform = rect_transform(anchor);
        if let Some(asset_id) = parts
            .picture_index
            .and_then(|index| self.picture_assets.get(&index))
        {
            let covers_slide = transform.x.abs() <= 1
                && transform.y.abs() <= 1
                && transform.width as u64 >= self.slide_size.width_emu.saturating_mul(95) / 100
                && transform.height as u64 >= self.slide_size.height_emu.saturating_mul(95) / 100;
            self.nodes.push(SlideNode::Image {
                id,
                name: format!("Legacy picture {}", parts.spid),
                transform,
                asset_id: asset_id.clone(),
                crop: None,
                opacity: None,
                preserve_aspect_ratio: !covers_slide,
            });
            return;
        }
        let paragraphs = parts
            .text
            .as_ref()
            .map(|body| paragraphs(body, self.font_family))
            .unwrap_or_default();
        self.nodes.push(SlideNode::Shape {
            id,
            name: format!("Legacy shape {}", parts.spid),
            transform,
            geometry: ShapeGeometry {
                preset: Some(shape_preset(parts.shape_type).into()),
                path: None,
            },
            fill: parts.fill,
            line: parts.line,
            paragraphs,
            vertical_alignment: parts.vertical_alignment,
            text_insets: None,
            autofit: None,
            text_rotation: None,
            vertical_text: None,
            horizontal_overflow: None,
            vertical_overflow: None,
            text_wrap: None,
            column_count: None,
            column_spacing: None,
            right_to_left_columns: None,
            space_first_last_paragraph: None,
        });
    }

    fn collect_group(&mut self, record: Record<'_>, parent: Option<GroupTransform>) {
        let children = child_records(record.payload);
        let Some(group_shape) = children
            .first()
            .copied()
            .filter(|child| child.record_type == OFFICE_ART_SP_CONTAINER)
        else {
            return;
        };
        let group_parts = shape_parts(group_shape, self.outline);
        let source = group_parts.group_coordinates;
        let mut target = group_parts.anchor;
        if group_parts.child_anchor {
            target = target.and_then(|rect| parent.map(|parent| parent.map(rect)));
        }
        let group = source
            .zip(target)
            .map(|(source, target)| GroupTransform { source, target });
        let active_group = group.or(parent).or_else(|| {
            source.map(|source| GroupTransform {
                source,
                target: source,
            })
        });
        for child in children.into_iter().skip(1) {
            match child.record_type {
                OFFICE_ART_SP_CONTAINER => {
                    self.push_shape(shape_parts(child, self.outline), active_group)
                }
                OFFICE_ART_SPGR_CONTAINER => self.collect_group(child, active_group),
                _ => {}
            }
        }
    }

    fn find_groups(&mut self, data: &[u8], depth: usize, max_depth: usize) {
        if depth > max_depth {
            return;
        }
        for record in child_records(data) {
            if record.record_type == OFFICE_ART_SPGR_CONTAINER {
                self.collect_group(record, None);
            } else if record.version == 0x0f {
                self.find_groups(record.payload, depth + 1, max_depth);
            }
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn slide_shapes(
    slide: Record<'_>,
    prefix: &str,
    font_family: Option<&str>,
    outline: &[TextBody],
    picture_assets: &BTreeMap<u32, String>,
    include_placeholders: bool,
    slide_size: &PresentationSize,
    max_depth: usize,
) -> Vec<SlideNode> {
    let mut collector = ShapeCollector {
        prefix,
        font_family,
        outline,
        picture_assets,
        include_placeholders,
        slide_size,
        nodes: Vec::new(),
    };
    collector.find_groups(slide.payload, 0, max_depth);
    collector.nodes
}

#[derive(Default)]
struct SlideReference {
    persist_id: u32,
    slide_id: u32,
    outline: Vec<TextBody>,
}

fn slide_references(document: Record<'_>) -> Vec<SlideReference> {
    let Some(slide_list) = child_records(document.payload)
        .into_iter()
        .find(|record| record.record_type == RT_SLIDE_LIST_WITH_TEXT && record.instance == 0)
    else {
        return Vec::new();
    };
    let mut slides = Vec::new();
    let mut current_text_type = None;
    for record in child_records(slide_list.payload) {
        match record.record_type {
            RT_SLIDE_PERSIST_ATOM => {
                slides.push(SlideReference {
                    persist_id: read_u32(record.payload, 0).unwrap_or(0),
                    slide_id: read_u32(record.payload, 12).unwrap_or(0),
                    outline: Vec::new(),
                });
                current_text_type = None;
            }
            RT_TEXT_HEADER_ATOM => current_text_type = read_u32(record.payload, 0),
            RT_TEXT_CHARS_ATOM | RT_TEXT_BYTES_ATOM => {
                if let (Some(slide), Some(text)) = (slides.last_mut(), decode_text(record)) {
                    slide.outline.push(TextBody {
                        text,
                        text_type: current_text_type,
                    });
                }
            }
            _ => {}
        }
    }
    slides
}

fn master_references(document: Record<'_>) -> BTreeMap<u32, u32> {
    let Some(master_list) = child_records(document.payload)
        .into_iter()
        .find(|record| record.record_type == RT_SLIDE_LIST_WITH_TEXT && record.instance == 1)
    else {
        return BTreeMap::new();
    };
    child_records(master_list.payload)
        .into_iter()
        .filter(|record| record.record_type == RT_SLIDE_PERSIST_ATOM)
        .filter_map(|record| Some((read_u32(record.payload, 12)?, read_u32(record.payload, 0)?)))
        .collect()
}

fn slide_master_id(slide: Record<'_>) -> Option<u32> {
    let atom = child_records(slide.payload)
        .into_iter()
        .find(|record| record.record_type == RT_SLIDE_ATOM)?;
    read_u32(atom.payload, 12)
}

fn collect_fbse_offsets(document: Record<'_>, max_depth: usize) -> Vec<u32> {
    fn find_store(data: &[u8], depth: usize, max_depth: usize) -> Option<Vec<u32>> {
        if depth > max_depth {
            return None;
        }
        for record in child_records(data) {
            if record.record_type == OFFICE_ART_BSTORE_CONTAINER {
                return Some(
                    child_records(record.payload)
                        .into_iter()
                        .filter(|child| child.record_type == OFFICE_ART_FBSE)
                        .filter_map(|child| read_u32(child.payload, 28))
                        .collect(),
                );
            }
            if record.version == 0x0f {
                if let Some(result) = find_store(record.payload, depth + 1, max_depth) {
                    return Some(result);
                }
            }
        }
        None
    }
    find_store(document.payload, 0, max_depth).unwrap_or_default()
}

fn find_signature(data: &[u8], signature: &[u8]) -> Option<usize> {
    data.windows(signature.len())
        .position(|window| window == signature)
}

fn decode_metafile(record: Record<'_>, max_bytes: usize) -> Option<Vec<u8>> {
    let mut header_offset = 16usize;
    if record.payload.len() < header_offset + 34 {
        return None;
    }
    let uncompressed_size = read_u32(record.payload, header_offset)? as usize;
    if uncompressed_size > max_bytes {
        return None;
    }
    let saved_size = read_u32(record.payload, header_offset + 28)? as usize;
    let compression = *record.payload.get(header_offset + 32)?;
    if saved_size > record.payload.len().saturating_sub(header_offset + 34) {
        header_offset = 32;
    }
    let compressed = record.payload.get(header_offset + 34..)?;
    match compression {
        0x00 => {
            let mut output = Vec::with_capacity(uncompressed_size);
            ZlibDecoder::new(compressed)
                .take(max_bytes.saturating_add(1) as u64)
                .read_to_end(&mut output)
                .ok()?;
            (output.len() == uncompressed_size && output.len() <= max_bytes).then_some(output)
        }
        0xfe if compressed.len() <= max_bytes => Some(compressed.to_vec()),
        _ => None,
    }
}

fn picture_data(
    record: Record<'_>,
    max_bytes: usize,
) -> Option<(&'static str, Vec<u8>, &'static str)> {
    match record.record_type {
        OFFICE_ART_BLIP_PNG => {
            let offset = find_signature(record.payload, b"\x89PNG\r\n\x1a\n")?;
            Some(("image/png", record.payload[offset..].to_vec(), "png"))
        }
        OFFICE_ART_BLIP_JPEG => {
            let offset = find_signature(record.payload, b"\xff\xd8\xff")?;
            Some(("image/jpeg", record.payload[offset..].to_vec(), "jpg"))
        }
        OFFICE_ART_BLIP_TIFF => {
            let offset = find_signature(record.payload, b"II*\0")
                .or_else(|| find_signature(record.payload, b"MM\0*"))?;
            Some(("image/tiff", record.payload[offset..].to_vec(), "tiff"))
        }
        OFFICE_ART_BLIP_EMF => Some(("image/x-emf", decode_metafile(record, max_bytes)?, "emf")),
        OFFICE_ART_BLIP_WMF => Some(("image/x-wmf", decode_metafile(record, max_bytes)?, "wmf")),
        OFFICE_ART_BLIP_PICT => Some(("image/x-pict", decode_metafile(record, max_bytes)?, "pict")),
        OFFICE_ART_BLIP_DIB => Some(("image/bmp", record.payload.to_vec(), "dib")),
        _ => None,
    }
}

fn picture_assets(
    pictures: Option<&[u8]>,
    offsets: &[u32],
    limits: &ParseLimits,
) -> (BTreeMap<String, PresentationAsset>, BTreeMap<u32, String>) {
    let mut assets = BTreeMap::new();
    let mut references = BTreeMap::new();
    let Some(pictures) = pictures else {
        return (assets, references);
    };
    for (index, offset) in offsets.iter().copied().enumerate() {
        let Some(record) = record_at(pictures, offset as usize) else {
            continue;
        };
        let Some((content_type, data, extension)) = picture_data(record, limits.max_entry_bytes)
        else {
            continue;
        };
        if data.len() > limits.max_entry_bytes {
            continue;
        }
        let asset_id = format!("legacy-picture-{}", index + 1);
        assets.insert(
            asset_id.clone(),
            PresentationAsset {
                id: asset_id.clone(),
                content_type: content_type.into(),
                byte_length: data.len(),
                file_name: Some(format!("picture-{}.{}", index + 1, extension)),
                data: Some(serde_bytes::ByteBuf::from(data)),
            },
        );
        references.insert(index as u32 + 1, asset_id);
    }
    (assets, references)
}

fn normalize_fallback_slide(
    index: usize,
    chunks: Vec<TextBody>,
    font_family: Option<&str>,
) -> PresentationSlide {
    let nodes = chunks
        .into_iter()
        .enumerate()
        .map(|(node_index, body)| SlideNode::Shape {
            id: format!("legacy-{index}-{node_index}"),
            name: format!("Legacy text {}", node_index + 1),
            transform: Transform {
                x: 457_200,
                y: 457_200 + node_index as i64 * 548_640,
                width: 8_229_600,
                height: 457_200,
                ..Default::default()
            },
            geometry: ShapeGeometry {
                preset: Some("rect".into()),
                path: None,
            },
            fill: None,
            line: None,
            paragraphs: paragraphs(&body, font_family),
            vertical_alignment: None,
            text_insets: None,
            autofit: None,
            text_rotation: None,
            vertical_text: None,
            horizontal_overflow: None,
            vertical_overflow: None,
            text_wrap: None,
            column_count: None,
            column_spacing: None,
            right_to_left_columns: None,
            space_first_last_paragraph: None,
        })
        .collect();
    PresentationSlide {
        id: format!("slide-{}", index + 1),
        index,
        name: None,
        hidden: None,
        master_id: None,
        layout_id: None,
        background: None,
        nodes,
        notes: Vec::new(),
        comments: Vec::new(),
        source_part: Some("/PowerPoint Document".into()),
        warnings: Vec::new(),
    }
}

fn fallback_document(data: &[u8], limits: &ParseLimits) -> PresentationDocument {
    let mut text_slides = child_records(data)
        .into_iter()
        .filter(|record| record.record_type == RT_SLIDE)
        .map(|record| text_bodies(record.payload))
        .collect::<Vec<_>>();
    if text_slides.is_empty() {
        text_slides.push(text_bodies(data));
    }
    if text_slides.is_empty() {
        text_slides.push(Vec::new());
    }
    let slides = text_slides
        .into_iter()
        .take(limits.max_entries)
        .enumerate()
        .map(|(index, chunks)| normalize_fallback_slide(index, chunks, None))
        .collect();
    PresentationDocument {
        format: PresentationFormat::Ppt,
        size: PresentationSize {
            width_emu: DEFAULT_WIDTH,
            height_emu: DEFAULT_HEIGHT,
        },
        slides,
        masters: Vec::new(),
        layouts: Vec::new(),
        themes: Vec::new(),
        assets: BTreeMap::new(),
        embedded_fonts: Vec::new(),
        warnings: vec![PresentationWarning::warning(
            "degraded-rendering",
            "This legacy .ppt has no Current User stream; slide text was recovered with the bounded sequential fallback parser.",
        )],
        metadata: None,
    }
}

fn parse_live_document(
    data: &[u8],
    current_user: &[u8],
    pictures: Option<&[u8]>,
    limits: &ParseLimits,
) -> Result<PresentationDocument, ParseError> {
    let (directory, document_persist_id) = live_persist_directory(data, current_user, limits)?;
    let document_offset = directory
        .get(&document_persist_id)
        .copied()
        .ok_or_else(|| {
            ParseError::Corrupt(format!(
                "document persist id {document_persist_id} is missing"
            ))
        })?;
    let document = record_at(data, document_offset as usize)
        .filter(|record| record.record_type == RT_DOCUMENT)
        .ok_or_else(|| {
            ParseError::Corrupt("document persist object is not a DocumentContainer".into())
        })?;
    let size = document_size(document);
    let font_family = default_font_family(document, limits.max_xml_depth);
    let fbse_offsets = collect_fbse_offsets(document, limits.max_xml_depth);
    let (assets, picture_references) = picture_assets(pictures, &fbse_offsets, limits);
    let masters = master_references(document);
    let references = slide_references(document);
    if references.len() > limits.max_entries {
        return Err(ParseError::ResourceLimit("legacy slide count".into()));
    }
    let mut slides = Vec::with_capacity(references.len());
    let mut warnings = vec![PresentationWarning::warning(
        "degraded-rendering",
        "Legacy .ppt parsing follows the live edit and persist-object graph, preserves ordered text, OfficeArt anchors, and supported pictures; advanced effects and editable chart/table semantics can render with static fallbacks.",
    )];
    for (index, reference) in references.into_iter().enumerate() {
        let slide_offset = directory
            .get(&reference.persist_id)
            .copied()
            .ok_or_else(|| {
                ParseError::Corrupt(format!(
                    "slide persist id {} is missing",
                    reference.persist_id
                ))
            })?;
        let slide = record_at(data, slide_offset as usize)
            .filter(|record| record.record_type == RT_SLIDE)
            .ok_or_else(|| {
                ParseError::Corrupt(format!(
                    "persist id {} is not a SlideContainer",
                    reference.persist_id
                ))
            })?;
        let mut nodes = Vec::new();
        if let Some(master_id) = slide_master_id(slide) {
            if let Some(master_persist_id) = masters.get(&master_id) {
                if let Some(master) = directory
                    .get(master_persist_id)
                    .and_then(|offset| record_at(data, *offset as usize))
                    .filter(|record| matches!(record.record_type, RT_MAIN_MASTER | RT_SLIDE))
                {
                    nodes.extend(slide_shapes(
                        master,
                        &format!("legacy-master-{index}"),
                        font_family.as_deref(),
                        &[],
                        &picture_references,
                        false,
                        &size,
                        limits.max_xml_depth,
                    ));
                }
            }
        }
        nodes.extend(slide_shapes(
            slide,
            &format!("legacy-slide-{index}"),
            font_family.as_deref(),
            &reference.outline,
            &picture_references,
            true,
            &size,
            limits.max_xml_depth,
        ));
        if nodes.is_empty() && !reference.outline.is_empty() {
            nodes =
                normalize_fallback_slide(index, reference.outline, font_family.as_deref()).nodes;
        }
        if nodes.is_empty() {
            let mut warning = PresentationWarning::warning(
                "degraded-rendering",
                "No renderable OfficeArt shapes were recovered for this slide.",
            );
            warning.slide_index = Some(index);
            warnings.push(warning.clone());
        }
        slides.push(PresentationSlide {
            id: if reference.slide_id == 0 {
                format!("slide-{}", index + 1)
            } else {
                format!("slide-{}", reference.slide_id)
            },
            index,
            name: None,
            hidden: None,
            master_id: None,
            layout_id: None,
            background: None,
            nodes,
            notes: Vec::new(),
            comments: Vec::new(),
            source_part: Some(format!("/PowerPoint Document#{slide_offset:#x}")),
            warnings: Vec::new(),
        });
    }
    Ok(PresentationDocument {
        format: PresentationFormat::Ppt,
        size,
        slides,
        masters: Vec::new(),
        layouts: Vec::new(),
        themes: Vec::new(),
        assets,
        embedded_fonts: Vec::new(),
        warnings,
        metadata: None,
    })
}

pub fn parse(bytes: &[u8], limits: &ParseLimits) -> Result<PresentationDocument, ParseError> {
    let cursor = Cursor::new(bytes.to_vec());
    let mut compound = cfb::CompoundFile::open(cursor)
        .map_err(|error| ParseError::Corrupt(format!("OLE: {error}")))?;
    if compound.exists("/EncryptedSummary") {
        return Err(ParseError::Encrypted);
    }
    let entries = compound
        .walk()
        .map(|entry| (entry.is_stream(), entry.len()))
        .collect::<Vec<_>>();
    if entries.len() > limits.max_entries {
        return Err(ParseError::ResourceLimit("OLE entry count".into()));
    }
    let mut total_stream_bytes = 0usize;
    for (is_stream, length) in entries {
        if !is_stream {
            continue;
        }
        let length = usize::try_from(length)
            .map_err(|_| ParseError::ResourceLimit("OLE stream size".into()))?;
        if length > limits.max_entry_bytes {
            return Err(ParseError::ResourceLimit("OLE stream size".into()));
        }
        total_stream_bytes = total_stream_bytes
            .checked_add(length)
            .ok_or_else(|| ParseError::ResourceLimit("OLE total stream size".into()))?;
    }
    if total_stream_bytes > limits.max_total_uncompressed_bytes {
        return Err(ParseError::ResourceLimit("OLE total stream size".into()));
    }
    let data = read_stream(
        &mut compound,
        "/PowerPoint Document",
        limits.max_total_uncompressed_bytes,
    )?;
    if !compound.exists("/Current User") {
        return Ok(fallback_document(&data, limits));
    }
    let current_user = read_stream(&mut compound, "/Current User", limits.max_entry_bytes)?;
    let pictures = if compound.exists("/Pictures") {
        Some(read_stream(
            &mut compound,
            "/Pictures",
            limits.max_total_uncompressed_bytes,
        )?)
    } else {
        None
    };
    parse_live_document(&data, &current_user, pictures.as_deref(), limits)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn atom(record_type: u16, payload: &[u8]) -> Vec<u8> {
        let mut record = Vec::new();
        record.extend_from_slice(&0u16.to_le_bytes());
        record.extend_from_slice(&record_type.to_le_bytes());
        record.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        record.extend_from_slice(payload);
        record
    }

    #[test]
    fn parses_persist_directory_ranges() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&((2u32 << 20) | 5).to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&8u32.to_le_bytes());
        let data = atom(RT_PERSIST_PTR_INCREMENTAL_BLOCK, &payload);
        assert_eq!(
            parse_persist_directory(&data, 0).unwrap(),
            vec![(5, 0), (6, 8)]
        );
    }

    #[test]
    fn maps_group_child_coordinates_to_the_slide() {
        let transform = GroupTransform {
            source: Rect {
                left: 0.0,
                top: 0.0,
                right: 100.0,
                bottom: 100.0,
            },
            target: Rect {
                left: 100.0,
                top: 200.0,
                right: 300.0,
                bottom: 600.0,
            },
        };
        let mapped = transform.map(Rect {
            left: 25.0,
            top: 25.0,
            right: 75.0,
            bottom: 75.0,
        });
        assert_eq!(mapped.left, 150.0);
        assert_eq!(mapped.top, 300.0);
        assert_eq!(mapped.right, 250.0);
        assert_eq!(mapped.bottom, 500.0);
    }

    #[test]
    fn decodes_windows_1252_text_atoms() {
        let record = atom(RT_TEXT_BYTES_ATOM, b"\x93Legacy\x94");
        assert_eq!(
            decode_text(record_at(&record, 0).unwrap()).unwrap(),
            "“Legacy”"
        );
    }
}

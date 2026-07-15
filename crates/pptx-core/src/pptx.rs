use std::{
    cell::RefCell,
    collections::{BTreeMap, HashMap},
    io::{Cursor, Read},
};

use quick_xml::{
    events::{BytesStart, Event},
    Reader,
};
use zip::ZipArchive;

use crate::{
    ChartSeries, ColorValue, FillStyle, GradientStop, LineStyle, ParseError, ParseLimits,
    PresentationAsset, PresentationDocument, PresentationEmbeddedFont, PresentationFormat,
    PresentationMetadata, PresentationSize, PresentationSlide, PresentationWarning, ShapeGeometry,
    SlideComment, SlideNode, SlideNote, TableCell, TextParagraph, TextRun, Transform,
    VerticalAlignment,
};

const DEFAULT_WIDTH: u64 = 9_144_000;
const DEFAULT_HEIGHT: u64 = 6_858_000;

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn attr(start: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    start
        .attributes()
        .with_checks(false)
        .flatten()
        .find_map(|value| {
            (local_name(value.key.as_ref()) == name)
                .then(|| String::from_utf8_lossy(value.value.as_ref()).into_owned())
        })
}

fn normalize_target(base: &str, target: &str) -> String {
    if target.contains("://")
        || target.starts_with("mailto:")
        || target.starts_with("tel:")
        || target.starts_with('#')
    {
        return target.to_owned();
    }
    let mut parts: Vec<&str> = base.split('/').collect();
    parts.pop();
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            value => parts.push(value),
        }
    }
    parts.join("/")
}

fn read_entry(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    name: &str,
    limits: &ParseLimits,
) -> Result<Option<Vec<u8>>, ParseError> {
    let Ok(mut entry) = archive.by_name(name) else {
        return Ok(None);
    };
    if entry.size() as usize > limits.max_entry_bytes {
        return Err(ParseError::ResourceLimit(format!(
            "entry {name} exceeds size limit"
        )));
    }
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry
        .read_to_end(&mut bytes)
        .map_err(|error| ParseError::Corrupt(error.to_string()))?;
    Ok(Some(bytes))
}

fn relationships(
    xml: &[u8],
    base: &str,
    max_depth: usize,
) -> Result<HashMap<String, String>, ParseError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut result = HashMap::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                depth += 1;
                if depth > max_depth {
                    return Err(ParseError::ResourceLimit("XML nesting depth".into()));
                }
                if local_name(start.name().as_ref()) == b"Relationship" {
                    if let (Some(id), Some(target)) = (attr(&start, b"Id"), attr(&start, b"Target"))
                    {
                        result.insert(id, normalize_target(base, &target));
                    }
                }
            }
            Ok(Event::Empty(start)) => {
                if local_name(start.name().as_ref()) == b"Relationship" {
                    if let (Some(id), Some(target)) = (attr(&start, b"Id"), attr(&start, b"Target"))
                    {
                        result.insert(id, normalize_target(base, &target));
                    }
                }
            }
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) => break,
            Err(error) => return Err(ParseError::Corrupt(format!("relationship XML: {error}"))),
            _ => {}
        }
    }
    Ok(result)
}

fn parse_presentation_xml(
    xml: &[u8],
    rels: &HashMap<String, String>,
    limits: &ParseLimits,
) -> Result<(PresentationSize, Vec<String>), ParseError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut size = PresentationSize {
        width_emu: DEFAULT_WIDTH,
        height_emu: DEFAULT_HEIGHT,
    };
    let mut slides = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                depth += 1;
                if depth > limits.max_xml_depth {
                    return Err(ParseError::ResourceLimit("XML nesting depth".into()));
                }
                match local_name(start.name().as_ref()) {
                    b"sldSz" => {
                        size.width_emu = attr(&start, b"cx")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(DEFAULT_WIDTH);
                        size.height_emu = attr(&start, b"cy")
                            .and_then(|v| v.parse().ok())
                            .unwrap_or(DEFAULT_HEIGHT);
                    }
                    b"sldId" => {
                        if let Some(id) =
                            attr(&start, b"id")
                                .filter(|value| value.starts_with("rId"))
                                .or_else(|| {
                                    start.attributes().with_checks(false).flatten().find_map(
                                        |value| {
                                            (local_name(value.key.as_ref()) == b"id"
                                                && String::from_utf8_lossy(value.value.as_ref())
                                                    .starts_with("rId"))
                                            .then(|| {
                                                String::from_utf8_lossy(value.value.as_ref())
                                                    .into_owned()
                                            })
                                        },
                                    )
                                })
                        {
                            if let Some(target) = rels.get(&id) {
                                slides.push(target.clone());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(start)) => match local_name(start.name().as_ref()) {
                b"sldSz" => {
                    size.width_emu = attr(&start, b"cx")
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(DEFAULT_WIDTH);
                    size.height_emu = attr(&start, b"cy")
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(DEFAULT_HEIGHT);
                }
                b"sldId" => {
                    if let Some(id) = attr(&start, b"id")
                        .filter(|value| value.starts_with("rId"))
                        .or_else(|| {
                            start
                                .attributes()
                                .with_checks(false)
                                .flatten()
                                .find_map(|value| {
                                    (local_name(value.key.as_ref()) == b"id"
                                        && String::from_utf8_lossy(value.value.as_ref())
                                            .starts_with("rId"))
                                    .then(|| {
                                        String::from_utf8_lossy(value.value.as_ref()).into_owned()
                                    })
                                })
                        })
                    {
                        if let Some(target) = rels.get(&id) {
                            slides.push(target.clone());
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(_)) => depth = depth.saturating_sub(1),
            Ok(Event::Eof) => break,
            Err(error) => return Err(ParseError::Corrupt(format!("presentation XML: {error}"))),
            _ => {}
        }
    }
    Ok((size, slides))
}

#[derive(Debug)]
struct EmbeddedFontReference {
    family: String,
    variants: Vec<(String, String, String)>,
}

fn parse_embedded_font_references(xml: &[u8]) -> Vec<EmbeddedFontReference> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut current: Option<EmbeddedFontReference> = None;
    let mut fonts = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => match local_name(start.name().as_ref()) {
                b"embeddedFont" => {
                    current = Some(EmbeddedFontReference {
                        family: String::new(),
                        variants: Vec::new(),
                    });
                }
                b"font" => {
                    if let Some(font) = current.as_mut() {
                        font.family = attr(&start, b"typeface").unwrap_or_default();
                    }
                }
                b"regular" | b"bold" | b"italic" | b"boldItalic" => {
                    if let (Some(font), Some(rel_id)) = (current.as_mut(), attr(&start, b"id")) {
                        let (style, weight) = match local_name(start.name().as_ref()) {
                            b"bold" => ("normal", "700"),
                            b"italic" => ("italic", "400"),
                            b"boldItalic" => ("italic", "700"),
                            _ => ("normal", "400"),
                        };
                        font.variants
                            .push((rel_id, style.to_owned(), weight.to_owned()));
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(start)) => match local_name(start.name().as_ref()) {
                b"font" => {
                    if let Some(font) = current.as_mut() {
                        font.family = attr(&start, b"typeface").unwrap_or_default();
                    }
                }
                b"regular" | b"bold" | b"italic" | b"boldItalic" => {
                    if let (Some(font), Some(rel_id)) = (current.as_mut(), attr(&start, b"id")) {
                        let (style, weight) = match local_name(start.name().as_ref()) {
                            b"bold" => ("normal", "700"),
                            b"italic" => ("italic", "400"),
                            b"boldItalic" => ("italic", "700"),
                            _ => ("normal", "400"),
                        };
                        font.variants
                            .push((rel_id, style.to_owned(), weight.to_owned()));
                    }
                }
                _ => {}
            },
            Ok(Event::End(end)) if local_name(end.name().as_ref()) == b"embeddedFont" => {
                if let Some(font) = current.take().filter(|font| !font.family.is_empty()) {
                    fonts.push(font);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    fonts
}

fn font_magic(data: &[u8]) -> Option<&'static str> {
    match data.get(..4) {
        Some(b"OTTO") => Some("font/otf"),
        Some(b"\x00\x01\x00\x00") | Some(b"true") | Some(b"typ1") => Some("font/ttf"),
        Some(b"wOFF") => Some("font/woff"),
        Some(b"wOF2") => Some("font/woff2"),
        _ => None,
    }
}

fn deobfuscate_embedded_font(path: &str, data: &mut [u8]) -> bool {
    if font_magic(data).is_some() {
        return true;
    }
    let stem = path
        .rsplit('/')
        .next()
        .unwrap_or_default()
        .split('.')
        .next()
        .unwrap_or_default();
    let hex: String = stem
        .chars()
        .filter(|value| value.is_ascii_hexdigit())
        .collect();
    if hex.len() != 32 || data.len() < 32 {
        return false;
    }
    let mut key = Vec::with_capacity(16);
    for index in (0..32).step_by(2) {
        let Ok(value) = u8::from_str_radix(&hex[index..index + 2], 16) else {
            return false;
        };
        key.push(value);
    }
    key.reverse();
    for index in 0..32 {
        data[index] ^= key[index % 16];
    }
    font_magic(data).is_some()
}

#[derive(Debug, Clone)]
struct XmlNode {
    name: String,
    attributes: Vec<(String, String)>,
    text: String,
    children: Vec<XmlNode>,
}

impl XmlNode {
    fn from_start(start: &BytesStart<'_>) -> Self {
        Self {
            name: String::from_utf8_lossy(local_name(start.name().as_ref())).into_owned(),
            attributes: start
                .attributes()
                .with_checks(false)
                .flatten()
                .map(|attribute| {
                    (
                        String::from_utf8_lossy(local_name(attribute.key.as_ref())).into_owned(),
                        String::from_utf8_lossy(attribute.value.as_ref()).into_owned(),
                    )
                })
                .collect(),
            text: String::new(),
            children: Vec::new(),
        }
    }

    fn attr(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find_map(|(key, value)| (key == name).then_some(value.as_str()))
    }

    fn child(&self, name: &str) -> Option<&XmlNode> {
        self.children.iter().find(|child| child.name == name)
    }

    fn children_named<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a XmlNode> {
        self.children.iter().filter(move |child| child.name == name)
    }

    fn descendant(&self, name: &str) -> Option<&XmlNode> {
        if self.name == name {
            return Some(self);
        }
        self.children
            .iter()
            .find_map(|child| child.descendant(name))
    }

    fn collect_descendants<'a>(&'a self, name: &str, output: &mut Vec<&'a XmlNode>) {
        if self.name == name {
            output.push(self);
        }
        for child in &self.children {
            child.collect_descendants(name, output);
        }
    }

    fn text_content(&self) -> String {
        let mut value = self.text.clone();
        for child in &self.children {
            value.push_str(&child.text_content());
        }
        value
    }
}

#[derive(Default)]
struct XmlBudget {
    nodes: usize,
    attributes: usize,
    text_bytes: usize,
    attribute_bytes: usize,
}

impl XmlBudget {
    fn consume_node(
        &mut self,
        start: &BytesStart<'_>,
        limits: &ParseLimits,
    ) -> Result<(), ParseError> {
        self.nodes = self.nodes.saturating_add(1);
        if self.nodes > limits.max_xml_nodes {
            return Err(ParseError::ResourceLimit("XML node count".into()));
        }
        for attribute in start.attributes().with_checks(false).flatten() {
            self.attributes = self.attributes.saturating_add(1);
            self.attribute_bytes = self
                .attribute_bytes
                .saturating_add(attribute.key.as_ref().len())
                .saturating_add(attribute.value.as_ref().len());
        }
        if self.attributes > limits.max_xml_attributes {
            return Err(ParseError::ResourceLimit("XML attribute count".into()));
        }
        if self.attribute_bytes > limits.max_xml_attribute_bytes {
            return Err(ParseError::ResourceLimit("XML attribute bytes".into()));
        }
        Ok(())
    }

    fn consume_text(&mut self, bytes: usize, limits: &ParseLimits) -> Result<(), ParseError> {
        self.text_bytes = self.text_bytes.saturating_add(bytes);
        if self.text_bytes > limits.max_xml_text_bytes {
            return Err(ParseError::ResourceLimit("XML text bytes".into()));
        }
        Ok(())
    }
}

fn parse_xml_tree(xml: &[u8], limits: &ParseLimits, context: &str) -> Result<XmlNode, ParseError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<XmlNode>::new();
    let mut root = None;
    let mut budget = XmlBudget::default();
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= limits.max_xml_depth {
                    return Err(ParseError::ResourceLimit("XML nesting depth".into()));
                }
                budget.consume_node(&start, limits)?;
                stack.push(XmlNode::from_start(&start));
            }
            Ok(Event::Empty(start)) => {
                budget.consume_node(&start, limits)?;
                let node = XmlNode::from_start(&start);
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else if root.replace(node).is_some() {
                    return Err(ParseError::Corrupt(format!(
                        "{context} contains multiple XML roots"
                    )));
                }
            }
            Ok(Event::Text(text)) => {
                if let Some(node) = stack.last_mut() {
                    let value = text
                        .unescape()
                        .map_err(|error| ParseError::Corrupt(format!("{context}: {error}")))?;
                    budget.consume_text(value.len(), limits)?;
                    node.text.push_str(&value);
                }
            }
            Ok(Event::CData(text)) => {
                if let Some(node) = stack.last_mut() {
                    budget.consume_text(text.as_ref().len(), limits)?;
                    node.text.push_str(&String::from_utf8_lossy(text.as_ref()));
                }
            }
            Ok(Event::End(_)) => {
                let node = stack.pop().ok_or_else(|| {
                    ParseError::Corrupt(format!("{context} contains an unmatched closing tag"))
                })?;
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else if root.replace(node).is_some() {
                    return Err(ParseError::Corrupt(format!(
                        "{context} contains multiple XML roots"
                    )));
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(ParseError::Corrupt(format!("{context}: {error}"))),
            _ => {}
        }
    }
    if !stack.is_empty() {
        return Err(ParseError::Corrupt(format!(
            "{context} contains unclosed XML elements"
        )));
    }
    root.ok_or_else(|| ParseError::Corrupt(format!("{context} has no XML root")))
}

fn bool_value(value: &str) -> Option<bool> {
    match value.to_ascii_lowercase().as_str() {
        "1" | "true" | "t" | "on" => Some(true),
        "0" | "false" | "f" | "off" => Some(false),
        _ => None,
    }
}

fn parse_transform(xfrm: Option<&XmlNode>) -> Transform {
    let Some(xfrm) = xfrm else {
        return Transform::default();
    };
    let off = xfrm.child("off");
    let ext = xfrm.child("ext");
    Transform {
        x: off
            .and_then(|node| node.attr("x"))
            .and_then(|value| value.parse().ok())
            .unwrap_or(0),
        y: off
            .and_then(|node| node.attr("y"))
            .and_then(|value| value.parse().ok())
            .unwrap_or(0),
        width: ext
            .and_then(|node| node.attr("cx"))
            .and_then(|value| value.parse().ok())
            .unwrap_or(0),
        height: ext
            .and_then(|node| node.attr("cy"))
            .and_then(|value| value.parse().ok())
            .unwrap_or(0),
        rotation: xfrm
            .attr("rot")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value / 60_000.0),
        flip_horizontal: xfrm.attr("flipH").and_then(bool_value),
        flip_vertical: xfrm.attr("flipV").and_then(bool_value),
    }
}

fn parse_child_transform(xfrm: Option<&XmlNode>) -> Option<Transform> {
    let xfrm = xfrm?;
    let off = xfrm.child("chOff")?;
    let ext = xfrm.child("chExt")?;
    Some(Transform {
        x: off
            .attr("x")
            .and_then(|value| value.parse().ok())
            .unwrap_or(0),
        y: off
            .attr("y")
            .and_then(|value| value.parse().ok())
            .unwrap_or(0),
        width: ext
            .attr("cx")
            .and_then(|value| value.parse().ok())
            .unwrap_or(0),
        height: ext
            .attr("cy")
            .and_then(|value| value.parse().ok())
            .unwrap_or(0),
        ..Default::default()
    })
}

fn node_identity(node: &XmlNode, slide_index: usize, node_index: usize) -> (String, String) {
    let properties = node.descendant("cNvPr");
    (
        properties
            .and_then(|value| value.attr("id"))
            .map(str::to_owned)
            .unwrap_or_else(|| format!("node-{slide_index}-{node_index}")),
        properties
            .and_then(|value| value.attr("name"))
            .unwrap_or_default()
            .to_owned(),
    )
}

fn color_element(node: &XmlNode) -> Option<&XmlNode> {
    if matches!(
        node.name.as_str(),
        "srgbClr" | "scrgbClr" | "schemeClr" | "sysClr" | "prstClr"
    ) {
        return Some(node);
    }
    node.children.iter().find_map(color_element)
}

#[derive(Debug, Clone)]
struct ThemeDataInternal {
    id: String,
    name: Option<String>,
    colors: BTreeMap<String, String>,
    major_fonts: BTreeMap<String, String>,
    minor_fonts: BTreeMap<String, String>,
}

impl ThemeDataInternal {
    fn public(&self) -> crate::PresentationTheme {
        crate::PresentationTheme {
            id: self.id.clone(),
            name: self.name.clone(),
            colors: self.colors.clone(),
            major_fonts: self.major_fonts.clone(),
            minor_fonts: self.minor_fonts.clone(),
        }
    }

    fn resolve_font(&self, typeface: &str) -> Option<String> {
        let fonts = if typeface.starts_with("+mj-") {
            &self.major_fonts
        } else if typeface.starts_with("+mn-") {
            &self.minor_fonts
        } else {
            return Some(typeface.to_owned());
        };
        let key = if typeface.ends_with("-ea") {
            "eastAsia"
        } else if typeface.ends_with("-cs") {
            "complexScript"
        } else {
            "latin"
        };
        fonts.get(key).cloned().filter(|font| !font.is_empty())
    }
}

#[derive(Debug)]
struct ParseDiagnostics {
    slide_index: Option<usize>,
    part_name: String,
    warnings: RefCell<Vec<PresentationWarning>>,
}

impl ParseDiagnostics {
    fn new(slide_index: Option<usize>, part_name: impl Into<String>) -> Self {
        Self {
            slide_index,
            part_name: part_name.into(),
            warnings: RefCell::new(Vec::new()),
        }
    }

    fn warn(&self, code: &str, message: impl Into<String>, feature: Option<&str>) {
        let mut warning = PresentationWarning::warning(code, message);
        warning.slide_index = self.slide_index;
        warning.part_name = Some(self.part_name.clone());
        warning.feature = feature.map(str::to_owned);
        self.warnings.borrow_mut().push(warning);
    }

    fn take(self) -> Vec<PresentationWarning> {
        self.warnings.into_inner()
    }
}

struct ColorContext<'a> {
    theme: Option<&'a ThemeDataInternal>,
    color_map: &'a BTreeMap<String, String>,
    diagnostics: &'a ParseDiagnostics,
}

fn preset_color(value: &str) -> Option<&'static str> {
    match value {
        "black" => Some("#000000"),
        "white" => Some("#FFFFFF"),
        "red" => Some("#FF0000"),
        "green" => Some("#008000"),
        "blue" => Some("#0000FF"),
        "yellow" => Some("#FFFF00"),
        "cyan" | "aqua" => Some("#00FFFF"),
        "magenta" | "fuchsia" => Some("#FF00FF"),
        "gray" | "grey" => Some("#808080"),
        "silver" => Some("#C0C0C0"),
        "maroon" | "darkRed" => Some("#800000"),
        "olive" | "darkYellow" => Some("#808000"),
        "lime" => Some("#00FF00"),
        "teal" | "darkCyan" => Some("#008080"),
        "navy" | "darkBlue" => Some("#000080"),
        "purple" | "darkMagenta" => Some("#800080"),
        "darkGreen" => Some("#006400"),
        "orange" => Some("#FFA500"),
        "gold" => Some("#FFD700"),
        "brown" => Some("#A52A2A"),
        "pink" => Some("#FFC0CB"),
        "violet" => Some("#EE82EE"),
        "lightGray" | "lightGrey" => Some("#D3D3D3"),
        _ => None,
    }
}

fn normalized_hex(value: &str) -> Option<String> {
    let value = value.strip_prefix('#').unwrap_or(value);
    (value.len() == 6 && value.bytes().all(|byte| byte.is_ascii_hexdigit()))
        .then(|| format!("#{}", value.to_ascii_uppercase()))
}

fn rgb_from_hex(value: &str) -> Option<[f64; 3]> {
    let value = value.strip_prefix('#').unwrap_or(value);
    if value.len() != 6 {
        return None;
    }
    Some([
        u8::from_str_radix(&value[0..2], 16).ok()? as f64,
        u8::from_str_radix(&value[2..4], 16).ok()? as f64,
        u8::from_str_radix(&value[4..6], 16).ok()? as f64,
    ])
}

fn transformed_color(value: &str, color: &XmlNode) -> Option<String> {
    let mut rgb = rgb_from_hex(value)?;
    for transform in &color.children {
        let amount = transform
            .attr("val")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value / 100_000.0);
        match (transform.name.as_str(), amount) {
            ("tint", Some(amount)) => {
                for channel in &mut rgb {
                    *channel += (255.0 - *channel) * amount;
                }
            }
            ("shade", Some(amount)) | ("lumMod", Some(amount)) => {
                for channel in &mut rgb {
                    *channel *= amount;
                }
            }
            ("lumOff", Some(amount)) => {
                for channel in &mut rgb {
                    *channel += 255.0 * amount;
                }
            }
            _ => {}
        }
    }
    Some(format!(
        "#{:02X}{:02X}{:02X}",
        rgb[0].round().clamp(0.0, 255.0) as u8,
        rgb[1].round().clamp(0.0, 255.0) as u8,
        rgb[2].round().clamp(0.0, 255.0) as u8
    ))
}

fn parse_color(node: &XmlNode, context: Option<&ColorContext<'_>>) -> Option<ColorValue> {
    let color = color_element(node)?;
    let base_value = match color.name.as_str() {
        "srgbClr" => normalized_hex(color.attr("val")?),
        "scrgbClr" => {
            let channel = |name: &str| {
                color
                    .attr(name)
                    .and_then(|value| value.parse::<f64>().ok())
                    .map(|value| (value * 255.0 / 100_000.0).round().clamp(0.0, 255.0) as u8)
                    .unwrap_or(0)
            };
            Some(format!(
                "#{:02X}{:02X}{:02X}",
                channel("r"),
                channel("g"),
                channel("b")
            ))
        }
        "schemeClr" => {
            let requested = color.attr("val")?;
            let context = context?;
            let mapped = context
                .color_map
                .get(requested)
                .map(String::as_str)
                .unwrap_or(requested);
            let resolved = context
                .theme
                .and_then(|theme| theme.colors.get(mapped))
                .cloned();
            if resolved.is_none() {
                context.diagnostics.warn(
                    "degraded-rendering",
                    format!("Theme color {requested} could not be resolved."),
                    Some("theme-color"),
                );
            }
            resolved
        }
        "sysClr" => color.attr("lastClr").and_then(normalized_hex),
        "prstClr" => {
            let preset = color.attr("val")?;
            let resolved = preset_color(preset).map(str::to_owned);
            if resolved.is_none() {
                if let Some(context) = context {
                    context.diagnostics.warn(
                        "degraded-rendering",
                        "An unsupported preset color was ignored.",
                        Some("preset-color"),
                    );
                }
            }
            resolved
        }
        _ => None,
    }?;
    let value = transformed_color(&base_value, color)?;
    let mut alpha = color
        .child("alpha")
        .and_then(|node| node.attr("val"))
        .and_then(|value| value.parse::<f64>().ok())
        .map(|value| (value / 100_000.0).clamp(0.0, 1.0))
        .unwrap_or(1.0);
    if let Some(value) = color
        .child("alphaMod")
        .and_then(|node| node.attr("val"))
        .and_then(|value| value.parse::<f64>().ok())
    {
        alpha *= value / 100_000.0;
    }
    if let Some(value) = color
        .child("alphaOff")
        .and_then(|node| node.attr("val"))
        .and_then(|value| value.parse::<f64>().ok())
    {
        alpha += value / 100_000.0;
    }
    let alpha = (alpha < 1.0).then(|| alpha.clamp(0.0, 1.0));
    Some(ColorValue { value, alpha })
}

fn parse_theme(
    xml: &[u8],
    path: &str,
    limits: &ParseLimits,
) -> Result<ThemeDataInternal, ParseError> {
    let root = parse_xml_tree(xml, limits, path)?;
    let color_scheme = root.descendant("clrScheme");
    let mut colors = BTreeMap::new();
    if let Some(color_scheme) = color_scheme {
        for entry in &color_scheme.children {
            if let Some(color) = parse_color(entry, None) {
                colors.insert(entry.name.clone(), color.value);
            }
        }
    }
    let parse_fonts = |node: Option<&XmlNode>| {
        let mut fonts = BTreeMap::new();
        if let Some(node) = node {
            for (element, key) in [
                ("latin", "latin"),
                ("ea", "eastAsia"),
                ("cs", "complexScript"),
            ] {
                if let Some(typeface) = node.child(element).and_then(|font| font.attr("typeface")) {
                    if !typeface.is_empty() {
                        fonts.insert(key.to_owned(), typeface.to_owned());
                    }
                }
            }
        }
        fonts
    };
    let font_scheme = root.descendant("fontScheme");
    Ok(ThemeDataInternal {
        id: path.to_owned(),
        name: root
            .attr("name")
            .or_else(|| color_scheme.and_then(|scheme| scheme.attr("name")))
            .map(str::to_owned),
        colors,
        major_fonts: parse_fonts(font_scheme.and_then(|scheme| scheme.child("majorFont"))),
        minor_fonts: parse_fonts(font_scheme.and_then(|scheme| scheme.child("minorFont"))),
    })
}

fn color_map(root: &XmlNode) -> BTreeMap<String, String> {
    let mapping = root.child("clrMap").or_else(|| {
        root.child("clrMapOvr")
            .and_then(|override_node| override_node.child("overrideClrMapping"))
    });
    mapping
        .map(|mapping| mapping.attributes.iter().cloned().collect())
        .unwrap_or_default()
}

fn merge_color_maps(
    base: &BTreeMap<String, String>,
    override_map: &BTreeMap<String, String>,
) -> BTreeMap<String, String> {
    let mut result = base.clone();
    result.extend(override_map.clone());
    result
}

fn parse_fill(properties: &XmlNode, context: Option<&ColorContext<'_>>) -> Option<FillStyle> {
    if properties.child("noFill").is_some() {
        return Some(FillStyle::None);
    }
    if let Some(solid) = properties.child("solidFill") {
        return parse_color(solid, context).map(|color| FillStyle::Solid { color });
    }
    if let Some(gradient) = properties.child("gradFill") {
        let stops = gradient
            .child("gsLst")
            .into_iter()
            .flat_map(|list| list.children_named("gs"))
            .filter_map(|stop| {
                Some(GradientStop {
                    position: stop
                        .attr("pos")
                        .and_then(|value| value.parse::<f64>().ok())
                        .unwrap_or(0.0)
                        / 100_000.0,
                    color: parse_color(stop, context)?,
                })
            })
            .collect::<Vec<_>>();
        if !stops.is_empty() {
            let angle = gradient
                .child("lin")
                .and_then(|node| node.attr("ang"))
                .and_then(|value| value.parse::<f64>().ok())
                .map(|value| value / 60_000.0);
            return Some(FillStyle::Gradient { angle, stops });
        }
    }
    if let Some(pattern) = properties.child("pattFill") {
        let foreground = parse_color(pattern.child("fgClr")?, context)?;
        let background = parse_color(pattern.child("bgClr")?, context)?;
        return Some(FillStyle::Pattern {
            preset: pattern.attr("prst").unwrap_or("pct5").to_owned(),
            foreground,
            background,
        });
    }
    None
}

fn parse_line_node(line: &XmlNode, context: Option<&ColorContext<'_>>) -> Option<LineStyle> {
    if line.child("noFill").is_some() {
        return None;
    }
    let color = line
        .child("solidFill")
        .and_then(|fill| parse_color(fill, context))
        .or_else(|| {
            line.child("gradFill").and_then(|gradient| {
                gradient
                    .child("gsLst")
                    .and_then(|list| list.child("gs"))
                    .and_then(|stop| parse_color(stop, context))
            })
        });
    let width = line
        .attr("w")
        .and_then(|value| value.parse::<f64>().ok())
        .map(|value| value / 9_525.0)
        .or_else(|| color.as_ref().map(|_| 1.0));
    let dash = line
        .child("prstDash")
        .and_then(|value| value.attr("val"))
        .map(str::to_owned);
    let start_arrow = line
        .child("headEnd")
        .and_then(|value| value.attr("type"))
        .filter(|value| *value != "none")
        .map(str::to_owned);
    let end_arrow = line
        .child("tailEnd")
        .and_then(|value| value.attr("type"))
        .filter(|value| *value != "none")
        .map(str::to_owned);
    (color.is_some()
        || width.is_some()
        || dash.is_some()
        || start_arrow.is_some()
        || end_arrow.is_some())
    .then_some(LineStyle {
        color,
        width,
        dash,
        start_arrow,
        end_arrow,
    })
}

#[derive(Debug, Clone, Default)]
struct RunStyle {
    font_family: Option<String>,
    font_size_pt: Option<f64>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    color: Option<ColorValue>,
    baseline: Option<f64>,
    language: Option<String>,
    hyperlink: Option<String>,
}

impl RunStyle {
    fn overlay(mut self, other: RunStyle) -> Self {
        if other.font_family.is_some() {
            self.font_family = other.font_family;
        }
        if other.font_size_pt.is_some() {
            self.font_size_pt = other.font_size_pt;
        }
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
        if other.strike.is_some() {
            self.strike = other.strike;
        }
        if other.color.is_some() {
            self.color = other.color;
        }
        if other.baseline.is_some() {
            self.baseline = other.baseline;
        }
        if other.language.is_some() {
            self.language = other.language;
        }
        if other.hyperlink.is_some() {
            self.hyperlink = other.hyperlink;
        }
        self
    }
}

fn parse_run_style(
    properties: Option<&XmlNode>,
    colors: Option<&ColorContext<'_>>,
    rels: &HashMap<String, String>,
) -> RunStyle {
    let Some(properties) = properties else {
        return RunStyle::default();
    };
    let font_family = properties
        .child("latin")
        .and_then(|node| node.attr("typeface"))
        .and_then(|typeface| {
            colors
                .and_then(|context| context.theme)
                .and_then(|theme| theme.resolve_font(typeface))
                .or_else(|| (!typeface.starts_with('+')).then(|| typeface.to_owned()))
        });
    let hyperlink = properties
        .child("hlinkClick")
        .and_then(|link| link.attr("id"))
        .and_then(|id| rels.get(id))
        .cloned();
    RunStyle {
        font_family,
        font_size_pt: properties
            .attr("sz")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value / 100.0),
        bold: properties.attr("b").and_then(bool_value),
        italic: properties.attr("i").and_then(bool_value),
        underline: properties
            .attr("u")
            .map(|value| value != "none" && value != "0"),
        strike: properties
            .attr("strike")
            .map(|value| value != "noStrike" && value != "none" && value != "0"),
        color: parse_color(properties, colors),
        baseline: properties
            .attr("baseline")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value / 1_000.0),
        language: properties.attr("lang").map(str::to_owned),
        hyperlink,
    }
}

fn text_run(text: String, style: RunStyle) -> TextRun {
    TextRun {
        text,
        font_family: style.font_family,
        font_size_pt: style.font_size_pt,
        bold: style.bold,
        italic: style.italic,
        underline: style.underline,
        strike: style.strike,
        color: style.color,
        baseline: style.baseline,
        language: style.language,
        hyperlink: style.hyperlink,
    }
}

#[derive(Clone, Default)]
struct ParagraphStyle {
    alignment: Option<String>,
    level: Option<usize>,
    bullet: Option<Option<crate::TextBullet>>,
    line_spacing: Option<f64>,
    space_before: Option<f64>,
    space_after: Option<f64>,
    rtl: Option<bool>,
}

impl ParagraphStyle {
    fn overlay(mut self, other: ParagraphStyle) -> Self {
        if other.alignment.is_some() {
            self.alignment = other.alignment;
        }
        if other.level.is_some() {
            self.level = other.level;
        }
        if other.bullet.is_some() {
            self.bullet = other.bullet;
        }
        if other.line_spacing.is_some() {
            self.line_spacing = other.line_spacing;
        }
        if other.space_before.is_some() {
            self.space_before = other.space_before;
        }
        if other.space_after.is_some() {
            self.space_after = other.space_after;
        }
        if other.rtl.is_some() {
            self.rtl = other.rtl;
        }
        self
    }
}

fn spacing_value(node: Option<&XmlNode>) -> Option<f64> {
    let node = node?;
    node.child("spcPts")
        .and_then(|value| value.attr("val"))
        .and_then(|value| value.parse::<f64>().ok())
        .map(|value| value / 100.0)
        .or_else(|| {
            node.child("spcPct")
                .and_then(|value| value.attr("val"))
                .and_then(|value| value.parse::<f64>().ok())
                .map(|value| value / 1_000.0)
        })
}

fn parse_paragraph_style(properties: Option<&XmlNode>) -> ParagraphStyle {
    let Some(properties) = properties else {
        return ParagraphStyle::default();
    };
    let alignment = properties.attr("algn").and_then(|value| {
        match value {
            "l" => Some("left"),
            "ctr" => Some("center"),
            "r" => Some("right"),
            "just" | "justLow" => Some("justify"),
            "dist" | "thaiDist" => Some("distributed"),
            _ => None,
        }
        .map(str::to_owned)
    });
    let bullet = if let Some(character) = properties.child("buChar") {
        Some(Some(crate::TextBullet {
            kind: "character".into(),
            value: character.attr("char").map(str::to_owned),
        }))
    } else if let Some(number) = properties.child("buAutoNum") {
        Some(Some(crate::TextBullet {
            kind: "number".into(),
            value: number.attr("type").map(str::to_owned),
        }))
    } else if properties.child("buBlip").is_some() {
        Some(Some(crate::TextBullet {
            kind: "picture".into(),
            value: None,
        }))
    } else if properties.child("buNone").is_some() {
        Some(None)
    } else {
        None
    };
    ParagraphStyle {
        alignment,
        level: properties
            .attr("lvl")
            .and_then(|value| value.parse::<usize>().ok()),
        bullet,
        line_spacing: spacing_value(properties.child("lnSpc")),
        space_before: spacing_value(properties.child("spcBef")),
        space_after: spacing_value(properties.child("spcAft")),
        rtl: properties.attr("rtl").and_then(bool_value),
    }
}

fn fallback_paragraph_properties(
    text_body: Option<&XmlNode>,
    paragraph_index: usize,
    level: usize,
) -> Option<&XmlNode> {
    let text_body = text_body?;
    text_body
        .children_named("p")
        .nth(paragraph_index)
        .and_then(|paragraph| paragraph.child("pPr"))
        .or_else(|| {
            text_body
                .child("lstStyle")
                .and_then(|style| style.child(&format!("lvl{}pPr", level.saturating_add(1))))
        })
}

fn style_level_properties(style: Option<&XmlNode>, level: usize) -> Option<&XmlNode> {
    style.and_then(|style| style.child(&format!("lvl{}pPr", level.saturating_add(1))))
}

fn parse_text_paragraphs(
    text_body: Option<&XmlNode>,
    layout_text_body: Option<&XmlNode>,
    master_text_body: Option<&XmlNode>,
    master_text_style: Option<&XmlNode>,
    colors: Option<&ColorContext<'_>>,
    rels: &HashMap<String, String>,
) -> Vec<TextParagraph> {
    let Some(text_body) = text_body else {
        return Vec::new();
    };
    text_body
        .children_named("p")
        .enumerate()
        .map(|(paragraph_index, paragraph)| {
            let own_properties = paragraph.child("pPr");
            let level = own_properties
                .and_then(|properties| properties.attr("lvl"))
                .and_then(|value| value.parse::<usize>().ok())
                .or_else(|| {
                    layout_text_body
                        .and_then(|body| {
                            fallback_paragraph_properties(Some(body), paragraph_index, 0)
                        })
                        .and_then(|properties| properties.attr("lvl"))
                        .and_then(|value| value.parse::<usize>().ok())
                })
                .unwrap_or(0);
            let master_style_properties = style_level_properties(master_text_style, level);
            let master_properties =
                fallback_paragraph_properties(master_text_body, paragraph_index, level);
            let layout_properties =
                fallback_paragraph_properties(layout_text_body, paragraph_index, level);
            let paragraph_style = parse_paragraph_style(master_style_properties)
                .overlay(parse_paragraph_style(master_properties))
                .overlay(parse_paragraph_style(layout_properties))
                .overlay(parse_paragraph_style(own_properties));
            let default_style = parse_run_style(
                master_style_properties.and_then(|properties| properties.child("defRPr")),
                colors,
                rels,
            )
            .overlay(parse_run_style(
                master_properties.and_then(|properties| properties.child("defRPr")),
                colors,
                rels,
            ))
            .overlay(parse_run_style(
                layout_properties.and_then(|properties| properties.child("defRPr")),
                colors,
                rels,
            ))
            .overlay(parse_run_style(
                own_properties.and_then(|properties| properties.child("defRPr")),
                colors,
                rels,
            ));
            let mut runs = Vec::new();
            for child in &paragraph.children {
                match child.name.as_str() {
                    "r" | "fld" => {
                        let style = default_style.clone().overlay(parse_run_style(
                            child.child("rPr"),
                            colors,
                            rels,
                        ));
                        let value = child
                            .children_named("t")
                            .map(XmlNode::text_content)
                            .collect::<String>();
                        runs.push(text_run(value, style));
                    }
                    "br" => {
                        let style = default_style.clone().overlay(parse_run_style(
                            child.child("rPr"),
                            colors,
                            rels,
                        ));
                        runs.push(text_run("\n".into(), style));
                    }
                    "t" => runs.push(text_run(child.text_content(), default_style.clone())),
                    _ => {}
                }
            }
            TextParagraph {
                runs,
                alignment: paragraph_style.alignment,
                level: paragraph_style
                    .level
                    .or(Some(level))
                    .filter(|level| *level > 0),
                bullet: paragraph_style.bullet.flatten(),
                line_spacing: paragraph_style.line_spacing,
                space_before: paragraph_style.space_before,
                space_after: paragraph_style.space_after,
                rtl: paragraph_style.rtl,
            }
        })
        .collect()
}

fn vertical_alignment(text_body: Option<&XmlNode>) -> Option<VerticalAlignment> {
    match text_body
        .and_then(|body| body.child("bodyPr"))
        .and_then(|properties| properties.attr("anchor"))
    {
        Some("ctr") | Some("just") | Some("dist") => Some(VerticalAlignment::Middle),
        Some("b") => Some(VerticalAlignment::Bottom),
        Some("t") => Some(VerticalAlignment::Top),
        _ => None,
    }
}

fn shape_tree(root: &XmlNode) -> Option<&XmlNode> {
    root.child("cSld").and_then(|slide| slide.child("spTree"))
}

#[derive(Clone)]
struct PlaceholderKey {
    index: Option<String>,
    kind: Option<String>,
}

fn normalized_placeholder_kind(kind: &str) -> &str {
    match kind {
        "ctrTitle" | "title" => "title",
        "body" | "obj" | "subTitle" => "body",
        value => value,
    }
}

fn placeholder_key(node: &XmlNode) -> Option<PlaceholderKey> {
    let placeholder = node.descendant("ph")?;
    Some(PlaceholderKey {
        index: placeholder.attr("idx").map(str::to_owned),
        kind: placeholder
            .attr("type")
            .map(normalized_placeholder_kind)
            .map(str::to_owned),
    })
}

fn find_placeholder<'a>(root: &'a XmlNode, source: &XmlNode) -> Option<&'a XmlNode> {
    let key = placeholder_key(source)?;
    let candidates = shape_tree(root)?.children.iter().filter(|candidate| {
        matches!(
            candidate.name.as_str(),
            "sp" | "cxnSp" | "pic" | "graphicFrame"
        )
    });
    if let Some(index) = key.index.as_deref() {
        if let Some(candidate) = candidates.clone().find(|candidate| {
            placeholder_key(candidate).and_then(|value| value.index) == Some(index.to_owned())
        }) {
            return Some(candidate);
        }
    }
    key.kind.as_deref().and_then(|kind| {
        candidates.into_iter().find(|candidate| {
            placeholder_key(candidate).and_then(|value| value.kind) == Some(kind.to_owned())
        })
    })
}

fn master_text_style<'a>(master: Option<&'a XmlNode>, source: &XmlNode) -> Option<&'a XmlNode> {
    let styles = master?.child("txStyles")?;
    match placeholder_key(source).and_then(|key| key.kind) {
        Some(kind) if kind == "title" => styles.child("titleStyle"),
        Some(kind) if kind == "body" => styles.child("bodyStyle"),
        _ => styles.child("otherStyle"),
    }
}

fn apply_transform(target: &mut Transform, xfrm: Option<&XmlNode>) {
    let Some(xfrm) = xfrm else {
        return;
    };
    if let Some(off) = xfrm.child("off") {
        if let Some(value) = off.attr("x").and_then(|value| value.parse().ok()) {
            target.x = value;
        }
        if let Some(value) = off.attr("y").and_then(|value| value.parse().ok()) {
            target.y = value;
        }
    }
    if let Some(ext) = xfrm.child("ext") {
        if let Some(value) = ext.attr("cx").and_then(|value| value.parse().ok()) {
            target.width = value;
        }
        if let Some(value) = ext.attr("cy").and_then(|value| value.parse().ok()) {
            target.height = value;
        }
    }
    if let Some(value) = xfrm.attr("rot").and_then(|value| value.parse::<f64>().ok()) {
        target.rotation = Some(value / 60_000.0);
    }
    if let Some(value) = xfrm.attr("flipH").and_then(bool_value) {
        target.flip_horizontal = Some(value);
    }
    if let Some(value) = xfrm.attr("flipV").and_then(bool_value) {
        target.flip_vertical = Some(value);
    }
}

fn shape_xfrm(node: Option<&XmlNode>) -> Option<&XmlNode> {
    node?.child("spPr")?.child("xfrm")
}

fn inherited_shape_transform(
    own: &XmlNode,
    layout: Option<&XmlNode>,
    master: Option<&XmlNode>,
) -> Transform {
    let mut transform = Transform::default();
    apply_transform(&mut transform, shape_xfrm(master));
    apply_transform(&mut transform, shape_xfrm(layout));
    apply_transform(&mut transform, shape_xfrm(Some(own)));
    transform
}

fn shape_fill(node: &XmlNode, colors: Option<&ColorContext<'_>>) -> Option<Option<FillStyle>> {
    if let Some(properties) = node.child("spPr") {
        let has_fill = [
            "noFill",
            "solidFill",
            "gradFill",
            "pattFill",
            "blipFill",
            "grpFill",
        ]
        .into_iter()
        .any(|name| properties.child(name).is_some());
        if has_fill {
            return Some(parse_fill(properties, colors));
        }
    }
    node.child("style")
        .and_then(|style| style.child("fillRef"))
        .map(|fill| parse_color(fill, colors).map(|color| FillStyle::Solid { color }))
}

fn shape_line(node: &XmlNode, colors: Option<&ColorContext<'_>>) -> Option<Option<LineStyle>> {
    if let Some(line) = node
        .child("spPr")
        .and_then(|properties| properties.child("ln"))
    {
        return Some(parse_line_node(line, colors));
    }
    node.child("style")
        .and_then(|style| style.child("lnRef"))
        .map(|line| {
            parse_color(line, colors).map(|color| LineStyle {
                color: Some(color),
                ..Default::default()
            })
        })
}

struct NodeParseContext<'a, 'color> {
    slide_index: usize,
    rels: &'a HashMap<String, String>,
    related_parts: &'a HashMap<String, Vec<u8>>,
    limits: &'a ParseLimits,
    colors: Option<&'a ColorContext<'color>>,
    diagnostics: &'a ParseDiagnostics,
    layout_root: Option<&'a XmlNode>,
    master_root: Option<&'a XmlNode>,
}

fn parse_shape_node(
    node: &XmlNode,
    node_index: usize,
    context: &NodeParseContext<'_, '_>,
) -> SlideNode {
    let (id, name) = node_identity(node, context.slide_index, node_index);
    let layout = context
        .layout_root
        .and_then(|root| find_placeholder(root, node));
    let master_source = layout.unwrap_or(node);
    let master = context
        .master_root
        .and_then(|root| find_placeholder(root, master_source));
    let text_body = node.child("txBody");
    let layout_text_body = layout.and_then(|shape| shape.child("txBody"));
    let master_text_body = master.and_then(|shape| shape.child("txBody"));
    let geometry = [Some(node), layout, master]
        .into_iter()
        .flatten()
        .find_map(|shape| {
            shape
                .child("spPr")
                .and_then(|properties| properties.child("prstGeom"))
                .and_then(|geometry| geometry.attr("prst"))
                .map(str::to_owned)
        });
    let fill = [Some(node), layout, master]
        .into_iter()
        .flatten()
        .find_map(|shape| shape_fill(shape, context.colors))
        .flatten();
    let line = [Some(node), layout, master]
        .into_iter()
        .flatten()
        .find_map(|shape| shape_line(shape, context.colors))
        .flatten();
    let vertical_alignment = [text_body, layout_text_body, master_text_body]
        .into_iter()
        .find_map(vertical_alignment);
    SlideNode::Shape {
        id,
        name,
        transform: inherited_shape_transform(node, layout, master),
        geometry: ShapeGeometry { preset: geometry },
        fill,
        line,
        paragraphs: parse_text_paragraphs(
            text_body,
            layout_text_body,
            master_text_body,
            master_text_style(context.master_root, master_source),
            context.colors,
            context.rels,
        ),
        vertical_alignment,
    }
}

fn parse_image_node(
    node: &XmlNode,
    node_index: usize,
    context: &NodeParseContext<'_, '_>,
) -> SlideNode {
    let (id, name) = node_identity(node, context.slide_index, node_index);
    let relationship = node
        .child("blipFill")
        .and_then(|fill| fill.child("blip"))
        .and_then(|blip| blip.attr("embed"));
    let asset_id = relationship
        .and_then(|relationship| context.rels.get(relationship))
        .cloned();
    if asset_id.is_none() {
        context.diagnostics.warn(
            "missing-asset",
            format!("Image {id} has no resolvable embedded asset."),
            Some("image"),
        );
    }
    SlideNode::Image {
        id,
        name,
        transform: parse_transform(
            node.child("spPr")
                .and_then(|properties| properties.child("xfrm")),
        ),
        asset_id: asset_id.unwrap_or_else(|| "missing-asset".into()),
        preserve_aspect_ratio: true,
    }
}

fn parse_table_cell(cell: &XmlNode, context: &NodeParseContext<'_, '_>) -> TableCell {
    let properties = cell.child("tcPr");
    let mut borders = BTreeMap::new();
    if let Some(properties) = properties {
        for (element, side) in [
            ("lnT", "top"),
            ("lnR", "right"),
            ("lnB", "bottom"),
            ("lnL", "left"),
        ] {
            if let Some(line) = properties
                .child(element)
                .and_then(|line| parse_line_node(line, context.colors))
            {
                borders.insert(side.to_owned(), line);
            }
        }
    }
    TableCell {
        row_span: cell
            .attr("rowSpan")
            .and_then(|value| value.parse().ok())
            .filter(|value| *value > 1),
        col_span: cell
            .attr("gridSpan")
            .and_then(|value| value.parse().ok())
            .filter(|value| *value > 1),
        fill: properties.and_then(|properties| parse_fill(properties, context.colors)),
        borders,
        paragraphs: parse_text_paragraphs(
            cell.child("txBody"),
            None,
            None,
            None,
            context.colors,
            context.rels,
        ),
    }
}

fn is_merge_continuation(cell: &XmlNode) -> bool {
    ["hMerge", "vMerge"]
        .into_iter()
        .filter_map(|name| cell.attr(name))
        .any(|value| bool_value(value).unwrap_or(false))
}

fn parse_table_node(
    frame: &XmlNode,
    table: &XmlNode,
    node_index: usize,
    context: &NodeParseContext<'_, '_>,
) -> SlideNode {
    let (id, name) = node_identity(frame, context.slide_index, node_index);
    let rows = table
        .children_named("tr")
        .map(|row| {
            row.children_named("tc")
                .filter(|cell| !is_merge_continuation(cell))
                .map(|cell| parse_table_cell(cell, context))
                .collect()
        })
        .collect::<Vec<Vec<TableCell>>>();
    let column_widths = table
        .child("tblGrid")
        .into_iter()
        .flat_map(|grid| grid.children_named("gridCol"))
        .filter_map(|column| column.attr("w").and_then(|value| value.parse().ok()))
        .collect();
    let row_heights = table
        .children_named("tr")
        .filter_map(|row| row.attr("h").and_then(|value| value.parse().ok()))
        .collect();
    SlideNode::Table {
        id,
        name,
        transform: parse_transform(frame.child("xfrm")),
        rows,
        column_widths,
        row_heights,
    }
}

fn indexed_cache_values(node: &XmlNode) -> Vec<Option<String>> {
    let cache = [
        "strCache",
        "numCache",
        "strLit",
        "numLit",
        "multiLvlStrCache",
    ]
    .into_iter()
    .find_map(|name| node.descendant(name))
    .unwrap_or(node);
    let mut points = Vec::new();
    cache.collect_descendants("pt", &mut points);
    let declared = cache
        .descendant("ptCount")
        .and_then(|count| count.attr("val"))
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(0);
    let observed = points
        .iter()
        .filter_map(|point| point.attr("idx")?.parse::<usize>().ok())
        .max()
        .map(|index| index.saturating_add(1))
        .unwrap_or(0);
    let mut values = vec![None; declared.max(observed)];
    for point in points {
        let Some(index) = point
            .attr("idx")
            .and_then(|value| value.parse::<usize>().ok())
        else {
            continue;
        };
        if index >= values.len() {
            values.resize(index.saturating_add(1), None);
        }
        values[index] = point
            .child("v")
            .or_else(|| point.child("t"))
            .map(XmlNode::text_content);
    }
    values
}

fn category_cache_is_numeric(node: &XmlNode) -> bool {
    node.descendant("numCache").is_some() || node.descendant("numLit").is_some()
}

fn chart_title(root: &XmlNode) -> Option<String> {
    let title = root.descendant("title")?;
    let mut values = Vec::new();
    for name in ["t", "v"] {
        let mut nodes = Vec::new();
        title.collect_descendants(name, &mut nodes);
        values.extend(
            nodes
                .into_iter()
                .map(XmlNode::text_content)
                .filter(|value| !value.trim().is_empty()),
        );
    }
    (!values.is_empty()).then(|| values.join(""))
}

struct ParsedChart {
    chart_type: String,
    title: Option<String>,
    series: Vec<ChartSeries>,
    has_legend: Option<bool>,
}

fn parse_chart_part(
    xml: &[u8],
    limits: &ParseLimits,
    part_name: &str,
    colors: Option<&ColorContext<'_>>,
) -> Result<ParsedChart, ParseError> {
    let root = parse_xml_tree(xml, limits, part_name)?;
    let chart_types = [
        ("barChart", "bar"),
        ("lineChart", "line"),
        ("pieChart", "pie"),
        ("doughnutChart", "doughnut"),
        ("areaChart", "area"),
        ("scatterChart", "scatter"),
        ("radarChart", "radar"),
        ("bubbleChart", "bubble"),
        ("stockChart", "stock"),
        ("surfaceChart", "surface"),
    ];
    let (chart_node, chart_type) = chart_types
        .iter()
        .find_map(|(element, kind)| root.descendant(element).map(|node| (node, *kind)))
        .unwrap_or((&root, "ooxml"));
    let series = chart_node
        .children_named("ser")
        .map(|series| {
            let name = series
                .child("tx")
                .and_then(|value| indexed_cache_values(value).into_iter().flatten().next())
                .or_else(|| {
                    series
                        .child("tx")
                        .and_then(|value| value.child("v"))
                        .map(XmlNode::text_content)
                });
            let categories = series
                .child("cat")
                .or_else(|| series.child("xVal"))
                .map(|categories| {
                    let numeric = category_cache_is_numeric(categories);
                    indexed_cache_values(categories)
                        .into_iter()
                        .map(|value| match value {
                            None => serde_json::Value::Null,
                            Some(value) if numeric => value
                                .parse::<f64>()
                                .ok()
                                .and_then(serde_json::Number::from_f64)
                                .map(serde_json::Value::Number)
                                .unwrap_or(serde_json::Value::Null),
                            Some(value) => serde_json::Value::String(value),
                        })
                        .collect()
                })
                .filter(|values: &Vec<serde_json::Value>| !values.is_empty());
            let values = series
                .child("val")
                .or_else(|| series.child("yVal"))
                .map(indexed_cache_values)
                .unwrap_or_default()
                .into_iter()
                .map(|value| value.and_then(|value| value.parse::<f64>().ok()))
                .collect();
            let color = series
                .child("spPr")
                .and_then(|properties| properties.child("solidFill"))
                .and_then(|fill| parse_color(fill, colors));
            ChartSeries {
                name,
                categories,
                values,
                color,
            }
        })
        .collect();
    Ok(ParsedChart {
        chart_type: chart_type.into(),
        title: chart_title(&root),
        series,
        has_legend: Some(root.descendant("legend").is_some()),
    })
}

fn parse_graphic_frame(
    node: &XmlNode,
    node_index: usize,
    context: &NodeParseContext<'_, '_>,
) -> Result<SlideNode, ParseError> {
    let graphic_data = node.descendant("graphicData");
    if let Some(table) = graphic_data.and_then(|value| value.child("tbl")) {
        return Ok(parse_table_node(node, table, node_index, context));
    }
    if let Some(chart) = graphic_data.and_then(|value| value.child("chart")) {
        let relationship = chart.attr("id");
        let path = relationship.and_then(|value| context.rels.get(value));
        let chart = if let Some((path, xml)) =
            path.and_then(|path| context.related_parts.get(path).map(|xml| (path, xml)))
        {
            parse_chart_part(xml, context.limits, path, context.colors)?
        } else {
            context.diagnostics.warn(
                "missing-part",
                "Chart relationship or chart part could not be resolved.",
                Some("chart"),
            );
            ParsedChart {
                chart_type: "ooxml".into(),
                title: None,
                series: Vec::new(),
                has_legend: None,
            }
        };
        let (id, name) = node_identity(node, context.slide_index, node_index);
        return Ok(SlideNode::Chart {
            id,
            name,
            transform: parse_transform(node.child("xfrm")),
            chart_type: chart.chart_type,
            title: chart.title,
            series: chart.series,
            has_legend: chart.has_legend,
        });
    }
    let (id, name) = node_identity(node, context.slide_index, node_index);
    context.diagnostics.warn(
        "unsupported-feature",
        format!("Graphic frame {id} contains an unsupported payload."),
        Some("graphic-frame"),
    );
    Ok(SlideNode::Unknown {
        id,
        name,
        transform: parse_transform(node.child("xfrm")),
        feature: "graphic-frame".into(),
    })
}

fn parse_group_node(
    node: &XmlNode,
    node_index: usize,
    context: &NodeParseContext<'_, '_>,
) -> Result<SlideNode, ParseError> {
    let (id, name) = node_identity(node, context.slide_index, node_index);
    let xfrm = node
        .child("grpSpPr")
        .and_then(|properties| properties.child("xfrm"));
    let mut children = Vec::new();
    for (child_index, child) in node.children.iter().enumerate() {
        if let Some(child) = parse_slide_node(child, child_index, context)? {
            children.push(child);
        }
    }
    Ok(SlideNode::Group {
        id,
        name,
        transform: parse_transform(xfrm),
        children,
        child_transform: parse_child_transform(xfrm),
    })
}

fn parse_slide_node(
    node: &XmlNode,
    node_index: usize,
    context: &NodeParseContext<'_, '_>,
) -> Result<Option<SlideNode>, ParseError> {
    match node.name.as_str() {
        "sp" | "cxnSp" => Ok(Some(parse_shape_node(node, node_index, context))),
        "pic" => Ok(Some(parse_image_node(node, node_index, context))),
        "grpSp" => Ok(Some(parse_group_node(node, node_index, context)?)),
        "graphicFrame" => Ok(Some(parse_graphic_frame(node, node_index, context)?)),
        "nvGrpSpPr" | "grpSpPr" | "extLst" => Ok(None),
        feature => {
            context.diagnostics.warn(
                "unsupported-feature",
                format!("Slide-tree node {feature} is not supported."),
                Some(feature),
            );
            Ok(None)
        }
    }
}

fn part_background(root: &XmlNode, colors: Option<&ColorContext<'_>>) -> Option<FillStyle> {
    let background = root.child("cSld")?.child("bg")?;
    background
        .child("bgPr")
        .and_then(|properties| parse_fill(properties, colors))
        .or_else(|| {
            background
                .child("bgRef")
                .and_then(|reference| parse_color(reference, colors))
                .map(|color| FillStyle::Solid { color })
        })
}

struct ParsedSlideContent {
    name: Option<String>,
    hidden: Option<bool>,
    background: Option<FillStyle>,
    nodes: Vec<SlideNode>,
}

fn parse_part_nodes(
    root: &XmlNode,
    context: &NodeParseContext<'_, '_>,
) -> Result<Vec<SlideNode>, ParseError> {
    let mut nodes = Vec::new();
    let Some(shape_tree) = shape_tree(root) else {
        return Ok(nodes);
    };
    for (node_index, node) in shape_tree.children.iter().enumerate() {
        if let Some(node) = parse_slide_node(node, node_index, context)? {
            nodes.push(node);
        }
    }
    Ok(nodes)
}

fn parse_slide(
    root: &XmlNode,
    context: &NodeParseContext<'_, '_>,
) -> Result<ParsedSlideContent, ParseError> {
    let background = [Some(root), context.layout_root, context.master_root]
        .into_iter()
        .flatten()
        .find_map(|part| part_background(part, context.colors));
    Ok(ParsedSlideContent {
        name: root
            .child("cSld")
            .and_then(|slide| slide.attr("name"))
            .map(str::to_owned),
        hidden: root
            .attr("show")
            .and_then(bool_value)
            .map(|show| !show)
            .filter(|hidden| *hidden),
        background,
        nodes: parse_part_nodes(root, context)?,
    })
}

fn parse_core_metadata(xml: &[u8]) -> PresentationMetadata {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut field: Option<String> = None;
    let mut metadata = PresentationMetadata::default();
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                field =
                    Some(String::from_utf8_lossy(local_name(start.name().as_ref())).into_owned())
            }
            Ok(Event::Text(text)) => {
                if let Some(name) = field.as_deref() {
                    let value = text
                        .unescape()
                        .map(|text| text.into_owned())
                        .unwrap_or_default();
                    match name {
                        "title" => metadata.title = Some(value),
                        "subject" => metadata.subject = Some(value),
                        "creator" => metadata.creator = Some(value),
                        _ => {}
                    }
                }
            }
            Ok(Event::End(_)) => field = None,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    metadata
}

fn parse_text_values(xml: &[u8]) -> Vec<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut in_text = false;
    let mut values = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                in_text = matches!(local_name(start.name().as_ref()), b"t" | b"text")
            }
            Ok(Event::Text(text)) if in_text => {
                if let Ok(value) = text.unescape() {
                    let value = value.trim();
                    if !value.is_empty() {
                        values.push(value.to_owned());
                    }
                }
            }
            Ok(Event::End(end)) if matches!(local_name(end.name().as_ref()), b"t" | b"text") => {
                in_text = false
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    values
}

fn parse_comments(xml: &[u8]) -> Vec<SlideComment> {
    parse_text_values(xml)
        .into_iter()
        .enumerate()
        .map(|(index, text)| SlideComment {
            id: format!("comment-{}", index + 1),
            text,
            author: None,
            created_at: None,
            x: None,
            y: None,
        })
        .collect()
}

fn content_type(path: &str) -> String {
    match path
        .rsplit('.')
        .next()
        .unwrap_or_default()
        .to_ascii_lowercase()
        .as_str()
    {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "svg" => "image/svg+xml",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        "mp4" => "video/mp4",
        "mp3" => "audio/mpeg",
        _ => "application/octet-stream",
    }
    .into()
}

#[derive(Clone)]
struct ParsedXmlPart {
    path: String,
    root: XmlNode,
    rels: HashMap<String, String>,
}

fn relationship_part_path(path: &str) -> String {
    match path.rsplit_once('/') {
        Some((directory, file_name)) => format!("{directory}/_rels/{file_name}.rels"),
        None => format!("_rels/{path}.rels"),
    }
}

fn load_xml_part(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    path: &str,
    limits: &ParseLimits,
) -> Result<Option<ParsedXmlPart>, ParseError> {
    let Some(xml) = read_entry(archive, path, limits)? else {
        return Ok(None);
    };
    let root = parse_xml_tree(&xml, limits, path)?;
    let rels = match read_entry(archive, &relationship_part_path(path), limits)? {
        Some(xml) => relationships(&xml, path, limits.max_xml_depth)?,
        None => HashMap::new(),
    };
    Ok(Some(ParsedXmlPart {
        path: path.to_owned(),
        root,
        rels,
    }))
}

fn related_path(rels: &HashMap<String, String>, prefix: &str) -> Option<String> {
    rels.values()
        .find(|target| target.starts_with(prefix))
        .cloned()
}

fn missing_part_warning(path: &str, slide_index: Option<usize>) -> PresentationWarning {
    let mut warning = PresentationWarning::warning(
        "missing-part",
        format!("Related presentation part {path} is missing."),
    );
    warning.slide_index = slide_index;
    warning.part_name = Some(path.to_owned());
    warning
}

pub fn parse(bytes: &[u8], limits: &ParseLimits) -> Result<PresentationDocument, ParseError> {
    let cursor = Cursor::new(bytes);
    let mut archive =
        ZipArchive::new(cursor).map_err(|error| ParseError::Corrupt(format!("ZIP: {error}")))?;
    if archive.len() > limits.max_entries {
        return Err(ParseError::ResourceLimit("ZIP entry count".into()));
    }
    let total: u64 = (0..archive.len())
        .filter_map(|index| archive.by_index(index).ok().map(|file| file.size()))
        .sum();
    if total > limits.max_total_uncompressed_bytes as u64 {
        return Err(ParseError::ResourceLimit("ZIP expanded size".into()));
    }
    if archive.by_name("EncryptedPackage").is_ok() {
        return Err(ParseError::Encrypted);
    }

    let presentation = read_entry(&mut archive, "ppt/presentation.xml", limits)?
        .ok_or_else(|| ParseError::Corrupt("missing ppt/presentation.xml".into()))?;
    let rel_bytes = read_entry(&mut archive, "ppt/_rels/presentation.xml.rels", limits)?
        .ok_or_else(|| ParseError::Corrupt("missing presentation relationships".into()))?;
    let rels = relationships(&rel_bytes, "ppt/presentation.xml", limits.max_xml_depth)?;
    let (size, slide_paths) = parse_presentation_xml(&presentation, &rels, limits)?;
    let embedded_font_refs = parse_embedded_font_references(&presentation);
    if slide_paths.is_empty() {
        return Err(ParseError::Corrupt(
            "presentation has no related slides".into(),
        ));
    }

    let mut slides = Vec::with_capacity(slide_paths.len());
    let mut asset_paths = BTreeMap::<String, String>::new();
    let mut layout_parts = BTreeMap::<String, ParsedXmlPart>::new();
    let mut master_parts = BTreeMap::<String, ParsedXmlPart>::new();
    let mut theme_parts = BTreeMap::<String, ThemeDataInternal>::new();
    let mut document_warnings = Vec::new();
    for (index, slide_path) in slide_paths.iter().enumerate() {
        let xml = read_entry(&mut archive, slide_path, limits)?
            .ok_or_else(|| ParseError::Corrupt(format!("missing {slide_path}")))?;
        let slide_root = parse_xml_tree(&xml, limits, slide_path)?;
        let slide_rels =
            match read_entry(&mut archive, &relationship_part_path(slide_path), limits)? {
                Some(bytes) => relationships(&bytes, slide_path, limits.max_xml_depth)?,
                None => HashMap::new(),
            };

        let layout_path = related_path(&slide_rels, "ppt/slideLayouts/");
        if let Some(path) = layout_path.as_deref() {
            if !layout_parts.contains_key(path) {
                match load_xml_part(&mut archive, path, limits)? {
                    Some(part) => {
                        layout_parts.insert(path.to_owned(), part);
                    }
                    None => document_warnings.push(missing_part_warning(path, Some(index))),
                }
            }
        }
        let layout = layout_path
            .as_ref()
            .and_then(|path| layout_parts.get(path))
            .cloned();
        let master_path = layout
            .as_ref()
            .and_then(|part| related_path(&part.rels, "ppt/slideMasters/"));
        if let Some(path) = master_path.as_deref() {
            if !master_parts.contains_key(path) {
                match load_xml_part(&mut archive, path, limits)? {
                    Some(part) => {
                        master_parts.insert(path.to_owned(), part);
                    }
                    None => document_warnings.push(missing_part_warning(path, Some(index))),
                }
            }
        }
        let master = master_path
            .as_ref()
            .and_then(|path| master_parts.get(path))
            .cloned();
        let theme_path = master
            .as_ref()
            .and_then(|part| related_path(&part.rels, "ppt/theme/"));
        if let Some(path) = theme_path.as_deref() {
            if !theme_parts.contains_key(path) {
                match read_entry(&mut archive, path, limits)? {
                    Some(xml) => {
                        theme_parts.insert(path.to_owned(), parse_theme(&xml, path, limits)?);
                    }
                    None => document_warnings.push(missing_part_warning(path, Some(index))),
                }
            }
        }
        let theme = theme_path
            .as_ref()
            .and_then(|path| theme_parts.get(path))
            .cloned();

        let related_paths = slide_rels
            .values()
            .filter(|target| target.starts_with("ppt/charts/"))
            .cloned()
            .collect::<Vec<_>>();
        let mut related_parts = HashMap::new();
        for target in related_paths {
            if let Some(bytes) = read_entry(&mut archive, &target, limits)? {
                related_parts.insert(target, bytes);
            }
        }
        for rels in [
            Some(&slide_rels),
            layout.as_ref().map(|part| &part.rels),
            master.as_ref().map(|part| &part.rels),
        ]
        .into_iter()
        .flatten()
        {
            for target in rels
                .values()
                .filter(|target| target.starts_with("ppt/media/"))
            {
                asset_paths.insert(target.clone(), target.clone());
            }
        }
        let notes = if let Some(target) = slide_rels
            .values()
            .find(|target| target.starts_with("ppt/notesSlides/"))
        {
            read_entry(&mut archive, target, limits)?
                .map(|xml| {
                    let text = parse_text_values(&xml).join("\n");
                    if text.is_empty() {
                        Vec::new()
                    } else {
                        vec![SlideNote { text }]
                    }
                })
                .unwrap_or_default()
        } else {
            Vec::new()
        };
        let comments = if let Some(target) = slide_rels
            .values()
            .find(|target| target.starts_with("ppt/comments/"))
        {
            read_entry(&mut archive, target, limits)?
                .map(|xml| parse_comments(&xml))
                .unwrap_or_default()
        } else {
            Vec::new()
        };
        let master_color_map = master
            .as_ref()
            .map(|part| color_map(&part.root))
            .unwrap_or_default();
        let layout_color_map = layout
            .as_ref()
            .map(|part| merge_color_maps(&master_color_map, &color_map(&part.root)))
            .unwrap_or_else(|| master_color_map.clone());
        let slide_color_map = merge_color_maps(&layout_color_map, &color_map(&slide_root));
        let diagnostics = ParseDiagnostics::new(Some(index), slide_path);
        let content = {
            let colors = ColorContext {
                theme: theme.as_ref(),
                color_map: &slide_color_map,
                diagnostics: &diagnostics,
            };
            let context = NodeParseContext {
                slide_index: index,
                rels: &slide_rels,
                related_parts: &related_parts,
                limits,
                colors: Some(&colors),
                diagnostics: &diagnostics,
                layout_root: layout.as_ref().map(|part| &part.root),
                master_root: master.as_ref().map(|part| &part.root),
            };
            parse_slide(&slide_root, &context)?
        };
        let slide_warnings = diagnostics.take();
        slides.push(PresentationSlide {
            id: format!("slide-{}", index + 1),
            index,
            name: content.name,
            hidden: content.hidden,
            master_id: master_path,
            layout_id: layout_path,
            background: content.background,
            nodes: content.nodes,
            notes,
            comments,
            source_part: Some(slide_path.clone()),
            warnings: slide_warnings,
        });
    }

    let declared_master_paths = rels
        .values()
        .filter(|target| target.starts_with("ppt/slideMasters/"))
        .cloned()
        .collect::<Vec<_>>();
    for path in declared_master_paths {
        if master_parts.contains_key(&path) {
            continue;
        }
        match load_xml_part(&mut archive, &path, limits)? {
            Some(part) => {
                master_parts.insert(path, part);
            }
            None => document_warnings.push(missing_part_warning(&path, None)),
        }
    }
    let declared_layout_paths = master_parts
        .values()
        .flat_map(|part| part.rels.values())
        .filter(|target| target.starts_with("ppt/slideLayouts/"))
        .cloned()
        .collect::<Vec<_>>();
    for path in declared_layout_paths {
        if layout_parts.contains_key(&path) {
            continue;
        }
        match load_xml_part(&mut archive, &path, limits)? {
            Some(part) => {
                layout_parts.insert(path, part);
            }
            None => document_warnings.push(missing_part_warning(&path, None)),
        }
    }
    let declared_theme_paths = master_parts
        .values()
        .flat_map(|part| part.rels.values())
        .filter(|target| target.starts_with("ppt/theme/"))
        .cloned()
        .collect::<Vec<_>>();
    for path in declared_theme_paths {
        if theme_parts.contains_key(&path) {
            continue;
        }
        match read_entry(&mut archive, &path, limits)? {
            Some(xml) => {
                theme_parts.insert(path.clone(), parse_theme(&xml, &path, limits)?);
            }
            None => document_warnings.push(missing_part_warning(&path, None)),
        }
    }
    for target in master_parts
        .values()
        .chain(layout_parts.values())
        .flat_map(|part| part.rels.values())
        .filter(|target| target.starts_with("ppt/media/"))
    {
        asset_paths.insert(target.clone(), target.clone());
    }

    let empty_related_parts = HashMap::new();
    let mut masters = Vec::new();
    for part in master_parts.values() {
        let theme_path = related_path(&part.rels, "ppt/theme/");
        let theme = theme_path.as_ref().and_then(|path| theme_parts.get(path));
        let part_color_map = color_map(&part.root);
        let diagnostics = ParseDiagnostics::new(None, &part.path);
        let nodes = {
            let colors = ColorContext {
                theme,
                color_map: &part_color_map,
                diagnostics: &diagnostics,
            };
            let context = NodeParseContext {
                slide_index: 0,
                rels: &part.rels,
                related_parts: &empty_related_parts,
                limits,
                colors: Some(&colors),
                diagnostics: &diagnostics,
                layout_root: None,
                master_root: None,
            };
            parse_part_nodes(&part.root, &context)?
        };
        document_warnings.extend(diagnostics.take());
        masters.push(crate::SlideMaster {
            id: part.path.clone(),
            name: part
                .root
                .child("cSld")
                .and_then(|slide| slide.attr("name"))
                .map(str::to_owned),
            theme_id: theme_path,
            nodes,
        });
    }

    let mut layouts = Vec::new();
    for part in layout_parts.values() {
        let master_path = related_path(&part.rels, "ppt/slideMasters/");
        let master = master_path.as_ref().and_then(|path| master_parts.get(path));
        let theme = master
            .and_then(|part| related_path(&part.rels, "ppt/theme/"))
            .and_then(|path| theme_parts.get(&path));
        let master_color_map = master.map(|part| color_map(&part.root)).unwrap_or_default();
        let part_color_map = merge_color_maps(&master_color_map, &color_map(&part.root));
        let diagnostics = ParseDiagnostics::new(None, &part.path);
        let nodes = {
            let colors = ColorContext {
                theme,
                color_map: &part_color_map,
                diagnostics: &diagnostics,
            };
            let context = NodeParseContext {
                slide_index: 0,
                rels: &part.rels,
                related_parts: &empty_related_parts,
                limits,
                colors: Some(&colors),
                diagnostics: &diagnostics,
                layout_root: None,
                master_root: master.map(|part| &part.root),
            };
            parse_part_nodes(&part.root, &context)?
        };
        document_warnings.extend(diagnostics.take());
        layouts.push(crate::SlideLayout {
            id: part.path.clone(),
            name: part
                .root
                .child("cSld")
                .and_then(|slide| slide.attr("name"))
                .map(str::to_owned),
            master_id: master_path,
            nodes,
        });
    }

    let mut assets = BTreeMap::new();
    for asset_path in asset_paths.keys() {
        if let Some(data) = read_entry(&mut archive, asset_path, limits)? {
            assets.insert(
                asset_path.clone(),
                PresentationAsset {
                    id: asset_path.clone(),
                    content_type: content_type(asset_path),
                    byte_length: data.len(),
                    file_name: asset_path.rsplit('/').next().map(str::to_owned),
                    data: Some(data.into()),
                },
            );
        }
    }
    let mut embedded_fonts = Vec::new();
    let mut font_warnings = Vec::new();
    for font in embedded_font_refs {
        for (rel_id, style, weight) in font.variants {
            let Some(path) = rels.get(&rel_id) else {
                font_warnings.push(PresentationWarning::warning(
                    "missing-font",
                    format!(
                        "Embedded font relationship {rel_id} for {} is missing.",
                        font.family
                    ),
                ));
                continue;
            };
            let Some(mut data) = read_entry(&mut archive, path, limits)? else {
                font_warnings.push(PresentationWarning::warning(
                    "missing-font",
                    format!("Embedded font data for {} is missing.", font.family),
                ));
                continue;
            };
            if !deobfuscate_embedded_font(path, &mut data) {
                font_warnings.push(PresentationWarning::warning(
                    "degraded-rendering",
                    format!("Embedded font {} could not be decoded safely.", font.family),
                ));
                continue;
            }
            let content_type = font_magic(&data)
                .unwrap_or("application/octet-stream")
                .to_owned();
            assets.insert(
                path.clone(),
                PresentationAsset {
                    id: path.clone(),
                    content_type,
                    byte_length: data.len(),
                    file_name: path.rsplit('/').next().map(str::to_owned),
                    data: Some(data.into()),
                },
            );
            embedded_fonts.push(PresentationEmbeddedFont {
                family: font.family.clone(),
                asset_id: path.clone(),
                style,
                weight,
            });
        }
    }
    let metadata =
        read_entry(&mut archive, "docProps/core.xml", limits)?.map(|xml| parse_core_metadata(&xml));
    document_warnings.extend(font_warnings);
    Ok(PresentationDocument {
        format: PresentationFormat::Pptx,
        size,
        slides,
        masters,
        layouts,
        themes: theme_parts
            .values()
            .map(ThemeDataInternal::public)
            .collect(),
        assets,
        embedded_fonts,
        warnings: document_warnings,
        metadata,
    })
}

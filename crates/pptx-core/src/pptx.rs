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
    // OPC relationship targets may be package-absolute ("/ppt/charts/chart1.xml").
    let (base, target) = match target.strip_prefix('/') {
        Some(absolute) => ("", absolute),
        None => (base, target),
    };
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

const CHART_STYLE_REL_TYPE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
const CHART_COLOR_STYLE_REL_TYPE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

fn is_chart_relationship_type(relationship_type: &str) -> bool {
    matches!(
        relationship_type.rsplit('/').next(),
        Some("chart" | "chartEx")
    )
}

fn is_diagram_relationship_type(relationship_type: &str) -> bool {
    relationship_type
        .rsplit('/')
        .next()
        .is_some_and(|name| name.starts_with("diagram"))
}

/// Companion Microsoft chart style parts resolved from a chart part's rels.
#[derive(Debug, Clone, Default)]
struct ChartCompanionParts {
    style_xml: Option<String>,
    colors_xml: Option<String>,
}

fn relationship_targets_by_type(
    xml: &[u8],
    base: &str,
    max_depth: usize,
) -> Result<Vec<(String, String)>, ParseError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut result = Vec::new();
    let mut depth = 0usize;
    loop {
        let event = reader.read_event();
        match event {
            Ok(Event::Start(ref start)) | Ok(Event::Empty(ref start)) => {
                if matches!(event, Ok(Event::Start(_))) {
                    depth += 1;
                    if depth > max_depth {
                        return Err(ParseError::ResourceLimit("XML nesting depth".into()));
                    }
                }
                if local_name(start.name().as_ref()) == b"Relationship" {
                    if let (Some(rel_type), Some(target)) =
                        (attr(start, b"Type"), attr(start, b"Target"))
                    {
                        result.push((rel_type, normalize_target(base, &target)));
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
                                .filter(|value| rels.contains_key(value))
                                .or_else(|| {
                                    start.attributes().with_checks(false).flatten().find_map(
                                        |value| {
                                            let candidate =
                                                String::from_utf8_lossy(value.value.as_ref())
                                                    .into_owned();
                                            (local_name(value.key.as_ref()) == b"id"
                                                && rels.contains_key(&candidate))
                                            .then_some(candidate)
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
                        .filter(|value| rels.contains_key(value))
                        .or_else(|| {
                            start
                                .attributes()
                                .with_checks(false)
                                .flatten()
                                .find_map(|value| {
                                    let candidate =
                                        String::from_utf8_lossy(value.value.as_ref()).into_owned();
                                    (local_name(value.key.as_ref()) == b"id"
                                        && rels.contains_key(&candidate))
                                    .then_some(candidate)
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
        "srgbClr" | "scrgbClr" | "hslClr" | "schemeClr" | "sysClr" | "prstClr"
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
    rels: HashMap<String, String>,
    fill_styles: Vec<XmlNode>,
    background_fill_styles: Vec<XmlNode>,
    line_styles: Vec<XmlNode>,
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

fn percentage(value: &str) -> Option<f64> {
    let value = value.trim();
    if let Some(value) = value.strip_suffix('%') {
        return value.trim().parse::<f64>().ok().map(|value| value / 100.0);
    }
    value.parse::<f64>().ok().map(|value| value / 100_000.0)
}

fn rgb_to_hsl(rgb: [f64; 3]) -> [f64; 3] {
    let [red, green, blue] = rgb.map(|channel| channel / 255.0);
    let maximum = red.max(green).max(blue);
    let minimum = red.min(green).min(blue);
    let luminance = (maximum + minimum) / 2.0;
    if (maximum - minimum).abs() < f64::EPSILON {
        return [0.0, 0.0, luminance];
    }
    let delta = maximum - minimum;
    let saturation = if luminance > 0.5 {
        delta / (2.0 - maximum - minimum)
    } else {
        delta / (maximum + minimum)
    };
    let hue = if (maximum - red).abs() < f64::EPSILON {
        (green - blue) / delta + if green < blue { 6.0 } else { 0.0 }
    } else if (maximum - green).abs() < f64::EPSILON {
        (blue - red) / delta + 2.0
    } else {
        (red - green) / delta + 4.0
    } / 6.0;
    [hue, saturation, luminance]
}

fn hue_channel(p: f64, q: f64, mut value: f64) -> f64 {
    if value < 0.0 {
        value += 1.0;
    }
    if value > 1.0 {
        value -= 1.0;
    }
    if value < 1.0 / 6.0 {
        p + (q - p) * 6.0 * value
    } else if value < 0.5 {
        q
    } else if value < 2.0 / 3.0 {
        p + (q - p) * (2.0 / 3.0 - value) * 6.0
    } else {
        p
    }
}

fn hsl_to_rgb([hue, saturation, luminance]: [f64; 3]) -> [f64; 3] {
    if saturation.abs() < f64::EPSILON {
        return [luminance * 255.0; 3];
    }
    let q = if luminance < 0.5 {
        luminance * (1.0 + saturation)
    } else {
        luminance + saturation - luminance * saturation
    };
    let p = 2.0 * luminance - q;
    [
        hue_channel(p, q, hue + 1.0 / 3.0) * 255.0,
        hue_channel(p, q, hue) * 255.0,
        hue_channel(p, q, hue - 1.0 / 3.0) * 255.0,
    ]
}

fn angle_degrees(value: &str) -> Option<f64> {
    value.parse::<f64>().ok().map(|value| value / 60_000.0)
}

fn srgb_channel_to_linear(channel: f64) -> f64 {
    let value = (channel / 255.0).clamp(0.0, 1.0);
    if value <= 0.04045 {
        value / 12.92
    } else {
        ((value + 0.055) / 1.055).powf(2.4)
    }
}

fn linear_channel_to_srgb(channel: f64) -> f64 {
    let value = channel.clamp(0.0, 1.0);
    let encoded = if value <= 0.003_130_8 {
        value * 12.92
    } else {
        1.055 * value.powf(1.0 / 2.4) - 0.055
    };
    encoded * 255.0
}

fn transformed_color(value: &str, color: &XmlNode) -> Option<String> {
    let mut rgb = rgb_from_hex(value)?;
    for transform in &color.children {
        let amount = transform.attr("val").and_then(percentage);
        match (transform.name.as_str(), amount) {
            ("tint", Some(amount)) => {
                for channel in &mut rgb {
                    let linear = srgb_channel_to_linear(*channel);
                    *channel = linear_channel_to_srgb(linear * amount + (1.0 - amount));
                }
            }
            ("shade", Some(amount)) => {
                for channel in &mut rgb {
                    *channel = linear_channel_to_srgb(srgb_channel_to_linear(*channel) * amount);
                }
            }
            ("lum", Some(amount)) => {
                let mut hsl = rgb_to_hsl(rgb);
                hsl[2] = amount.clamp(0.0, 1.0);
                rgb = hsl_to_rgb(hsl);
            }
            ("lumMod", Some(amount)) => {
                let mut hsl = rgb_to_hsl(rgb);
                hsl[2] = (hsl[2] * amount).clamp(0.0, 1.0);
                rgb = hsl_to_rgb(hsl);
            }
            ("lumOff", Some(amount)) => {
                let mut hsl = rgb_to_hsl(rgb);
                hsl[2] = (hsl[2] + amount).clamp(0.0, 1.0);
                rgb = hsl_to_rgb(hsl);
            }
            ("sat", Some(amount)) => {
                let mut hsl = rgb_to_hsl(rgb);
                hsl[1] = amount.clamp(0.0, 1.0);
                rgb = hsl_to_rgb(hsl);
            }
            ("satMod", Some(amount)) => {
                let mut hsl = rgb_to_hsl(rgb);
                hsl[1] = (hsl[1] * amount).clamp(0.0, 1.0);
                rgb = hsl_to_rgb(hsl);
            }
            ("satOff", Some(amount)) => {
                let mut hsl = rgb_to_hsl(rgb);
                hsl[1] = (hsl[1] + amount).clamp(0.0, 1.0);
                rgb = hsl_to_rgb(hsl);
            }
            ("hueMod", Some(amount)) => {
                let mut hsl = rgb_to_hsl(rgb);
                hsl[0] = (hsl[0] * amount).rem_euclid(1.0);
                rgb = hsl_to_rgb(hsl);
            }
            ("red", Some(amount)) => rgb[0] = 255.0 * amount,
            ("green", Some(amount)) => rgb[1] = 255.0 * amount,
            ("blue", Some(amount)) => rgb[2] = 255.0 * amount,
            ("redMod", Some(amount)) => rgb[0] *= amount,
            ("greenMod", Some(amount)) => rgb[1] *= amount,
            ("blueMod", Some(amount)) => rgb[2] *= amount,
            ("redOff", Some(amount)) => rgb[0] += 255.0 * amount,
            ("greenOff", Some(amount)) => rgb[1] += 255.0 * amount,
            ("blueOff", Some(amount)) => rgb[2] += 255.0 * amount,
            ("hue", _) => {
                if let Some(value) = transform.attr("val").and_then(angle_degrees) {
                    let mut hsl = rgb_to_hsl(rgb);
                    hsl[0] = (value / 360.0).rem_euclid(1.0);
                    rgb = hsl_to_rgb(hsl);
                }
            }
            ("hueOff", _) => {
                if let Some(value) = transform.attr("val").and_then(angle_degrees) {
                    let mut hsl = rgb_to_hsl(rgb);
                    hsl[0] = (hsl[0] + value / 360.0).rem_euclid(1.0);
                    rgb = hsl_to_rgb(hsl);
                }
            }
            ("comp", _) => {
                let mut hsl = rgb_to_hsl(rgb);
                hsl[0] = (hsl[0] + 0.5).rem_euclid(1.0);
                rgb = hsl_to_rgb(hsl);
            }
            ("inv", _) => {
                rgb = rgb.map(|channel| 255.0 - channel);
            }
            ("gray", _) => {
                let gray = rgb[0] * 0.2126 + rgb[1] * 0.7152 + rgb[2] * 0.0722;
                rgb = [gray; 3];
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

fn parse_color_with_placeholder(
    node: &XmlNode,
    context: Option<&ColorContext<'_>>,
    placeholder: Option<&ColorValue>,
) -> Option<ColorValue> {
    let color = color_element(node)?;
    let inherited_alpha = (color.name == "schemeClr" && color.attr("val") == Some("phClr"))
        .then(|| placeholder.and_then(|value| value.alpha))
        .flatten()
        .unwrap_or(1.0);
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
        "hslClr" => {
            let hue = color.attr("hue").and_then(angle_degrees).unwrap_or(0.0) / 360.0;
            let saturation = color.attr("sat").and_then(percentage).unwrap_or(0.0);
            let luminance = color.attr("lum").and_then(percentage).unwrap_or(0.0);
            let rgb = hsl_to_rgb([hue, saturation, luminance]);
            Some(format!(
                "#{:02X}{:02X}{:02X}",
                rgb[0].round().clamp(0.0, 255.0) as u8,
                rgb[1].round().clamp(0.0, 255.0) as u8,
                rgb[2].round().clamp(0.0, 255.0) as u8
            ))
        }
        "schemeClr" => {
            let requested = color.attr("val")?;
            if requested == "phClr" {
                placeholder.map(|value| value.value.clone())
            } else {
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
        }
        "sysClr" => color
            .attr("lastClr")
            .and_then(normalized_hex)
            .or_else(|| color.attr("val").and_then(preset_color).map(str::to_owned)),
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
        .unwrap_or(inherited_alpha);
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

fn parse_color(node: &XmlNode, context: Option<&ColorContext<'_>>) -> Option<ColorValue> {
    parse_color_with_placeholder(node, context, None)
}

fn parse_theme(
    xml: &[u8],
    path: &str,
    rels: HashMap<String, String>,
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
    let format_scheme = root.descendant("fmtScheme");
    Ok(ThemeDataInternal {
        id: path.to_owned(),
        name: root
            .attr("name")
            .or_else(|| color_scheme.and_then(|scheme| scheme.attr("name")))
            .map(str::to_owned),
        colors,
        major_fonts: parse_fonts(font_scheme.and_then(|scheme| scheme.child("majorFont"))),
        minor_fonts: parse_fonts(font_scheme.and_then(|scheme| scheme.child("minorFont"))),
        rels,
        fill_styles: format_scheme
            .and_then(|scheme| scheme.child("fillStyleLst"))
            .map(|list| list.children.clone())
            .unwrap_or_default(),
        background_fill_styles: format_scheme
            .and_then(|scheme| scheme.child("bgFillStyleLst"))
            .map(|list| list.children.clone())
            .unwrap_or_default(),
        line_styles: format_scheme
            .and_then(|scheme| scheme.child("lnStyleLst"))
            .map(|list| list.children.clone())
            .unwrap_or_default(),
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

fn named_fill<'a>(properties: &'a XmlNode, name: &str) -> Option<&'a XmlNode> {
    (properties.name == name)
        .then_some(properties)
        .or_else(|| properties.child(name))
}

fn image_crop(fill: &XmlNode) -> Option<crate::ImageCrop> {
    let source = fill.child("srcRect")?;
    let edge = |name: &str| {
        source
            .attr(name)
            .and_then(percentage)
            .unwrap_or(0.0)
            .clamp(-1.0, 1.0)
    };
    let crop = crate::ImageCrop {
        top: edge("t"),
        right: edge("r"),
        bottom: edge("b"),
        left: edge("l"),
    };
    (crop.top != 0.0 || crop.right != 0.0 || crop.bottom != 0.0 || crop.left != 0.0).then_some(crop)
}

/// Parse DrawingML recolor effects declared directly on an `a:blip`.
fn parse_blip_effects(
    blip: &XmlNode,
    context: Option<&ColorContext<'_>>,
) -> Option<crate::ImageEffects> {
    let mut effects = crate::ImageEffects::default();
    let mut any = false;
    if let Some(node) = blip.child("biLevel") {
        effects.bi_level_threshold = Some(
            node.attr("thresh")
                .and_then(percentage)
                .map(|value| value.clamp(0.0, 1.0))
                .unwrap_or(0.5),
        );
        any = true;
    }
    if blip.child("grayscl").is_some() {
        effects.grayscale = Some(true);
        any = true;
    }
    if let Some(node) = blip.child("duotone") {
        let colors = node
            .children
            .iter()
            .filter_map(|child| parse_color(child, context))
            .collect::<Vec<_>>();
        if colors.len() >= 2 {
            effects.duotone = Some(colors.into_iter().take(2).collect());
            any = true;
        }
    }
    if let Some(node) = blip.child("lum") {
        let brightness = node
            .attr("bright")
            .and_then(percentage)
            .map(|value| value.clamp(-1.0, 1.0));
        let contrast = node
            .attr("contrast")
            .and_then(percentage)
            .map(|value| value.clamp(-1.0, 1.0));
        if brightness.is_some() || contrast.is_some() {
            effects.brightness = brightness;
            effects.contrast = contrast;
            any = true;
        }
    }
    any.then_some(effects)
}

fn parse_blip_fill(
    fill: &XmlNode,
    context: Option<&ColorContext<'_>>,
    rels: &HashMap<String, String>,
) -> Option<FillStyle> {
    let blip = fill.child("blip")?;
    let relationship = blip.attr("embed").or_else(|| blip.attr("link"))?;
    let asset_id = rels.get(relationship)?.clone();
    let opacity = blip
        .child("alphaModFix")
        .or_else(|| blip.child("alphaMod"))
        .and_then(|effect| effect.attr("amt").or_else(|| effect.attr("val")))
        .and_then(percentage)
        .map(|value| value.clamp(0.0, 1.0));
    Some(FillStyle::Image {
        asset_id,
        mode: if fill.child("tile").is_some() {
            crate::FillImageMode::Tile
        } else {
            crate::FillImageMode::Stretch
        },
        crop: image_crop(fill),
        opacity,
        effects: parse_blip_effects(blip, context),
    })
}

fn parse_fill_with_placeholder(
    properties: &XmlNode,
    context: Option<&ColorContext<'_>>,
    rels: &HashMap<String, String>,
    placeholder: Option<&ColorValue>,
) -> Option<FillStyle> {
    if named_fill(properties, "noFill").is_some() {
        return Some(FillStyle::None);
    }
    if let Some(solid) = named_fill(properties, "solidFill") {
        return parse_color_with_placeholder(solid, context, placeholder)
            .map(|color| FillStyle::Solid { color });
    }
    if let Some(gradient) = named_fill(properties, "gradFill") {
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
                    color: parse_color_with_placeholder(stop, context, placeholder)?,
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
    if let Some(pattern) = named_fill(properties, "pattFill") {
        let foreground =
            parse_color_with_placeholder(pattern.child("fgClr")?, context, placeholder)?;
        let background =
            parse_color_with_placeholder(pattern.child("bgClr")?, context, placeholder)?;
        return Some(FillStyle::Pattern {
            preset: pattern.attr("prst").unwrap_or("pct5").to_owned(),
            foreground,
            background,
        });
    }
    if let Some(fill) = named_fill(properties, "blipFill") {
        return parse_blip_fill(fill, context, rels);
    }
    None
}

fn parse_fill(
    properties: &XmlNode,
    context: Option<&ColorContext<'_>>,
    rels: &HashMap<String, String>,
) -> Option<FillStyle> {
    parse_fill_with_placeholder(properties, context, rels, None)
}

fn parse_line_node_with_placeholder(
    line: &XmlNode,
    context: Option<&ColorContext<'_>>,
    placeholder: Option<&ColorValue>,
) -> Option<LineStyle> {
    if line.child("noFill").is_some() {
        return None;
    }
    let color = line
        .child("solidFill")
        .and_then(|fill| parse_color_with_placeholder(fill, context, placeholder))
        .or_else(|| {
            line.child("gradFill").and_then(|gradient| {
                gradient
                    .child("gsLst")
                    .and_then(|list| list.child("gs"))
                    .and_then(|stop| parse_color_with_placeholder(stop, context, placeholder))
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

fn parse_line_node(line: &XmlNode, context: Option<&ColorContext<'_>>) -> Option<LineStyle> {
    parse_line_node_with_placeholder(line, context, None)
}

fn style_matrix_fill(reference: &XmlNode, context: Option<&ColorContext<'_>>) -> Option<FillStyle> {
    let index = reference.attr("idx")?.parse::<usize>().ok()?;
    if matches!(index, 0 | 1000) {
        return Some(FillStyle::None);
    }
    let theme = context?.theme?;
    let placeholder = parse_color(reference, context);
    let style = if index >= 1001 {
        theme.background_fill_styles.get(index - 1001)
    } else {
        theme.fill_styles.get(index - 1)
    };
    style
        .and_then(|style| {
            parse_fill_with_placeholder(style, context, &theme.rels, placeholder.as_ref())
        })
        .or_else(|| placeholder.map(|color| FillStyle::Solid { color }))
}

fn style_matrix_line(reference: &XmlNode, context: Option<&ColorContext<'_>>) -> Option<LineStyle> {
    let index = reference.attr("idx")?.parse::<usize>().ok()?;
    if index == 0 {
        return None;
    }
    let theme = context?.theme?;
    let placeholder = parse_color(reference, context);
    theme
        .line_styles
        .get(index - 1)
        .and_then(|line| parse_line_node_with_placeholder(line, context, placeholder.as_ref()))
        .or_else(|| {
            placeholder.map(|color| LineStyle {
                color: Some(color),
                ..Default::default()
            })
        })
}

#[derive(Debug, Clone, Default)]
struct RunStyle {
    font_family: Option<String>,
    east_asian_font_family: Option<String>,
    complex_script_font_family: Option<String>,
    symbol_font_family: Option<String>,
    font_size_pt: Option<f64>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    color: Option<ColorValue>,
    baseline: Option<f64>,
    language: Option<String>,
    alternative_language: Option<String>,
    right_to_left: Option<bool>,
    hyperlink: Option<String>,
    character_spacing_pt: Option<f64>,
    kerning_threshold_pt: Option<f64>,
}

impl RunStyle {
    fn overlay(mut self, other: RunStyle) -> Self {
        if other.font_family.is_some() {
            self.font_family = other.font_family;
        }
        if other.east_asian_font_family.is_some() {
            self.east_asian_font_family = other.east_asian_font_family;
        }
        if other.complex_script_font_family.is_some() {
            self.complex_script_font_family = other.complex_script_font_family;
        }
        if other.symbol_font_family.is_some() {
            self.symbol_font_family = other.symbol_font_family;
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
        if other.alternative_language.is_some() {
            self.alternative_language = other.alternative_language;
        }
        if other.right_to_left.is_some() {
            self.right_to_left = other.right_to_left;
        }
        if other.hyperlink.is_some() {
            self.hyperlink = other.hyperlink;
        }
        if other.character_spacing_pt.is_some() {
            self.character_spacing_pt = other.character_spacing_pt;
        }
        if other.kerning_threshold_pt.is_some() {
            self.kerning_threshold_pt = other.kerning_threshold_pt;
        }
        self
    }
}

fn run_font_family(
    properties: &XmlNode,
    element: &str,
    colors: Option<&ColorContext<'_>>,
) -> Option<String> {
    properties
        .child(element)
        .and_then(|node| node.attr("typeface"))
        .and_then(|typeface| {
            colors
                .and_then(|context| context.theme)
                .and_then(|theme| theme.resolve_font(typeface))
                .or_else(|| (!typeface.starts_with('+')).then(|| typeface.to_owned()))
        })
}

fn parse_run_style(
    properties: Option<&XmlNode>,
    colors: Option<&ColorContext<'_>>,
    rels: &HashMap<String, String>,
) -> RunStyle {
    let Some(properties) = properties else {
        return RunStyle::default();
    };
    let font_family = run_font_family(properties, "latin", colors);
    let east_asian_font_family = run_font_family(properties, "ea", colors);
    let complex_script_font_family = run_font_family(properties, "cs", colors);
    let symbol_font_family = run_font_family(properties, "sym", colors);
    let hyperlink = properties
        .child("hlinkClick")
        .and_then(|link| link.attr("id"))
        .and_then(|id| rels.get(id))
        .cloned();
    RunStyle {
        font_family: font_family
            .clone()
            .or_else(|| east_asian_font_family.clone())
            .or_else(|| complex_script_font_family.clone()),
        east_asian_font_family,
        complex_script_font_family,
        symbol_font_family,
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
        alternative_language: properties.attr("altLang").map(str::to_owned),
        right_to_left: properties
            .child("rtl")
            .map(|node| node.attr("val").and_then(bool_value).unwrap_or(true)),
        hyperlink,
        character_spacing_pt: properties
            .attr("spc")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value / 100.0),
        kerning_threshold_pt: properties
            .attr("kern")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value / 100.0),
    }
}

fn text_run(text: String, style: RunStyle) -> TextRun {
    TextRun {
        text,
        font_family: style.font_family,
        east_asian_font_family: style.east_asian_font_family,
        complex_script_font_family: style.complex_script_font_family,
        symbol_font_family: style.symbol_font_family,
        font_size_pt: style.font_size_pt,
        bold: style.bold,
        italic: style.italic,
        underline: style.underline,
        strike: style.strike,
        color: style.color,
        baseline: style.baseline,
        language: style.language,
        alternative_language: style.alternative_language,
        right_to_left: style.right_to_left,
        hyperlink: style.hyperlink,
        character_spacing_pt: style.character_spacing_pt,
        kerning_threshold_pt: style.kerning_threshold_pt,
    }
}

#[derive(Clone, Default)]
struct ParagraphStyle {
    alignment: Option<String>,
    level: Option<usize>,
    bullet: Option<Option<crate::TextBullet>>,
    line_spacing: Option<crate::TextSpacing>,
    space_before: Option<crate::TextSpacing>,
    space_after: Option<crate::TextSpacing>,
    rtl: Option<bool>,
    margin_left_emu: Option<i64>,
    indent_emu: Option<i64>,
    default_tab_size_emu: Option<i64>,
    tab_stops: Option<Vec<crate::TextTabStop>>,
    east_asian_line_break: Option<bool>,
    latin_line_break: Option<bool>,
    hanging_punctuation: Option<bool>,
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
        if other.margin_left_emu.is_some() {
            self.margin_left_emu = other.margin_left_emu;
        }
        if other.indent_emu.is_some() {
            self.indent_emu = other.indent_emu;
        }
        if other.default_tab_size_emu.is_some() {
            self.default_tab_size_emu = other.default_tab_size_emu;
        }
        if other.tab_stops.is_some() {
            self.tab_stops = other.tab_stops;
        }
        if other.east_asian_line_break.is_some() {
            self.east_asian_line_break = other.east_asian_line_break;
        }
        if other.latin_line_break.is_some() {
            self.latin_line_break = other.latin_line_break;
        }
        if other.hanging_punctuation.is_some() {
            self.hanging_punctuation = other.hanging_punctuation;
        }
        self
    }
}

fn spacing_value(node: Option<&XmlNode>) -> Option<crate::TextSpacing> {
    let node = node?;
    node.child("spcPts")
        .and_then(|value| value.attr("val"))
        .and_then(|value| value.parse::<f64>().ok())
        .map(|value| crate::TextSpacing {
            value: value / 100.0,
            unit: crate::TextSpacingUnit::Points,
        })
        .or_else(|| {
            node.child("spcPct")
                .and_then(|value| value.attr("val"))
                .and_then(percentage)
                .map(|value| crate::TextSpacing {
                    value,
                    unit: crate::TextSpacingUnit::Percent,
                })
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
    let bullet_font = properties
        .child("buFont")
        .and_then(|font| font.attr("typeface"))
        .map(str::to_owned);
    let bullet_size_pt = properties
        .child("buSzPts")
        .and_then(|size| size.attr("val"))
        .and_then(|value| value.parse::<f64>().ok())
        .map(|value| value / 100.0);
    let bullet_size_percent = properties
        .child("buSzPct")
        .and_then(|size| size.attr("val"))
        .and_then(percentage);
    let bullet = if let Some(character) = properties.child("buChar") {
        Some(Some(crate::TextBullet {
            kind: "character".into(),
            value: character.attr("char").map(str::to_owned),
            font_family: bullet_font.clone(),
            font_size_pt: bullet_size_pt,
            size_percent: bullet_size_percent,
            start_at: None,
        }))
    } else if let Some(number) = properties.child("buAutoNum") {
        Some(Some(crate::TextBullet {
            kind: "number".into(),
            value: number.attr("type").map(str::to_owned),
            font_family: bullet_font.clone(),
            font_size_pt: bullet_size_pt,
            size_percent: bullet_size_percent,
            start_at: number
                .attr("startAt")
                .and_then(|value| value.parse::<usize>().ok()),
        }))
    } else if properties.child("buBlip").is_some() {
        Some(Some(crate::TextBullet {
            kind: "picture".into(),
            value: None,
            font_family: bullet_font,
            font_size_pt: bullet_size_pt,
            size_percent: bullet_size_percent,
            start_at: None,
        }))
    } else if properties.child("buNone").is_some() {
        Some(None)
    } else {
        None
    };
    let tab_stops = properties.child("tabLst").map(|list| {
        list.children_named("tab")
            .filter_map(|tab| {
                let position_emu = tab.attr("pos")?.parse::<i64>().ok()?;
                let alignment = match tab.attr("algn") {
                    Some("ctr") => crate::TextTabAlignment::Center,
                    Some("r") => crate::TextTabAlignment::Right,
                    Some("dec") => crate::TextTabAlignment::Decimal,
                    _ => crate::TextTabAlignment::Left,
                };
                Some(crate::TextTabStop {
                    position_emu,
                    alignment,
                })
            })
            .collect()
    });
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
        margin_left_emu: properties
            .attr("marL")
            .and_then(|value| value.parse::<i64>().ok()),
        indent_emu: properties
            .attr("indent")
            .and_then(|value| value.parse::<i64>().ok()),
        default_tab_size_emu: properties
            .attr("defTabSz")
            .and_then(|value| value.parse::<i64>().ok()),
        tab_stops,
        east_asian_line_break: properties.attr("eaLnBrk").and_then(bool_value),
        latin_line_break: properties.attr("latinLnBrk").and_then(bool_value),
        hanging_punctuation: properties.attr("hangingPunct").and_then(bool_value),
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
    style.and_then(|style| {
        style
            .child(&format!("lvl{}pPr", level.saturating_add(1)))
            .or_else(|| style.child("defPPr"))
    })
}

struct TextParseContext<'a, 'color> {
    presentation_text_style: Option<&'a XmlNode>,
    shape_default_style: RunStyle,
    colors: Option<&'a ColorContext<'color>>,
    rels: &'a HashMap<String, String>,
}

fn parse_text_paragraphs(
    text_body: Option<&XmlNode>,
    layout_text_body: Option<&XmlNode>,
    master_text_body: Option<&XmlNode>,
    master_text_style: Option<&XmlNode>,
    context: &TextParseContext<'_, '_>,
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
            let presentation_style_properties =
                style_level_properties(context.presentation_text_style, level);
            let master_properties =
                fallback_paragraph_properties(master_text_body, paragraph_index, level);
            let layout_properties =
                fallback_paragraph_properties(layout_text_body, paragraph_index, level);
            let paragraph_style = parse_paragraph_style(presentation_style_properties)
                .overlay(parse_paragraph_style(master_style_properties))
                .overlay(parse_paragraph_style(master_properties))
                .overlay(parse_paragraph_style(layout_properties))
                .overlay(parse_paragraph_style(own_properties));
            // Presentation defaults are the base of the cascade. An explicit
            // shape or table-cell style must override them, not vice versa.
            let default_style = parse_run_style(
                presentation_style_properties.and_then(|properties| properties.child("defRPr")),
                context.colors,
                context.rels,
            )
            .overlay(context.shape_default_style.clone())
            .overlay(parse_run_style(
                master_style_properties.and_then(|properties| properties.child("defRPr")),
                context.colors,
                context.rels,
            ))
            .overlay(parse_run_style(
                master_properties.and_then(|properties| properties.child("defRPr")),
                context.colors,
                context.rels,
            ))
            .overlay(parse_run_style(
                layout_properties.and_then(|properties| properties.child("defRPr")),
                context.colors,
                context.rels,
            ))
            .overlay(parse_run_style(
                own_properties.and_then(|properties| properties.child("defRPr")),
                context.colors,
                context.rels,
            ));
            let mut runs = Vec::new();
            for child in &paragraph.children {
                match child.name.as_str() {
                    "r" | "fld" => {
                        let style = default_style.clone().overlay(parse_run_style(
                            child.child("rPr"),
                            context.colors,
                            context.rels,
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
                            context.colors,
                            context.rels,
                        ));
                        runs.push(text_run("\n".into(), style));
                    }
                    "tab" => {
                        let style = default_style.clone().overlay(parse_run_style(
                            child.child("rPr"),
                            context.colors,
                            context.rels,
                        ));
                        runs.push(text_run("\t".into(), style));
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
                margin_left_emu: paragraph_style.margin_left_emu,
                indent_emu: paragraph_style.indent_emu,
                default_tab_size_emu: paragraph_style.default_tab_size_emu,
                tab_stops: paragraph_style.tab_stops,
                east_asian_line_break: paragraph_style.east_asian_line_break,
                latin_line_break: paragraph_style.latin_line_break,
                hanging_punctuation: paragraph_style.hanging_punctuation,
            }
        })
        .collect()
}

fn inherited_body_attribute<'a>(bodies: &[Option<&'a XmlNode>], name: &str) -> Option<&'a str> {
    bodies.iter().find_map(|body| {
        body.and_then(|body| body.child("bodyPr"))
            .and_then(|properties| properties.attr(name))
    })
}

fn vertical_alignment_from_value(value: Option<&str>) -> Option<VerticalAlignment> {
    match value {
        Some("ctr") | Some("just") | Some("dist") => Some(VerticalAlignment::Middle),
        Some("b") => Some(VerticalAlignment::Bottom),
        Some("t") => Some(VerticalAlignment::Top),
        _ => None,
    }
}

fn inherited_vertical_alignment(bodies: &[Option<&XmlNode>]) -> Option<VerticalAlignment> {
    vertical_alignment_from_value(inherited_body_attribute(bodies, "anchor"))
}

fn inherited_text_insets(bodies: &[Option<&XmlNode>]) -> Option<crate::TextInsets> {
    if bodies.iter().all(Option::is_none) {
        return None;
    }
    let edge = |name: &str, default: i64| {
        inherited_body_attribute(bodies, name)
            .and_then(|value| value.parse::<i64>().ok())
            .unwrap_or(default)
    };
    Some(crate::TextInsets {
        top: edge("tIns", 45_720),
        right: edge("rIns", 91_440),
        bottom: edge("bIns", 45_720),
        left: edge("lIns", 91_440),
    })
}

fn inherited_autofit(bodies: &[Option<&XmlNode>]) -> Option<crate::TextAutofit> {
    let autofit = bodies.iter().find_map(|body| {
        let properties = body.and_then(|body| body.child("bodyPr"))?;
        if properties.child("noAutofit").is_some() {
            Some(crate::TextAutofit {
                mode: crate::TextAutoFitMode::None,
                font_scale: None,
                line_spacing_reduction: None,
            })
        } else if let Some(normal) = properties.child("normAutofit") {
            Some(crate::TextAutofit {
                mode: crate::TextAutoFitMode::Normal,
                font_scale: normal.attr("fontScale").and_then(percentage),
                line_spacing_reduction: normal.attr("lnSpcReduction").and_then(percentage),
            })
        } else if properties.child("spAutoFit").is_some() {
            Some(crate::TextAutofit {
                mode: crate::TextAutoFitMode::Shape,
                font_scale: None,
                line_spacing_reduction: None,
            })
        } else {
            None
        }
    });
    autofit.or_else(|| {
        bodies
            .iter()
            .any(Option::is_some)
            .then_some(crate::TextAutofit {
                mode: crate::TextAutoFitMode::None,
                font_scale: None,
                line_spacing_reduction: None,
            })
    })
}

fn shape_font_style(node: Option<&XmlNode>, colors: Option<&ColorContext<'_>>) -> RunStyle {
    let Some(reference) = node
        .and_then(|node| node.child("style"))
        .and_then(|style| style.child("fontRef"))
    else {
        return RunStyle::default();
    };
    let fonts = match reference.attr("idx") {
        Some("major") => colors
            .and_then(|context| context.theme)
            .map(|theme| &theme.major_fonts),
        Some("minor") => colors
            .and_then(|context| context.theme)
            .map(|theme| &theme.minor_fonts),
        _ => None,
    };
    RunStyle {
        font_family: fonts.and_then(|fonts| fonts.get("latin")).cloned(),
        east_asian_font_family: fonts.and_then(|fonts| fonts.get("eastAsia")).cloned(),
        complex_script_font_family: fonts.and_then(|fonts| fonts.get("complexScript")).cloned(),
        color: parse_color(reference, colors),
        ..Default::default()
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

fn compact_path_number(value: f64) -> String {
    let mut rendered = format!("{value:.8}");
    while rendered.contains('.') && rendered.ends_with('0') {
        rendered.pop();
    }
    if rendered.ends_with('.') {
        rendered.pop();
    }
    if rendered == "-0" {
        "0".into()
    } else {
        rendered
    }
}

fn custom_path_point(node: &XmlNode) -> Option<(f64, f64)> {
    Some((
        node.attr("x")?.parse::<f64>().ok()?,
        node.attr("y")?.parse::<f64>().ok()?,
    ))
}

/// Converts DrawingML custom paths into SVG path data normalized to a 0..1
/// view box. Most authored freeforms use literal path coordinates; unsupported
/// guide formulas are skipped instead of degrading the shape to a rectangle.
fn custom_geometry_path(node: &XmlNode) -> Option<String> {
    let paths = node.child("spPr")?.child("custGeom")?.child("pathLst")?;
    let mut output = Vec::new();
    let mut has_drawn_segment = false;
    for path in paths.children_named("path") {
        let width = path
            .attr("w")
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|value| *value > 0.0)?;
        let height = path
            .attr("h")
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|value| *value > 0.0)?;
        let normalized = |(x, y): (f64, f64)| {
            format!(
                "{} {}",
                compact_path_number(x / width),
                compact_path_number(y / height)
            )
        };
        let mut current = None;
        for command in &path.children {
            match command.name.as_str() {
                "moveTo" => {
                    if let Some(point) = command.child("pt").and_then(custom_path_point) {
                        output.push(format!("M {}", normalized(point)));
                        current = Some(point);
                    }
                }
                "lnTo" => {
                    if let Some(point) = command.child("pt").and_then(custom_path_point) {
                        output.push(format!("L {}", normalized(point)));
                        current = Some(point);
                        has_drawn_segment = true;
                    }
                }
                "cubicBezTo" => {
                    let points = command
                        .children_named("pt")
                        .filter_map(custom_path_point)
                        .collect::<Vec<_>>();
                    if let [first, second, third] = points.as_slice() {
                        output.push(format!(
                            "C {} {} {}",
                            normalized(*first),
                            normalized(*second),
                            normalized(*third)
                        ));
                        current = Some(*third);
                        has_drawn_segment = true;
                    }
                }
                "quadBezTo" => {
                    let points = command
                        .children_named("pt")
                        .filter_map(custom_path_point)
                        .collect::<Vec<_>>();
                    if let [first, second] = points.as_slice() {
                        output.push(format!("Q {} {}", normalized(*first), normalized(*second)));
                        current = Some(*second);
                        has_drawn_segment = true;
                    }
                }
                "arcTo" => {
                    let Some((start_x, start_y)) = current else {
                        continue;
                    };
                    let Some(radius_x) = command
                        .attr("wR")
                        .and_then(|value| value.parse::<f64>().ok())
                    else {
                        continue;
                    };
                    let Some(radius_y) = command
                        .attr("hR")
                        .and_then(|value| value.parse::<f64>().ok())
                    else {
                        continue;
                    };
                    let Some(start_angle) = command.attr("stAng").and_then(angle_degrees) else {
                        continue;
                    };
                    let Some(sweep_angle) = command.attr("swAng").and_then(angle_degrees) else {
                        continue;
                    };
                    let start_radians = start_angle.to_radians();
                    let end_radians = (start_angle + sweep_angle).to_radians();
                    let center_x = start_x - radius_x * start_radians.cos();
                    let center_y = start_y - radius_y * start_radians.sin();
                    let endpoint = (
                        center_x + radius_x * end_radians.cos(),
                        center_y + radius_y * end_radians.sin(),
                    );
                    output.push(format!(
                        "A {} {} 0 {} {} {}",
                        compact_path_number(radius_x / width),
                        compact_path_number(radius_y / height),
                        usize::from(sweep_angle.abs() > 180.0),
                        usize::from(sweep_angle >= 0.0),
                        normalized(endpoint)
                    ));
                    current = Some(endpoint);
                    has_drawn_segment = true;
                }
                "close" => output.push("Z".into()),
                _ => {}
            }
        }
    }
    (has_drawn_segment && !output.is_empty()).then(|| output.join(" "))
}

fn shape_fill(
    node: &XmlNode,
    colors: Option<&ColorContext<'_>>,
    rels: &HashMap<String, String>,
) -> Option<Option<FillStyle>> {
    // The slide already paints its background, so a background-fill shape is
    // transparent in the normalized composition. Falling through to fillRef
    // here incorrectly paints it with the theme accent color.
    if node.attr("useBgFill").and_then(bool_value).unwrap_or(false) {
        return Some(Some(FillStyle::None));
    }
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
            return Some(parse_fill(properties, colors, rels));
        }
    }
    node.child("style")
        .and_then(|style| style.child("fillRef"))
        .map(|fill| style_matrix_fill(fill, colors))
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
        .map(|line| style_matrix_line(line, colors))
}

struct NodeParseContext<'a, 'color> {
    slide_index: usize,
    rels: &'a HashMap<String, String>,
    related_parts: &'a HashMap<String, Vec<u8>>,
    chart_companions: &'a HashMap<String, ChartCompanionParts>,
    limits: &'a ParseLimits,
    colors: Option<&'a ColorContext<'color>>,
    diagnostics: &'a ParseDiagnostics,
    layout_root: Option<&'a XmlNode>,
    master_root: Option<&'a XmlNode>,
    layout_rels: Option<&'a HashMap<String, String>>,
    master_rels: Option<&'a HashMap<String, String>>,
    presentation_text_style: Option<&'a XmlNode>,
    table_styles: Option<&'a XmlNode>,
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
    let fill = [
        (Some(node), Some(context.rels)),
        (layout, context.layout_rels),
        (master, context.master_rels),
    ]
    .into_iter()
    .find_map(|(shape, rels)| shape_fill(shape?, context.colors, rels?))
    .flatten();
    let line = [Some(node), layout, master]
        .into_iter()
        .flatten()
        .find_map(|shape| shape_line(shape, context.colors))
        .flatten();
    let text_bodies = [text_body, layout_text_body, master_text_body];
    let shape_default_style = shape_font_style(master, context.colors)
        .overlay(shape_font_style(layout, context.colors))
        .overlay(shape_font_style(Some(node), context.colors));
    let text_context = TextParseContext {
        presentation_text_style: context.presentation_text_style,
        shape_default_style,
        colors: context.colors,
        rels: context.rels,
    };
    SlideNode::Shape {
        id,
        name,
        transform: inherited_shape_transform(node, layout, master),
        geometry: ShapeGeometry {
            preset: geometry,
            path: custom_geometry_path(node),
        },
        fill,
        line,
        paragraphs: parse_text_paragraphs(
            text_body,
            layout_text_body,
            master_text_body,
            master_text_style(context.master_root, master_source),
            &text_context,
        ),
        vertical_alignment: inherited_vertical_alignment(&text_bodies),
        text_insets: inherited_text_insets(&text_bodies),
        autofit: inherited_autofit(&text_bodies),
        text_rotation: inherited_body_attribute(&text_bodies, "rot").and_then(angle_degrees),
        vertical_text: inherited_body_attribute(&text_bodies, "vert").map(str::to_owned),
        horizontal_overflow: inherited_body_attribute(&text_bodies, "horzOverflow")
            .map(str::to_owned),
        vertical_overflow: inherited_body_attribute(&text_bodies, "vertOverflow")
            .map(str::to_owned),
        text_wrap: text_bodies.iter().any(Option::is_some).then(|| {
            inherited_body_attribute(&text_bodies, "wrap")
                .unwrap_or("square")
                .to_owned()
        }),
        column_count: text_bodies.iter().any(Option::is_some).then(|| {
            inherited_body_attribute(&text_bodies, "numCol")
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(1)
        }),
        column_spacing: inherited_body_attribute(&text_bodies, "spcCol")
            .and_then(|value| value.parse::<i64>().ok()),
        right_to_left_columns: text_bodies.iter().any(Option::is_some).then(|| {
            inherited_body_attribute(&text_bodies, "rtlCol")
                .and_then(bool_value)
                .unwrap_or(false)
        }),
        space_first_last_paragraph: text_bodies.iter().any(Option::is_some).then(|| {
            inherited_body_attribute(&text_bodies, "spcFirstLastPara")
                .and_then(bool_value)
                .unwrap_or(false)
        }),
    }
}

fn parse_image_node(
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
    let resolved_fill = [
        (Some(node), Some(context.rels)),
        (layout, context.layout_rels),
        (master, context.master_rels),
    ]
    .into_iter()
    .find_map(|(source, rels)| {
        let fill = source?.child("blipFill")?;
        let relationship = fill
            .child("blip")?
            .attr("embed")
            .or_else(|| fill.child("blip")?.attr("link"))?;
        let asset_id = rels?.get(relationship)?.clone();
        Some((fill, asset_id))
    });
    let asset_id = resolved_fill.as_ref().map(|(_, asset_id)| asset_id.clone());
    let crop = resolved_fill
        .as_ref()
        .and_then(|(fill, _)| image_crop(fill));
    let opacity = resolved_fill.as_ref().and_then(|(fill, _)| {
        fill.child("blip")
            .and_then(|blip| blip.child("alphaModFix"))
            .and_then(|effect| effect.attr("amt"))
            .and_then(percentage)
            .map(|value| value.clamp(0.0, 1.0))
    });
    let preserve_aspect_ratio = resolved_fill
        .as_ref()
        .map(|(fill, _)| fill.child("stretch").is_none())
        .unwrap_or(true);
    let effects = resolved_fill.as_ref().and_then(|(fill, _)| {
        fill.child("blip")
            .and_then(|blip| parse_blip_effects(blip, context.colors))
    });
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
        transform: inherited_shape_transform(node, layout, master),
        asset_id: asset_id.unwrap_or_else(|| "missing-asset".into()),
        crop,
        opacity,
        preserve_aspect_ratio,
        effects,
    }
}

fn selected_table_style<'a>(table: &XmlNode, styles: Option<&'a XmlNode>) -> Option<&'a XmlNode> {
    let styles = styles?;
    let requested = table
        .child("tblPr")
        .and_then(|properties| properties.child("tableStyleId"))
        .map(XmlNode::text_content)
        .filter(|value| !value.trim().is_empty())
        .or_else(|| styles.attr("def").map(str::to_owned))?;
    styles
        .children_named("tblStyle")
        .find(|style| style.attr("styleId") == Some(requested.trim()))
}

fn table_flag(table: &XmlNode, name: &str) -> bool {
    table
        .child("tblPr")
        .and_then(|properties| properties.attr(name))
        .and_then(bool_value)
        .unwrap_or(false)
}

fn table_style_regions<'a>(
    table: &XmlNode,
    style: Option<&'a XmlNode>,
    row_index: usize,
    column_index: usize,
    row_count: usize,
    column_count: usize,
) -> Vec<&'a XmlNode> {
    let Some(style) = style else {
        return Vec::new();
    };
    let mut regions = Vec::new();
    if let Some(region) = style.child("wholeTbl") {
        regions.push(region);
    }
    let first_row = table_flag(table, "firstRow");
    let last_row = table_flag(table, "lastRow");
    let first_column = table_flag(table, "firstCol");
    let last_column = table_flag(table, "lastCol");
    if table_flag(table, "bandRow") {
        let band_index = row_index.saturating_sub(usize::from(first_row));
        let name = if band_index.is_multiple_of(2) {
            "band1H"
        } else {
            "band2H"
        };
        if let Some(region) = style.child(name) {
            regions.push(region);
        }
    }
    if table_flag(table, "bandCol") {
        let band_index = column_index.saturating_sub(usize::from(first_column));
        let name = if band_index.is_multiple_of(2) {
            "band1V"
        } else {
            "band2V"
        };
        if let Some(region) = style.child(name) {
            regions.push(region);
        }
    }
    if first_column && column_index == 0 {
        if let Some(region) = style.child("firstCol") {
            regions.push(region);
        }
    }
    if last_column && column_index.saturating_add(1) == column_count {
        if let Some(region) = style.child("lastCol") {
            regions.push(region);
        }
    }
    if first_row && row_index == 0 {
        if let Some(region) = style.child("firstRow") {
            regions.push(region);
        }
    }
    if last_row && row_index.saturating_add(1) == row_count {
        if let Some(region) = style.child("lastRow") {
            regions.push(region);
        }
    }
    let corner = match (
        row_index == 0,
        row_index.saturating_add(1) == row_count,
        column_index == 0,
        column_index.saturating_add(1) == column_count,
    ) {
        (true, _, true, _) => Some("nwCell"),
        (true, _, _, true) => Some("neCell"),
        (_, true, true, _) => Some("swCell"),
        (_, true, _, true) => Some("seCell"),
        _ => None,
    };
    if let Some(region) = corner.and_then(|name| style.child(name)) {
        regions.push(region);
    }
    regions
}

fn table_region_fill(region: &XmlNode, context: &NodeParseContext<'_, '_>) -> Option<FillStyle> {
    let style = region.child("tcStyle")?;
    style
        .child("fill")
        .and_then(|fill| parse_fill(fill, context.colors, context.rels))
        .or_else(|| {
            style
                .child("fillRef")
                .and_then(|reference| style_matrix_fill(reference, context.colors))
        })
}

fn table_region_text_style(region: &XmlNode, context: &NodeParseContext<'_, '_>) -> RunStyle {
    let Some(style) = region.child("tcTxStyle") else {
        return RunStyle::default();
    };
    let reference = style.child("fontRef");
    let fonts = match reference.and_then(|reference| reference.attr("idx")) {
        Some("major") => context
            .colors
            .and_then(|colors| colors.theme)
            .map(|theme| &theme.major_fonts),
        Some("minor") => context
            .colors
            .and_then(|colors| colors.theme)
            .map(|theme| &theme.minor_fonts),
        _ => None,
    };
    let color = style
        .children
        .iter()
        .find(|child| {
            matches!(
                child.name.as_str(),
                "srgbClr" | "scrgbClr" | "hslClr" | "schemeClr" | "sysClr" | "prstClr"
            )
        })
        .and_then(|color| parse_color(color, context.colors))
        .or_else(|| reference.and_then(|reference| parse_color(reference, context.colors)));
    RunStyle {
        font_family: fonts.and_then(|fonts| fonts.get("latin")).cloned(),
        east_asian_font_family: fonts.and_then(|fonts| fonts.get("eastAsia")).cloned(),
        complex_script_font_family: fonts.and_then(|fonts| fonts.get("complexScript")).cloned(),
        bold: style.attr("b").and_then(bool_value),
        italic: style.attr("i").and_then(bool_value),
        color,
        ..Default::default()
    }
}

fn table_region_borders(
    region: &XmlNode,
    context: &NodeParseContext<'_, '_>,
    row_index: usize,
    column_index: usize,
    row_count: usize,
    column_count: usize,
) -> BTreeMap<String, LineStyle> {
    let Some(borders) = region
        .child("tcStyle")
        .and_then(|style| style.child("tcBdr"))
    else {
        return BTreeMap::new();
    };
    let mut result = BTreeMap::new();
    for (side, interior) in [
        ("top", (row_index > 0).then_some("insideH")),
        (
            "right",
            (column_index.saturating_add(1) < column_count).then_some("insideV"),
        ),
        (
            "bottom",
            (row_index.saturating_add(1) < row_count).then_some("insideH"),
        ),
        ("left", (column_index > 0).then_some("insideV")),
    ] {
        let line = borders
            .child(side)
            .or_else(|| interior.and_then(|name| borders.child(name)))
            .and_then(|container| container.child("ln"))
            .and_then(|line| parse_line_node(line, context.colors));
        if let Some(line) = line {
            result.insert(side.to_owned(), line);
        }
    }
    result
}

fn parse_table_cell(
    cell: &XmlNode,
    context: &NodeParseContext<'_, '_>,
    regions: &[&XmlNode],
    row_index: usize,
    column_index: usize,
    row_count: usize,
    column_count: usize,
) -> TableCell {
    let properties = cell.child("tcPr");
    let mut borders = BTreeMap::new();
    let mut style_fill = None;
    let mut text_style = RunStyle::default();
    for region in regions {
        if let Some(fill) = table_region_fill(region, context) {
            style_fill = Some(fill);
        }
        borders.extend(table_region_borders(
            region,
            context,
            row_index,
            column_index,
            row_count,
            column_count,
        ));
        text_style = text_style.overlay(table_region_text_style(region, context));
    }
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
    let text_context = TextParseContext {
        presentation_text_style: context.presentation_text_style,
        shape_default_style: text_style,
        colors: context.colors,
        rels: context.rels,
    };
    TableCell {
        row_span: cell
            .attr("rowSpan")
            .and_then(|value| value.parse().ok())
            .filter(|value| *value > 1),
        col_span: cell
            .attr("gridSpan")
            .and_then(|value| value.parse().ok())
            .filter(|value| *value > 1),
        fill: properties
            .and_then(|properties| parse_fill(properties, context.colors, context.rels))
            .or(style_fill),
        borders,
        paragraphs: parse_text_paragraphs(cell.child("txBody"), None, None, None, &text_context),
        text_insets: cell.child("txBody").map(|_| {
            let value = |name: &str, default: i64| {
                properties
                    .and_then(|properties| properties.attr(name))
                    .and_then(|value| value.parse::<i64>().ok())
                    .unwrap_or(default)
            };
            crate::TextInsets {
                top: value("marT", 45_720),
                right: value("marR", 91_440),
                bottom: value("marB", 45_720),
                left: value("marL", 91_440),
            }
        }),
        vertical_alignment: vertical_alignment_from_value(
            properties.and_then(|properties| properties.attr("anchor")),
        ),
        text_rotation: properties
            .and_then(|properties| properties.attr("vert"))
            .and_then(|value| match value {
                "vert270" => Some(270.0),
                "vert" | "eaVert" | "wordArtVert" | "wordArtVertRtl" => Some(90.0),
                _ => None,
            }),
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
    let column_widths = table
        .child("tblGrid")
        .into_iter()
        .flat_map(|grid| grid.children_named("gridCol"))
        .filter_map(|column| column.attr("w").and_then(|value| value.parse().ok()))
        .collect();
    let row_nodes = table.children_named("tr").collect::<Vec<_>>();
    let row_count = row_nodes.len();
    let column_count = table
        .child("tblGrid")
        .map(|grid| grid.children_named("gridCol").count())
        .filter(|count| *count > 0)
        .unwrap_or_else(|| {
            row_nodes
                .iter()
                .map(|row| row.children_named("tc").count())
                .max()
                .unwrap_or(0)
        });
    let style = selected_table_style(table, context.table_styles);
    let rows = row_nodes
        .iter()
        .enumerate()
        .map(|(row_index, row)| {
            let mut logical_column = 0usize;
            row.children_named("tc")
                .filter_map(|cell| {
                    let column_index = logical_column;
                    logical_column = logical_column.saturating_add(
                        cell.attr("gridSpan")
                            .and_then(|value| value.parse::<usize>().ok())
                            .unwrap_or(1),
                    );
                    if is_merge_continuation(cell) {
                        return None;
                    }
                    let regions = table_style_regions(
                        table,
                        style,
                        row_index,
                        column_index,
                        row_count,
                        column_count,
                    );
                    Some(parse_table_cell(
                        cell,
                        context,
                        &regions,
                        row_index,
                        column_index,
                        row_count,
                        column_count,
                    ))
                })
                .collect()
        })
        .collect::<Vec<Vec<TableCell>>>();
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
        .unwrap_or_else(|| {
            // Modern chartEx parts (`cx:chartSpace`) carry their cached data in
            // `chartData`; classic parts always declare one of the types above.
            if root.descendant("chartData").is_some() {
                (&root, "chartEx")
            } else {
                (&root, "ooxml")
            }
        });
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

#[derive(Default)]
struct ParsedSmartArtData {
    semantic_text: HashMap<String, Vec<TextParagraph>>,
    semantic_labels: Vec<(String, Vec<TextParagraph>)>,
    presentation_to_semantic: HashMap<String, String>,
    drawing_relationship_id: Option<String>,
    background: Option<FillStyle>,
}

fn paragraphs_have_text(paragraphs: &[TextParagraph]) -> bool {
    paragraphs
        .iter()
        .any(|paragraph| paragraph.runs.iter().any(|run| !run.text.trim().is_empty()))
}

fn parse_smartart_data(root: &XmlNode, context: &NodeParseContext<'_, '_>) -> ParsedSmartArtData {
    let mut parsed = ParsedSmartArtData {
        drawing_relationship_id: root
            .descendant("dataModelExt")
            .and_then(|extension| extension.attr("relId"))
            .map(str::to_owned),
        background: root
            .child("bg")
            .and_then(|background| parse_fill(background, context.colors, context.rels)),
        ..Default::default()
    };
    let Some(points) = root.child("ptLst") else {
        return parsed;
    };
    for point in points.children_named("pt") {
        let Some(model_id) = point.attr("modelId") else {
            continue;
        };
        let point_type = point.attr("type");
        if point_type == Some("pres") {
            if let Some(semantic_id) = point
                .child("prSet")
                .and_then(|properties| properties.attr("presAssocID"))
            {
                parsed
                    .presentation_to_semantic
                    .insert(model_id.to_owned(), semantic_id.to_owned());
            }
            continue;
        }
        if matches!(point_type, Some("doc" | "parTrans" | "sibTrans")) {
            continue;
        }
        let text_context = TextParseContext {
            presentation_text_style: context.presentation_text_style,
            shape_default_style: RunStyle::default(),
            colors: context.colors,
            rels: context.rels,
        };
        let paragraphs = parse_text_paragraphs(point.child("t"), None, None, None, &text_context);
        if paragraphs_have_text(&paragraphs) {
            parsed
                .semantic_text
                .insert(model_id.to_owned(), paragraphs.clone());
            parsed
                .semantic_labels
                .push((model_id.to_owned(), paragraphs));
        }
    }
    parsed
}

fn smartart_child_transform(frame: &Transform) -> Transform {
    Transform {
        x: 0,
        y: 0,
        width: frame.width.max(1),
        height: frame.height.max(1),
        ..Default::default()
    }
}

fn mapped_smartart_text<'a>(
    model_id: &str,
    data: &'a ParsedSmartArtData,
) -> Option<&'a Vec<TextParagraph>> {
    data.presentation_to_semantic
        .get(model_id)
        .and_then(|semantic_id| data.semantic_text.get(semantic_id))
        .or_else(|| data.semantic_text.get(model_id))
}

fn parse_materialized_smartart(
    frame_node: &XmlNode,
    node_index: usize,
    drawing_root: &XmlNode,
    data: &ParsedSmartArtData,
    context: &NodeParseContext<'_, '_>,
) -> Option<SlideNode> {
    let shape_tree = drawing_root.descendant("spTree")?;
    let (frame_id, frame_name) = node_identity(frame_node, context.slide_index, node_index);
    let mut children = Vec::new();
    for (shape_index, shape) in shape_tree.children.iter().enumerate() {
        if !matches!(shape.name.as_str(), "sp" | "cxnSp") {
            continue;
        }
        let model_id = shape
            .attr("modelId")
            .map(str::to_owned)
            .unwrap_or_else(|| format!("shape-{shape_index}"));
        let mut parsed = parse_shape_node(shape, shape_index, context);
        if let SlideNode::Shape {
            id,
            name,
            paragraphs,
            ..
        } = &mut parsed
        {
            *id = format!("{frame_id}:smartart:{model_id}");
            if name.is_empty() {
                *name = format!("SmartArt shape {}", shape_index + 1);
            }
            if !paragraphs_have_text(paragraphs) {
                if let Some(mapped) = mapped_smartart_text(&model_id, data) {
                    *paragraphs = mapped.clone();
                }
            }
        }
        children.push(parsed);
    }
    if children.is_empty() {
        return None;
    }
    let frame = parse_transform(frame_node.child("xfrm"));
    let drawing_xfrm = shape_tree
        .child("grpSpPr")
        .and_then(|properties| properties.child("xfrm"));
    Some(SlideNode::Group {
        id: frame_id,
        name: frame_name,
        transform: frame.clone(),
        children,
        child_transform: parse_child_transform(drawing_xfrm)
            .or_else(|| Some(smartart_child_transform(&frame))),
    })
}

fn normalize_smartart_fallback_text(paragraphs: &mut [TextParagraph]) {
    for paragraph in paragraphs {
        paragraph.alignment = Some("center".into());
        paragraph.bullet = None;
        paragraph.level = None;
        paragraph.margin_left_emu = None;
        paragraph.indent_emu = None;
        paragraph.line_spacing = None;
        paragraph.space_before = None;
        paragraph.space_after = None;
        for run in &mut paragraph.runs {
            run.font_size_pt = Some(run.font_size_pt.unwrap_or(12.0).clamp(8.0, 18.0));
            run.color = Some(ColorValue {
                value: "#172033".into(),
                alpha: None,
            });
        }
    }
}

fn smartart_fallback_group(
    frame_node: &XmlNode,
    node_index: usize,
    data: &ParsedSmartArtData,
    context: &NodeParseContext<'_, '_>,
) -> Option<SlideNode> {
    if data.semantic_labels.is_empty() {
        return None;
    }
    let (frame_id, frame_name) = node_identity(frame_node, context.slide_index, node_index);
    let frame = parse_transform(frame_node.child("xfrm"));
    let width = frame.width.max(1);
    let height = frame.height.max(1);
    let count = data.semantic_labels.len();
    let aspect = width as f64 / height as f64;
    let columns = ((count as f64 * aspect).sqrt().ceil() as usize).clamp(1, count);
    let rows = count.div_ceil(columns);
    let gap = (width.min(height) / 40).max(1);
    let padding = gap.saturating_mul(2);
    let horizontal_gaps = gap.saturating_mul(columns.saturating_sub(1) as i64);
    let vertical_gaps = gap.saturating_mul(rows.saturating_sub(1) as i64);
    let cell_width = width
        .saturating_sub(padding.saturating_mul(2))
        .saturating_sub(horizontal_gaps)
        .max(columns as i64)
        / columns as i64;
    let cell_height = height
        .saturating_sub(padding.saturating_mul(2))
        .saturating_sub(vertical_gaps)
        .max(rows as i64)
        / rows as i64;
    let mut children =
        Vec::with_capacity(count.saturating_add(usize::from(data.background.is_some())));
    if let Some(fill) = data.background.clone() {
        children.push(SlideNode::Shape {
            id: format!("{frame_id}:smartart:background"),
            name: "SmartArt background".into(),
            transform: smartart_child_transform(&frame),
            geometry: ShapeGeometry {
                preset: Some("rect".into()),
                path: None,
            },
            fill: Some(fill),
            line: None,
            paragraphs: Vec::new(),
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
        });
    }
    for (label_index, (model_id, source_paragraphs)) in data.semantic_labels.iter().enumerate() {
        let column = label_index % columns;
        let row = label_index / columns;
        let mut paragraphs = source_paragraphs.clone();
        normalize_smartart_fallback_text(&mut paragraphs);
        let horizontal_inset = (cell_width / 20).max(0);
        let vertical_inset = (cell_height / 12).max(0);
        children.push(SlideNode::Shape {
            id: format!("{frame_id}:smartart:fallback:{model_id}"),
            name: format!("SmartArt item {}", label_index + 1),
            transform: Transform {
                x: padding + column as i64 * (cell_width + gap),
                y: padding + row as i64 * (cell_height + gap),
                width: cell_width,
                height: cell_height,
                ..Default::default()
            },
            geometry: ShapeGeometry {
                preset: Some("roundRect".into()),
                path: None,
            },
            fill: Some(FillStyle::Solid {
                color: ColorValue {
                    value: "#E8EEF8".into(),
                    alpha: None,
                },
            }),
            line: Some(LineStyle {
                color: Some(ColorValue {
                    value: "#5273A8".into(),
                    alpha: None,
                }),
                width: Some(1.0),
                ..Default::default()
            }),
            paragraphs,
            vertical_alignment: Some(VerticalAlignment::Middle),
            text_insets: Some(crate::TextInsets {
                top: vertical_inset,
                right: horizontal_inset,
                bottom: vertical_inset,
                left: horizontal_inset,
            }),
            autofit: Some(crate::TextAutofit {
                mode: crate::TextAutoFitMode::Normal,
                font_scale: Some(0.85),
                line_spacing_reduction: Some(0.1),
            }),
            text_rotation: None,
            vertical_text: None,
            horizontal_overflow: Some("clip".into()),
            vertical_overflow: Some("clip".into()),
            text_wrap: Some("square".into()),
            column_count: Some(1),
            column_spacing: None,
            right_to_left_columns: Some(false),
            space_first_last_paragraph: Some(false),
        });
    }
    context.diagnostics.warn(
        "degraded-rendering",
        format!(
            "SmartArt frame {frame_id} has no materialized drawing; rendered a semantic fallback."
        ),
        Some("smartart"),
    );
    Some(SlideNode::Group {
        id: frame_id,
        name: frame_name,
        transform: frame.clone(),
        children,
        child_transform: Some(smartart_child_transform(&frame)),
    })
}

fn parse_smartart_frame(
    frame_node: &XmlNode,
    node_index: usize,
    graphic_data: &XmlNode,
    context: &NodeParseContext<'_, '_>,
) -> Result<Option<SlideNode>, ParseError> {
    let Some(relationships) = graphic_data.descendant("relIds") else {
        return Ok(None);
    };
    let Some(data_relationship_id) = relationships.attr("dm") else {
        return Ok(None);
    };
    let Some(data_path) = context.rels.get(data_relationship_id) else {
        context.diagnostics.warn(
            "missing-part",
            "SmartArt data relationship could not be resolved.",
            Some("smartart"),
        );
        return Ok(None);
    };
    let Some(data_xml) = context.related_parts.get(data_path) else {
        context.diagnostics.warn(
            "missing-part",
            format!("SmartArt data part {data_path} is missing."),
            Some("smartart"),
        );
        return Ok(None);
    };
    let data_root = parse_xml_tree(data_xml, context.limits, data_path)?;
    let data = parse_smartart_data(&data_root, context);
    if let Some(drawing_relationship_id) = data.drawing_relationship_id.as_deref() {
        if let Some(drawing_path) = context.rels.get(drawing_relationship_id) {
            if let Some(drawing_xml) = context.related_parts.get(drawing_path) {
                let drawing_root = parse_xml_tree(drawing_xml, context.limits, drawing_path)?;
                if let Some(group) = parse_materialized_smartart(
                    frame_node,
                    node_index,
                    &drawing_root,
                    &data,
                    context,
                ) {
                    return Ok(Some(group));
                }
            } else {
                context.diagnostics.warn(
                    "missing-part",
                    format!("SmartArt drawing part {drawing_path} is missing."),
                    Some("smartart"),
                );
            }
        } else {
            context.diagnostics.warn(
                "missing-part",
                "SmartArt drawing relationship could not be resolved.",
                Some("smartart"),
            );
        }
    }
    Ok(smartart_fallback_group(
        frame_node, node_index, &data, context,
    ))
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
        let resolved = path.and_then(|path| context.related_parts.get(path).map(|xml| (path, xml)));
        let (chart, chart_xml, companions) = if let Some((path, xml)) = resolved {
            (
                parse_chart_part(xml, context.limits, path, context.colors)?,
                Some(String::from_utf8_lossy(xml).into_owned()),
                context.chart_companions.get(path),
            )
        } else {
            context.diagnostics.warn(
                "missing-part",
                "Chart relationship or chart part could not be resolved.",
                Some("chart"),
            );
            (
                ParsedChart {
                    chart_type: "ooxml".into(),
                    title: None,
                    series: Vec::new(),
                    has_legend: None,
                },
                None,
                None,
            )
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
            chart_xml,
            chart_style_xml: companions.and_then(|parts| parts.style_xml.clone()),
            chart_colors_xml: companions.and_then(|parts| parts.colors_xml.clone()),
        });
    }
    if let Some(smartart) = graphic_data.filter(|value| value.descendant("relIds").is_some()) {
        if let Some(group) = parse_smartart_frame(node, node_index, smartart, context)? {
            return Ok(group);
        }
    }
    if let Some(picture) = graphic_data.and_then(|value| value.descendant("pic")) {
        let mut fallback = parse_image_node(picture, node_index, context);
        let (frame_id, frame_name) = node_identity(node, context.slide_index, node_index);
        if let SlideNode::Image { id, name, .. } = &mut fallback {
            *id = frame_id;
            *name = frame_name;
        }
        return Ok(fallback);
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
        // Markup-compatibility wrapper used by PowerPoint for modern content
        // (chartEx funnel/treemap/waterfall charts, newer media). Prefer the
        // richest branch that parses to a supported node, otherwise keep any
        // fallback rendering.
        "AlternateContent" => {
            let mut fallback: Option<SlideNode> = None;
            for branch_name in ["Choice", "Fallback"] {
                for branch in node.children_named(branch_name) {
                    for child in &branch.children {
                        let Some(parsed) = parse_slide_node(child, node_index, context)? else {
                            continue;
                        };
                        if !matches!(parsed, SlideNode::Unknown { .. }) {
                            return Ok(Some(parsed));
                        }
                        if fallback.is_none() {
                            fallback = Some(parsed);
                        }
                    }
                }
            }
            Ok(fallback)
        }
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

fn part_background(
    root: &XmlNode,
    colors: Option<&ColorContext<'_>>,
    rels: &HashMap<String, String>,
) -> Option<FillStyle> {
    let background = root.child("cSld")?.child("bg")?;
    background
        .child("bgPr")
        .and_then(|properties| parse_fill(properties, colors, rels))
        .or_else(|| {
            background
                .child("bgRef")
                .and_then(|reference| style_matrix_fill(reference, colors))
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

fn parse_decorative_part_nodes(
    root: &XmlNode,
    context: &NodeParseContext<'_, '_>,
) -> Result<Vec<SlideNode>, ParseError> {
    let mut nodes = Vec::new();
    let Some(shape_tree) = shape_tree(root) else {
        return Ok(nodes);
    };
    for (node_index, node) in shape_tree.children.iter().enumerate() {
        if placeholder_key(node).is_some() {
            continue;
        }
        if let Some(node) = parse_slide_node(node, node_index, context)? {
            nodes.push(node);
        }
    }
    Ok(nodes)
}

fn prefix_node_ids(node: &mut SlideNode, prefix: &str) {
    match node {
        SlideNode::Shape { id, .. }
        | SlideNode::Image { id, .. }
        | SlideNode::Table { id, .. }
        | SlideNode::Chart { id, .. }
        | SlideNode::Unknown { id, .. } => *id = format!("{prefix}:{id}"),
        SlideNode::Group { id, children, .. } => {
            *id = format!("{prefix}:{id}");
            for child in children {
                prefix_node_ids(child, prefix);
            }
        }
    }
}

fn master_shapes_are_visible(root: &XmlNode) -> bool {
    root.attr("showMasterSp")
        .and_then(bool_value)
        .unwrap_or(true)
}

fn parse_slide(
    root: &XmlNode,
    context: &NodeParseContext<'_, '_>,
) -> Result<ParsedSlideContent, ParseError> {
    let background = [
        (Some(root), Some(context.rels)),
        (context.layout_root, context.layout_rels),
        (context.master_root, context.master_rels),
    ]
    .into_iter()
    .find_map(|(part, rels)| part_background(part?, context.colors, rels?));
    let mut nodes = Vec::new();
    if master_shapes_are_visible(root)
        && context
            .layout_root
            .map(master_shapes_are_visible)
            .unwrap_or(true)
    {
        if let (Some(master), Some(master_rels)) = (context.master_root, context.master_rels) {
            let master_context = NodeParseContext {
                slide_index: context.slide_index,
                rels: master_rels,
                related_parts: context.related_parts,
                chart_companions: context.chart_companions,
                limits: context.limits,
                colors: context.colors,
                diagnostics: context.diagnostics,
                layout_root: None,
                master_root: None,
                layout_rels: None,
                master_rels: None,
                presentation_text_style: context.presentation_text_style,
                table_styles: context.table_styles,
            };
            let mut inherited = parse_decorative_part_nodes(master, &master_context)?;
            for node in &mut inherited {
                prefix_node_ids(node, "master");
            }
            nodes.extend(inherited);
        }
    }
    if let (Some(layout), Some(layout_rels)) = (context.layout_root, context.layout_rels) {
        let layout_context = NodeParseContext {
            slide_index: context.slide_index,
            rels: layout_rels,
            related_parts: context.related_parts,
            chart_companions: context.chart_companions,
            limits: context.limits,
            colors: context.colors,
            diagnostics: context.diagnostics,
            layout_root: None,
            master_root: None,
            layout_rels: None,
            master_rels: None,
            presentation_text_style: context.presentation_text_style,
            table_styles: context.table_styles,
        };
        let mut inherited = parse_decorative_part_nodes(layout, &layout_context)?;
        for node in &mut inherited {
            prefix_node_ids(node, "layout");
        }
        nodes.extend(inherited);
    }
    nodes.extend(parse_part_nodes(root, context)?);
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
        nodes,
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
        "bmp" | "dib" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "webp" => "image/webp",
        "ico" => "image/x-icon",
        "wdp" => "image/vnd.ms-photo",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        "mp4" => "video/mp4",
        "mp3" => "audio/mpeg",
        "wav" => "audio/wav",
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

fn load_theme_part(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    path: &str,
    limits: &ParseLimits,
) -> Result<Option<ThemeDataInternal>, ParseError> {
    let Some(xml) = read_entry(archive, path, limits)? else {
        return Ok(None);
    };
    let rels = match read_entry(archive, &relationship_part_path(path), limits)? {
        Some(xml) => relationships(&xml, path, limits.max_xml_depth)?,
        None => HashMap::new(),
    };
    Ok(Some(parse_theme(&xml, path, rels, limits)?))
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
    let presentation_root = parse_xml_tree(&presentation, limits, "ppt/presentation.xml")?;
    let presentation_text_style = presentation_root.child("defaultTextStyle");
    let table_styles_path =
        related_path(&rels, "ppt/tableStyles").unwrap_or_else(|| "ppt/tableStyles.xml".to_owned());
    let table_styles_root = match read_entry(&mut archive, &table_styles_path, limits)? {
        Some(xml) => Some(parse_xml_tree(&xml, limits, &table_styles_path)?),
        None => None,
    };
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
        let slide_relationship_xml =
            read_entry(&mut archive, &relationship_part_path(slide_path), limits)?;
        let slide_rels = match slide_relationship_xml.as_deref() {
            Some(bytes) => relationships(bytes, slide_path, limits.max_xml_depth)?,
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
                match load_theme_part(&mut archive, path, limits)? {
                    Some(theme) => {
                        theme_parts.insert(path.to_owned(), theme);
                    }
                    None => document_warnings.push(missing_part_warning(path, Some(index))),
                }
            }
        }
        let theme = theme_path
            .as_ref()
            .and_then(|path| theme_parts.get(path))
            .cloned();

        let mut related_paths = BTreeMap::<String, bool>::new();
        for target in slide_rels.values().filter(|target| {
            target.starts_with("ppt/charts/") || target.starts_with("ppt/diagrams/")
        }) {
            related_paths.insert(target.clone(), target.starts_with("ppt/charts/"));
        }
        if let Some(bytes) = slide_relationship_xml.as_deref() {
            for (relationship_type, target) in
                relationship_targets_by_type(bytes, slide_path, limits.max_xml_depth)?
            {
                let is_chart = is_chart_relationship_type(&relationship_type);
                if is_chart || is_diagram_relationship_type(&relationship_type) {
                    related_paths
                        .entry(target)
                        .and_modify(|existing| *existing |= is_chart)
                        .or_insert(is_chart);
                }
            }
        }
        let mut related_parts = HashMap::new();
        let mut chart_companions: HashMap<String, ChartCompanionParts> = HashMap::new();
        for (target, is_chart_part) in related_paths {
            if let Some(bytes) = read_entry(&mut archive, &target, limits)? {
                related_parts.insert(target.clone(), bytes);
                if !is_chart_part || chart_companions.contains_key(&target) {
                    continue;
                }
                let Some(rels_xml) =
                    read_entry(&mut archive, &relationship_part_path(&target), limits)?
                else {
                    continue;
                };
                let mut companions = ChartCompanionParts::default();
                for (rel_type, rel_target) in
                    relationship_targets_by_type(&rels_xml, &target, limits.max_xml_depth)?
                {
                    let slot = match rel_type.as_str() {
                        CHART_STYLE_REL_TYPE => &mut companions.style_xml,
                        CHART_COLOR_STYLE_REL_TYPE => &mut companions.colors_xml,
                        _ => continue,
                    };
                    if slot.is_none() {
                        if let Some(bytes) = read_entry(&mut archive, &rel_target, limits)? {
                            *slot = Some(String::from_utf8_lossy(&bytes).into_owned());
                        }
                    }
                }
                if companions.style_xml.is_some() || companions.colors_xml.is_some() {
                    chart_companions.insert(target, companions);
                }
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
                chart_companions: &chart_companions,
                limits,
                colors: Some(&colors),
                diagnostics: &diagnostics,
                layout_root: layout.as_ref().map(|part| &part.root),
                master_root: master.as_ref().map(|part| &part.root),
                layout_rels: layout.as_ref().map(|part| &part.rels),
                master_rels: master.as_ref().map(|part| &part.rels),
                presentation_text_style,
                table_styles: table_styles_root.as_ref(),
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
        match load_theme_part(&mut archive, &path, limits)? {
            Some(theme) => {
                theme_parts.insert(path.clone(), theme);
            }
            None => document_warnings.push(missing_part_warning(&path, None)),
        }
    }
    for target in master_parts
        .values()
        .chain(layout_parts.values())
        .flat_map(|part| part.rels.values())
        .chain(theme_parts.values().flat_map(|theme| theme.rels.values()))
        .filter(|target| target.starts_with("ppt/media/"))
    {
        asset_paths.insert(target.clone(), target.clone());
    }

    let empty_related_parts = HashMap::new();
    let empty_chart_companions = HashMap::new();
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
                chart_companions: &empty_chart_companions,
                limits,
                colors: Some(&colors),
                diagnostics: &diagnostics,
                layout_root: None,
                master_root: None,
                layout_rels: None,
                master_rels: None,
                presentation_text_style,
                table_styles: table_styles_root.as_ref(),
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
                chart_companions: &empty_chart_companions,
                limits,
                colors: Some(&colors),
                diagnostics: &diagnostics,
                layout_root: None,
                master_root: master.map(|part| &part.root),
                layout_rels: None,
                master_rels: master.map(|part| &part.rels),
                presentation_text_style,
                table_styles: table_styles_root.as_ref(),
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

#[cfg(test)]
mod tests {
    use super::*;

    fn xml(value: &[u8]) -> XmlNode {
        parse_xml_tree(value, &ParseLimits::default(), "test.xml").unwrap()
    }

    #[test]
    fn drawingml_tint_and_shade_mix_in_linear_light() {
        let tinted =
            xml(br#"<a:srgbClr xmlns:a="a" val="00FF00"><a:tint val="50000"/></a:srgbClr>"#);
        let shaded =
            xml(br#"<a:srgbClr xmlns:a="a" val="00FF00"><a:shade val="50000"/></a:srgbClr>"#);
        assert_eq!(parse_color(&tinted, None).unwrap().value, "#BCFFBC");
        assert_eq!(parse_color(&shaded, None).unwrap().value, "#00BC00");
    }

    #[test]
    fn image_fill_preserves_blip_opacity() {
        let fill = xml(
            br#"<a:blipFill xmlns:a="a" xmlns:r="r"><a:blip r:embed="rId1"><a:alphaModFix amt="14000"/></a:blip><a:stretch/></a:blipFill>"#,
        );
        let rels = HashMap::from([("rId1".into(), "ppt/media/image.png".into())]);
        assert!(matches!(
            parse_blip_fill(&fill, None, &rels),
            Some(FillStyle::Image { opacity: Some(value), .. }) if (value - 0.14).abs() < f64::EPSILON
        ));
    }

    #[test]
    fn background_fill_shapes_do_not_fall_through_to_theme_styles() {
        let shape = xml(br#"<p:sp xmlns:p="p" useBgFill="1"><p:spPr/></p:sp>"#);
        assert!(matches!(
            shape_fill(&shape, None, &HashMap::new()),
            Some(Some(FillStyle::None))
        ));
    }

    #[test]
    fn custom_geometry_is_normalized_as_svg_path_data() {
        let shape = xml(
            br#"<p:sp xmlns:p="p" xmlns:a="a"><p:spPr><a:custGeom><a:pathLst><a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="200" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom></p:spPr></p:sp>"#,
        );
        assert_eq!(
            custom_geometry_path(&shape).as_deref(),
            Some("M 0 0 L 1 1 Z")
        );
    }
}

use pptx_core::{detect_format, parse_presentation as parse_core_presentation, ParseLimits};
use serde::Serialize;
use serde_wasm_bindgen::Serializer;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub fn detect_presentation_format(bytes: &[u8]) -> Result<String, JsValue> {
    detect_format(bytes)
        .map(|format| {
            match format {
                pptx_core::PresentationFormat::Pptx => "pptx",
                pptx_core::PresentationFormat::Ppt => "ppt",
            }
            .into()
        })
        .map_err(|error| JsValue::from_str(&error.to_string()))
}

#[wasm_bindgen]
pub fn parse_presentation(bytes: &[u8]) -> Result<JsValue, JsValue> {
    let document = parse_core_presentation(bytes, &ParseLimits::default())
        .map_err(|error| JsValue::from_str(&error.to_string()))?;

    // The TypeScript model exposes keyed collections as `Record<string, T>` rather than
    // `Map<string, T>`. Keep byte buffers as Uint8Array and optional values as undefined,
    // while making every Rust map cross the Wasm boundary as a plain JavaScript object.
    document
        .serialize(&Serializer::new().serialize_maps_as_objects(true))
        .map_err(|error| JsValue::from_str(&error.to_string()))
}

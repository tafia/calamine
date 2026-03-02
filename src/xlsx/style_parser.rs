// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use quick_xml::{
    events::{attributes::Attribute, BytesStart, Event},
    name::QName,
    Reader,
};
use std::io::BufRead;

use crate::style::*;
use crate::utils::unescape_entity_to_buffer;
use crate::XlsxError;

/// Default Office 2007 theme colors, indexed by Excel's theme attribute value.
///
/// Excel swaps the first four entries relative to the clrScheme XML order:
///   theme="0" -> lt1, theme="1" -> dk1, theme="2" -> lt2, theme="3" -> dk2
pub fn default_theme_colors() -> Vec<Color> {
    vec![
        Color::rgb(255, 255, 255), // 0: lt1 (White)
        Color::rgb(0, 0, 0),       // 1: dk1 (Black)
        Color::rgb(238, 236, 225), // 2: lt2 (EEECE1)
        Color::rgb(31, 73, 125),   // 3: dk2 (1F497D)
        Color::rgb(79, 129, 189),  // 4: accent1
        Color::rgb(192, 80, 77),   // 5: accent2
        Color::rgb(155, 187, 89),  // 6: accent3
        Color::rgb(128, 100, 162), // 7: accent4
        Color::rgb(75, 172, 198),  // 8: accent5
        Color::rgb(247, 150, 70),  // 9: accent6
        Color::rgb(0, 0, 255),     // 10: hlink
        Color::rgb(128, 0, 128),   // 11: folHlink
    ]
}

/// Get theme color using the workbook's palette when available.
fn get_theme_color(theme: u8, theme_colors: &[Color]) -> Color {
    theme_colors
        .get(theme as usize)
        .copied()
        .unwrap_or_else(|| Color::rgb(0, 0, 0))
}

/// Apply tint to a color per the OOXML spec (Section 18.8.19).
///
/// Negative tint darkens:  result = channel * (1 + tint)
/// Positive tint lightens: result = channel + (255 - channel) * tint
fn apply_tint(color: Color, tint: f64) -> Color {
    let apply_channel = |c: u8| -> u8 {
        let val = if tint < 0.0 {
            c as f64 * (1.0 + tint)
        } else {
            c as f64 + (255.0 - c as f64) * tint
        };
        val.round().clamp(0.0, 255.0) as u8
    };
    Color::new(
        color.alpha,
        apply_channel(color.red),
        apply_channel(color.green),
        apply_channel(color.blue),
    )
}

/// Resolve theme colors from `xl/theme/theme1.xml`.
///
/// Reads the clrScheme element which lists colors in order:
///   dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink
///
/// Excel swaps the first four when mapping to theme indices, so:
///   theme="0" -> lt1, theme="1" -> dk1, theme="2" -> lt2, theme="3" -> dk2
pub fn read_theme_colors<RS: BufRead>(xml: &mut Reader<RS>) -> Vec<Color> {
    let mut xml_order: Vec<Color> = Vec::new();
    let mut buf = Vec::new();
    let mut in_clr_scheme = false;
    let clr_elements: &[&[u8]] = &[
        b"dk1", b"lt1", b"dk2", b"lt2", b"accent1", b"accent2", b"accent3", b"accent4",
        b"accent5", b"accent6", b"hlink", b"folHlink",
    ];
    let mut current_element: Option<Vec<u8>> = None;

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"clrScheme" {
                    in_clr_scheme = true;
                } else if in_clr_scheme {
                    if clr_elements.iter().any(|&el| local.as_ref() == el) {
                        current_element = Some(local.as_ref().to_vec());
                    } else if current_element.is_some() {
                        if let Some(color) = parse_theme_srgb_or_sys(e.attributes()) {
                            xml_order.push(color);
                            current_element = None;
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) if in_clr_scheme && current_element.is_some() => {
                if let Some(color) = parse_theme_srgb_or_sys(e.attributes()) {
                    xml_order.push(color);
                    current_element = None;
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"clrScheme" {
                    break;
                }
                if current_element
                    .as_ref()
                    .is_some_and(|el| el == local.as_ref())
                {
                    current_element = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    if xml_order.len() >= 4 {
        // Apply Excel's swap: indices 0/1 and 2/3
        xml_order.swap(0, 1);
        xml_order.swap(2, 3);
        xml_order
    } else {
        default_theme_colors()
    }
}

/// Parse an srgbClr or sysClr element's attributes into a Color.
///
/// For srgbClr: reads the `val` attribute (a 6-char hex string like "1F497D").
/// For sysClr: prefers `lastClr` (resolved color) over `val` (system name like "window").
fn parse_theme_srgb_or_sys(
    attributes: quick_xml::events::attributes::Attributes,
) -> Option<Color> {
    let mut color: Option<Color> = None;

    for attr in attributes.flatten() {
        match attr.key.as_ref() {
            b"lastClr" => {
                if let Some(c) = parse_hex_attr(attr.value.as_ref()) {
                    return Some(c);
                }
            }
            b"val" => {
                if color.is_none() {
                    color = parse_hex_attr(attr.value.as_ref());
                }
            }
            _ => {}
        }
    }

    color
}

fn parse_hex_attr(val: &[u8]) -> Option<Color> {
    if val.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&String::from_utf8_lossy(&val[0..2]), 16).ok()?;
    let g = u8::from_str_radix(&String::from_utf8_lossy(&val[2..4]), 16).ok()?;
    let b = u8::from_str_radix(&String::from_utf8_lossy(&val[4..6]), 16).ok()?;
    Some(Color::rgb(r, g, b))
}

/// Get indexed color from Excel's official color index palette
/// Based on: https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
fn get_indexed_color(index: u8) -> Color {
    match index {
        1 => Color::rgb(0, 0, 0),       // Black
        2 => Color::rgb(255, 255, 255), // White
        3 => Color::rgb(255, 0, 0),     // Red
        4 => Color::rgb(0, 255, 0),     // Green
        5 => Color::rgb(0, 0, 255),     // Blue
        6 => Color::rgb(255, 255, 0),   // Yellow
        7 => Color::rgb(255, 0, 255),   // Magenta
        8 => Color::rgb(0, 255, 255),   // Cyan
        9 => Color::rgb(128, 0, 0),      // Dark Red
        10 => Color::rgb(0, 128, 0),     // Dark Green
        11 => Color::rgb(0, 0, 128),     // Dark Blue
        12 => Color::rgb(128, 128, 0),   // Dark Yellow
        13 => Color::rgb(128, 0, 128),   // Dark Magenta
        14 => Color::rgb(0, 128, 128),   // Dark Cyan
        15 => Color::rgb(192, 192, 192), // Light Gray
        16 => Color::rgb(128, 128, 128), // Gray
        17 => Color::rgb(153, 153, 255), // Light Blue
        18 => Color::rgb(153, 51, 102),  // Dark Pink
        19 => Color::rgb(255, 255, 204), // Light Yellow
        20 => Color::rgb(204, 255, 255), // Light Cyan
        21 => Color::rgb(102, 0, 102),   // Dark Purple
        22 => Color::rgb(255, 128, 128), // Light Red
        23 => Color::rgb(0, 102, 204),   // Medium Blue
        24 => Color::rgb(204, 204, 255), // Light Purple
        25 => Color::rgb(0, 0, 128),   // Navy
        26 => Color::rgb(255, 0, 255), // Fuchsia
        27 => Color::rgb(255, 255, 0), // Yellow
        28 => Color::rgb(0, 255, 255), // Aqua
        29 => Color::rgb(128, 0, 128), // Purple
        30 => Color::rgb(128, 0, 0),   // Maroon
        31 => Color::rgb(0, 128, 128), // Teal
        32 => Color::rgb(0, 0, 255),   // Blue
        33 => Color::rgb(0, 204, 255),   // Sky Blue
        34 => Color::rgb(204, 255, 255), // Light Turquoise
        35 => Color::rgb(204, 255, 204), // Light Green
        36 => Color::rgb(255, 255, 153), // Light Yellow
        37 => Color::rgb(153, 204, 255), // Pale Blue
        38 => Color::rgb(255, 153, 204), // Pink
        39 => Color::rgb(204, 153, 255), // Lavender
        40 => Color::rgb(255, 204, 153), // Tan
        41 => Color::rgb(51, 102, 255),  // Bright Blue
        42 => Color::rgb(51, 204, 204),  // Aqua
        43 => Color::rgb(153, 204, 0),   // Lime
        44 => Color::rgb(255, 204, 0),   // Gold
        45 => Color::rgb(255, 153, 0),   // Orange
        46 => Color::rgb(255, 102, 0),   // Orange Red
        47 => Color::rgb(102, 102, 153), // Blue Gray
        48 => Color::rgb(150, 150, 150), // Gray 40%
        49 => Color::rgb(0, 51, 102),   // Dark Teal
        50 => Color::rgb(51, 153, 102), // Sea Green
        51 => Color::rgb(0, 51, 0),     // Dark Green
        52 => Color::rgb(51, 51, 0),    // Olive
        53 => Color::rgb(153, 51, 0),   // Brown
        54 => Color::rgb(153, 51, 102), // Plum
        55 => Color::rgb(51, 51, 153),  // Indigo
        56 => Color::rgb(51, 51, 51),   // Gray 80%
        0 => Color::rgb(0, 0, 0),        // Auto (Black)
        64 => Color::rgb(192, 192, 192), // System window background
        65 => Color::rgb(0, 0, 0),       // System auto color
        _ => Color::rgb(0, 0, 0), // Black for unknown indices
    }
}

/// Parse color from XML attributes, resolving theme/indexed references and
/// applying the tint modifier when present.
fn parse_color(
    attributes: &[Attribute],
    theme_colors: &[Color],
) -> Result<Option<Color>, XlsxError> {
    let mut color: Option<Color> = None;
    let mut tint: Option<f64> = None;

    for attr in attributes {
        match attr.key.as_ref() {
            b"rgb" => {
                let rgb_str: &[u8] = attr.value.as_ref();
                if rgb_str.len() == 6 {
                    let r = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[0..2]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid red color value"))?;
                    let g = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[2..4]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid green color value"))?;
                    let b = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[4..6]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid blue color value"))?;
                    color = Some(Color::rgb(r, g, b));
                } else if rgb_str.len() == 8 {
                    let a = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[0..2]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid alpha color value"))?;
                    let r = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[2..4]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid red color value"))?;
                    let g = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[4..6]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid green color value"))?;
                    let b = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[6..8]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid blue color value"))?;
                    color = Some(Color::new(a, r, g, b));
                }
            }
            b"theme" => {
                let theme_str = String::from_utf8_lossy(&attr.value);
                if let Ok(theme_value) = theme_str.parse::<u8>() {
                    color = Some(get_theme_color(theme_value, theme_colors));
                }
            }
            b"indexed" => {
                let indexed_str = String::from_utf8_lossy(&attr.value);
                if let Ok(indexed_value) = indexed_str.parse::<u8>() {
                    color = Some(get_indexed_color(indexed_value));
                }
            }
            b"tint" => {
                let tint_str = String::from_utf8_lossy(&attr.value);
                tint = tint_str.parse::<f64>().ok();
            }
            _ => {}
        }
    }

    Ok(color.map(|c| match tint {
        Some(t) if t != 0.0 => apply_tint(c, t),
        _ => c,
    }))
}

/// Parse font weight from string
fn parse_font_weight(s: &str) -> FontWeight {
    match s {
        "bold" | "700" => FontWeight::Bold,
        "normal" | "400" => FontWeight::Normal,
        _ => {
            if let Ok(weight) = s.parse::<u16>() {
                if weight >= 600 {
                    FontWeight::Bold
                } else {
                    FontWeight::Normal
                }
            } else {
                FontWeight::Normal
            }
        }
    }
}

fn parse_font_style(s: &str) -> FontStyle {
    match s {
        "italic" | "oblique" => FontStyle::Italic,
        _ => FontStyle::Normal,
    }
}

fn parse_underline_style(s: &str) -> UnderlineStyle {
    match s {
        "single" => UnderlineStyle::Single,
        "double" => UnderlineStyle::Double,
        "singleAccounting" => UnderlineStyle::SingleAccounting,
        "doubleAccounting" => UnderlineStyle::DoubleAccounting,
        _ => UnderlineStyle::None,
    }
}

fn parse_horizontal_alignment(s: &str) -> HorizontalAlignment {
    match s {
        "left" => HorizontalAlignment::Left,
        "center" => HorizontalAlignment::Center,
        "right" => HorizontalAlignment::Right,
        "justify" => HorizontalAlignment::Justify,
        "distributed" => HorizontalAlignment::Distributed,
        "fill" => HorizontalAlignment::Fill,
        _ => HorizontalAlignment::General,
    }
}

fn parse_vertical_alignment(s: &str) -> VerticalAlignment {
    match s {
        "top" => VerticalAlignment::Top,
        "center" => VerticalAlignment::Center,
        "bottom" => VerticalAlignment::Bottom,
        "justify" => VerticalAlignment::Justify,
        "distributed" => VerticalAlignment::Distributed,
        _ => VerticalAlignment::Bottom,
    }
}

fn parse_fill_pattern(s: &str) -> FillPattern {
    match s {
        "solid" => FillPattern::Solid,
        "darkGray" => FillPattern::DarkGray,
        "mediumGray" => FillPattern::MediumGray,
        "lightGray" => FillPattern::LightGray,
        "gray125" => FillPattern::Gray125,
        "gray0625" => FillPattern::Gray0625,
        "darkHorizontal" => FillPattern::DarkHorizontal,
        "darkVertical" => FillPattern::DarkVertical,
        "darkDown" => FillPattern::DarkDown,
        "darkUp" => FillPattern::DarkUp,
        "darkGrid" => FillPattern::DarkGrid,
        "darkTrellis" => FillPattern::DarkTrellis,
        "lightHorizontal" => FillPattern::LightHorizontal,
        "lightVertical" => FillPattern::LightVertical,
        "lightDown" => FillPattern::LightDown,
        "lightUp" => FillPattern::LightUp,
        "lightGrid" => FillPattern::LightGrid,
        "lightTrellis" => FillPattern::LightTrellis,
        _ => FillPattern::None,
    }
}

fn parse_border_style(s: &str) -> BorderStyle {
    match s {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "slantDashDot" => BorderStyle::SlantDashDot,
        _ => BorderStyle::None,
    }
}

/// Parse font element
pub fn parse_font<RS: BufRead>(
    xml: &mut Reader<RS>,
    _start_elem: &BytesStart,
    theme_colors: &[Color],
) -> Result<Font, XlsxError> {
    let mut font = Font::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"name" => {
                    let mut name = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            name = Some(String::from_utf8_lossy(&attr.value).to_string());
                            break;
                        }
                    }
                    if name.is_none() {
                        name = read_string(xml, QName(b"name"))?;
                    } else {
                        xml.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                    if let Some(n) = name {
                        font = font.with_name(n);
                    }
                }
                b"sz" => {
                    let mut size_str = None;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            size_str = Some(String::from_utf8_lossy(&attr.value).to_string());
                            break;
                        }
                    }
                    if size_str.is_none() {
                        size_str = read_string(xml, QName(b"sz"))?;
                    } else {
                        xml.read_to_end_into(e.name(), &mut Vec::new())?;
                    }
                    if let Some(s) = size_str {
                        if let Ok(size) = s.parse::<f64>() {
                            font = font.with_size(size);
                        }
                    }
                }
                b"b" => {
                    let mut weight = FontWeight::Bold;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            let val_str = String::from_utf8_lossy(&attr.value);
                            weight = parse_font_weight(&val_str);
                            break;
                        }
                    }
                    font = font.with_weight(weight);
                }
                b"i" => {
                    let mut style = FontStyle::Italic;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            let val_str = String::from_utf8_lossy(&attr.value);
                            style = parse_font_style(&val_str);
                            break;
                        }
                    }
                    font = font.with_style(style);
                }
                b"u" => {
                    let mut underline_style = UnderlineStyle::Single;
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"val" {
                            let val_str = String::from_utf8_lossy(&attr.value);
                            underline_style = parse_underline_style(&val_str);
                            break;
                        }
                    }
                    font = font.with_underline(underline_style);
                }
                b"strike" => {
                    font = font.with_strikethrough(true);
                }
                b"color" => {
                    if let Some(color) = parse_color(
                        &e.attributes().collect::<Result<Vec<_>, _>>()?,
                        theme_colors,
                    )? {
                        font = font.with_color(color);
                    }
                }
                b"family" => {
                    if let Some(family) = read_string(xml, QName(b"family"))? {
                        font = font.with_family(family);
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"font" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("font")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(font)
}

/// Parse fill element
pub fn parse_fill<RS: BufRead>(
    xml: &mut Reader<RS>,
    _start_elem: &BytesStart,
    theme_colors: &[Color],
) -> Result<Fill, XlsxError> {
    let mut fill = Fill::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"patternFill" => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"patternType" {
                            let pattern_str = String::from_utf8_lossy(&attr.value);
                            fill = fill.with_pattern(parse_fill_pattern(&pattern_str));
                        }
                    }
                }
                b"fgColor" => {
                    if let Some(color) = parse_color(
                        &e.attributes().collect::<Result<Vec<_>, _>>()?,
                        theme_colors,
                    )? {
                        fill = fill.with_foreground_color(color);
                    }
                }
                b"bgColor" => {
                    if let Some(color) = parse_color(
                        &e.attributes().collect::<Result<Vec<_>, _>>()?,
                        theme_colors,
                    )? {
                        fill = fill.with_background_color(color);
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fill" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("fill")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(fill)
}

/// Parse border element
pub fn parse_border<RS: BufRead>(
    xml: &mut Reader<RS>,
    _start_elem: &BytesStart,
    theme_colors: &[Color],
) -> Result<Borders, XlsxError> {
    let mut borders = Borders::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                    b"left" | b"right" | b"top" | b"bottom" | b"diagonal" => {
                        let mut style = BorderStyle::None;
                        let mut color = None;

                        for attr in e.attributes() {
                            let attr = attr?;
                            if attr.key.as_ref() == b"style" {
                                let style_str = String::from_utf8_lossy(&attr.value);
                                style = parse_border_style(&style_str);
                            }
                        }

                        if let Some(border_color) = parse_color(
                            &e.attributes().collect::<Result<Vec<_>, _>>()?,
                            theme_colors,
                        )? {
                            color = Some(border_color);
                        }

                        let mut inner_buf = Vec::new();
                        loop {
                            inner_buf.clear();
                            match xml.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref inner_e)) => {
                                    if inner_e.local_name().as_ref() == b"color" {
                                        if let Some(border_color) = parse_color(
                                            &inner_e
                                                .attributes()
                                                .collect::<Result<Vec<_>, _>>()?,
                                            theme_colors,
                                        )? {
                                            color = Some(border_color);
                                        }
                                    }
                                }
                                Ok(Event::End(ref inner_e))
                                    if inner_e.local_name().as_ref()
                                        == e.local_name().as_ref() =>
                                {
                                    break
                                }
                                Ok(Event::Eof) => {
                                    return Err(XlsxError::XmlEof("border side"))
                                }
                                Err(e) => return Err(XlsxError::Xml(e)),
                                _ => {}
                            }
                        }

                        let border = if let Some(c) = color {
                            Border::with_color(style, c)
                        } else {
                            Border::new(style)
                        };

                        match e.local_name().as_ref() {
                            b"left" => borders.left = border,
                            b"right" => borders.right = border,
                            b"top" => borders.top = border,
                            b"bottom" => borders.bottom = border,
                            b"diagonal" => {
                                for attr in e.attributes() {
                                    let attr = attr?;
                                    if attr.key.as_ref() == b"diagonalDown" {
                                        borders.diagonal_down = border.clone();
                                    } else if attr.key.as_ref() == b"diagonalUp" {
                                        borders.diagonal_up = border.clone();
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"border" => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("border")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(borders)
}

/// Parse alignment element
pub fn parse_alignment<RS: BufRead>(
    _xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Alignment, XlsxError> {
    let mut alignment = Alignment::new();

    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"horizontal" => {
                let horizontal_str = String::from_utf8_lossy(&attr.value);
                alignment = alignment.with_horizontal(parse_horizontal_alignment(&horizontal_str));
            }
            b"vertical" => {
                let vertical_str = String::from_utf8_lossy(&attr.value);
                alignment = alignment.with_vertical(parse_vertical_alignment(&vertical_str));
            }
            b"wrapText" => {
                let wrap_str = String::from_utf8_lossy(&attr.value);
                if wrap_str == "1" || wrap_str == "true" {
                    alignment = alignment.with_wrap_text(true);
                }
            }
            b"textRotation" => {
                if let Ok(rotation) = String::from_utf8_lossy(&attr.value).parse::<u16>() {
                    alignment = alignment.with_text_rotation(TextRotation::Degrees(rotation));
                }
            }
            b"indent" => {
                if let Ok(indent) = String::from_utf8_lossy(&attr.value).parse::<u8>() {
                    alignment = alignment.with_indent(indent);
                }
            }
            b"shrinkToFit" => {
                let shrink_str = String::from_utf8_lossy(&attr.value);
                if shrink_str == "1" || shrink_str == "true" {
                    alignment = alignment.with_shrink_to_fit(true);
                }
            }
            _ => {}
        }
    }

    Ok(alignment)
}

/// Parse protection element
pub fn parse_protection<RS: BufRead>(
    _xml: &mut Reader<RS>,
    start_elem: &BytesStart,
) -> Result<Protection, XlsxError> {
    let mut protection = Protection::new();

    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"locked" => {
                let locked_str = String::from_utf8_lossy(&attr.value);
                if locked_str == "1" || locked_str == "true" {
                    protection = protection.with_locked(true);
                }
            }
            b"hidden" => {
                let hidden_str = String::from_utf8_lossy(&attr.value);
                if hidden_str == "1" || hidden_str == "true" {
                    protection = protection.with_hidden(true);
                }
            }
            _ => {}
        }
    }

    Ok(protection)
}

/// Read string content from XML element
fn read_string<RS: BufRead>(
    xml: &mut Reader<RS>,
    closing: QName,
) -> Result<Option<String>, XlsxError> {
    let mut buf = Vec::new();
    let mut content = String::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Text(e)) => {
                content.push_str(&e.xml10_content()?);
            }
            Ok(Event::GeneralRef(e)) => {
                unescape_entity_to_buffer(&e, &mut content)?;
            }
            Ok(Event::End(ref e)) if e.local_name() == closing.into() => break,
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("string")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    if content.is_empty() {
        Ok(None)
    } else {
        Ok(Some(content))
    }
}

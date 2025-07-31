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
use crate::XlsxError;

/// Get theme color by index (Excel's default theme colors)
fn get_theme_color(theme_index: u8) -> Color {
    match theme_index {
        0 => Color::rgb(0, 0, 0),       // dk1 - dark 1 (black/window text)
        1 => Color::rgb(255, 255, 255), // lt1 - light 1 (white/window background)
        2 => Color::rgb(31, 73, 125),   // dk2 - dark 2 (dark blue)
        3 => Color::rgb(238, 236, 225), // lt2 - light 2 (light cream)
        4 => Color::rgb(79, 129, 189),  // accent1 - blue
        5 => Color::rgb(192, 80, 77),   // accent2 - red
        6 => Color::rgb(155, 187, 89),  // accent3 - green
        7 => Color::rgb(128, 100, 162), // accent4 - purple
        8 => Color::rgb(75, 172, 198),  // accent5 - cyan
        9 => Color::rgb(247, 150, 70),  // accent6 - orange
        10 => Color::rgb(0, 0, 255),    // hlink - hyperlink (blue)
        11 => Color::rgb(128, 0, 128),  // folHlink - followed hyperlink (purple)
        _ => Color::rgb(0, 0, 0),       // Default to black for unknown theme colors
    }
}

/// Get indexed color by value (Excel's default indexed color palette)
fn get_indexed_color(indexed_value: u8) -> Color {
    match indexed_value {
        // System colors (0-7)
        0 => Color::rgb(0, 0, 0),       // Black
        1 => Color::rgb(255, 255, 255), // White
        2 => Color::rgb(255, 0, 0),     // Red
        3 => Color::rgb(0, 255, 0),     // Green
        4 => Color::rgb(0, 0, 255),     // Blue
        5 => Color::rgb(255, 255, 0),   // Yellow
        6 => Color::rgb(255, 0, 255),   // Magenta
        7 => Color::rgb(0, 255, 255),   // Cyan

        // Default palette colors (8-63)
        8 => Color::rgb(0, 0, 0),        // Black
        9 => Color::rgb(255, 255, 255),  // White
        10 => Color::rgb(255, 0, 0),     // Red
        11 => Color::rgb(0, 255, 0),     // Green
        12 => Color::rgb(0, 0, 255),     // Blue
        13 => Color::rgb(255, 255, 0),   // Yellow
        14 => Color::rgb(255, 0, 255),   // Magenta
        15 => Color::rgb(0, 255, 255),   // Cyan
        16 => Color::rgb(128, 0, 0),     // Dark Red
        17 => Color::rgb(0, 128, 0),     // Dark Green
        18 => Color::rgb(0, 0, 128),     // Dark Blue
        19 => Color::rgb(128, 128, 0),   // Dark Yellow
        20 => Color::rgb(128, 0, 128),   // Dark Magenta
        21 => Color::rgb(0, 128, 128),   // Dark Cyan
        22 => Color::rgb(192, 192, 192), // Light Gray
        23 => Color::rgb(128, 128, 128), // Gray
        24 => Color::rgb(153, 153, 255), // Light Blue
        25 => Color::rgb(153, 51, 102),  // Dark Pink
        26 => Color::rgb(255, 255, 204), // Light Yellow
        27 => Color::rgb(204, 255, 255), // Light Cyan
        28 => Color::rgb(102, 0, 102),   // Dark Purple
        29 => Color::rgb(255, 128, 128), // Light Red
        30 => Color::rgb(0, 102, 204),   // Medium Blue
        31 => Color::rgb(204, 204, 255), // Light Purple
        32 => Color::rgb(0, 0, 128),     // Dark Blue
        33 => Color::rgb(255, 0, 255),   // Magenta
        34 => Color::rgb(255, 255, 0),   // Yellow
        35 => Color::rgb(0, 255, 255),   // Cyan
        36 => Color::rgb(128, 0, 128),   // Purple
        37 => Color::rgb(128, 0, 0),     // Maroon
        38 => Color::rgb(0, 128, 128),   // Teal
        39 => Color::rgb(0, 0, 255),     // Blue
        40 => Color::rgb(0, 204, 255),   // Sky Blue
        41 => Color::rgb(204, 255, 255), // Light Turquoise
        42 => Color::rgb(204, 255, 204), // Light Green
        43 => Color::rgb(255, 255, 153), // Light Yellow
        44 => Color::rgb(153, 204, 255), // Pale Blue
        45 => Color::rgb(255, 153, 204), // Pink
        46 => Color::rgb(204, 153, 255), // Lavender
        47 => Color::rgb(255, 204, 153), // Tan
        48 => Color::rgb(51, 102, 255),  // Light Blue
        49 => Color::rgb(51, 204, 204),  // Aqua
        50 => Color::rgb(153, 204, 0),   // Lime
        51 => Color::rgb(255, 204, 0),   // Gold
        52 => Color::rgb(255, 153, 0),   // Orange
        53 => Color::rgb(255, 102, 0),   // Orange Red
        54 => Color::rgb(102, 102, 153), // Blue Gray
        55 => Color::rgb(150, 150, 150), // Gray 40%
        56 => Color::rgb(0, 51, 102),    // Dark Teal
        57 => Color::rgb(51, 153, 102),  // Sea Green
        58 => Color::rgb(0, 51, 0),      // Dark Green
        59 => Color::rgb(51, 51, 0),     // Olive
        60 => Color::rgb(153, 51, 0),    // Brown
        61 => Color::rgb(153, 51, 102),  // Plum
        62 => Color::rgb(51, 51, 153),   // Indigo
        63 => Color::rgb(51, 51, 51),    // Gray 80%

        // Special values
        64 => Color::rgb(192, 192, 192), // System window background
        65 => Color::rgb(0, 0, 0),       // System auto color

        _ => Color::rgb(0, 0, 0), // Default to black for unknown indexed colors
    }
}

/// Parse color from XML attributes
fn parse_color(attributes: &[Attribute]) -> Result<Option<Color>, XlsxError> {
    for attr in attributes {
        match attr.key.as_ref() {
            b"rgb" => {
                let rgb_str = attr.value.as_ref();
                if rgb_str.len() == 6 {
                    // RGB format (6 characters)
                    let r = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[0..2]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid red color value"))?;
                    let g = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[2..4]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid green color value"))?;
                    let b = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[4..6]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid blue color value"))?;
                    return Ok(Some(Color::rgb(r, g, b)));
                } else if rgb_str.len() == 8 {
                    // ARGB format (8 characters)
                    let a = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[0..2]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid alpha color value"))?;
                    let r = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[2..4]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid red color value"))?;
                    let g = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[4..6]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid green color value"))?;
                    let b = u8::from_str_radix(&String::from_utf8_lossy(&rgb_str[6..8]), 16)
                        .map_err(|_| XlsxError::Unexpected("Invalid blue color value"))?;
                    return Ok(Some(Color::new(a, r, g, b)));
                }
            }
            b"theme" => {
                let theme_str = String::from_utf8_lossy(&attr.value);
                if let Ok(theme_index) = theme_str.parse::<u8>() {
                    // Convert theme color index to actual color using Excel's default theme colors
                    return Ok(Some(get_theme_color(theme_index)));
                }
            }
            b"indexed" => {
                let indexed_str = String::from_utf8_lossy(&attr.value);
                if let Ok(indexed_value) = indexed_str.parse::<u8>() {
                    // Convert indexed color to actual color using Excel's indexed color palette
                    return Ok(Some(get_indexed_color(indexed_value)));
                }
            }
            _ => {}
        }
    }
    Ok(None)
}

/// Parse font weight from string
fn parse_font_weight(s: &str) -> FontWeight {
    match s {
        "bold" | "700" => FontWeight::Bold,
        "normal" | "400" => FontWeight::Normal,
        _ => {
            // Try to parse as numeric weight
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

/// Parse font style from string
fn parse_font_style(s: &str) -> FontStyle {
    match s {
        "italic" | "oblique" => FontStyle::Italic,
        "normal" => FontStyle::Normal,
        _ => FontStyle::Normal,
    }
}

/// Parse underline style from string
fn parse_underline_style(s: &str) -> UnderlineStyle {
    match s {
        "single" => UnderlineStyle::Single,
        "double" => UnderlineStyle::Double,
        "singleAccounting" => UnderlineStyle::SingleAccounting,
        "doubleAccounting" => UnderlineStyle::DoubleAccounting,
        _ => UnderlineStyle::None,
    }
}

/// Parse horizontal alignment from string
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

/// Parse vertical alignment from string
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

/// Parse fill pattern from string
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

/// Parse border style from string
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
    start_elem: &BytesStart,
) -> Result<Font, XlsxError> {
    let mut font = Font::new();

    // Parse attributes from the opening font element
    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            // Font elements can have attributes like outline, shadow, etc.
            // Add specific font attribute parsing here if needed
            _ => {}
        }
    }

    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"name" => {
                    if let Some(name) = read_string(xml, QName(b"name"))? {
                        font = font.with_name(name);
                    }
                }
                b"sz" => {
                    if let Some(size_str) = read_string(xml, QName(b"sz"))? {
                        if let Ok(size) = size_str.parse::<f64>() {
                            font = font.with_size(size);
                        }
                    }
                }
                b"b" => {
                    // Check if the element has a 'val' attribute
                    let mut weight = FontWeight::Bold; // Default to bold
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
                    // Check if the element has a 'val' attribute
                    let mut style = FontStyle::Italic; // Default to italic
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
                    // Check if the element has a 'val' attribute
                    let mut underline_style = UnderlineStyle::Single; // Default to single underline
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
                    if let Some(color) =
                        parse_color(&e.attributes().collect::<Result<Vec<_>, _>>()?)?
                    {
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
    start_elem: &BytesStart,
) -> Result<Fill, XlsxError> {
    let mut fill = Fill::new();

    // Parse attributes from the opening fill element
    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            // Fill elements can have attributes like type, etc.
            // Add specific fill attribute parsing here if needed
            _ => {}
        }
    }

    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"patternFill" => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        match attr.key.as_ref() {
                            b"patternType" => {
                                let pattern_str = String::from_utf8_lossy(&attr.value);
                                fill = fill.with_pattern(parse_fill_pattern(&pattern_str));
                            }
                            _ => {}
                        }
                    }
                }
                b"fgColor" => {
                    if let Some(color) =
                        parse_color(&e.attributes().collect::<Result<Vec<_>, _>>()?)?
                    {
                        fill = fill.with_foreground_color(color);
                    }
                }
                b"bgColor" => {
                    if let Some(color) =
                        parse_color(&e.attributes().collect::<Result<Vec<_>, _>>()?)?
                    {
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
    start_elem: &BytesStart,
) -> Result<Borders, XlsxError> {
    let mut borders = Borders::new();

    // Parse attributes from the opening border element
    for attr in start_elem.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            // Border elements can have attributes like diagonalUp, diagonalDown, etc.
            // Add specific border attribute parsing here if needed
            _ => {}
        }
    }

    let mut buf = Vec::new();

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                    b"left" | b"right" | b"top" | b"bottom" | b"diagonal" => {
                        let mut style = BorderStyle::None;
                        let mut color = None;

                        // Parse attributes for style
                        for attr in e.attributes() {
                            let attr = attr?;
                            match attr.key.as_ref() {
                                b"style" => {
                                    let style_str = String::from_utf8_lossy(&attr.value);
                                    style = parse_border_style(&style_str);
                                }
                                _ => {}
                            }
                        }

                        // Check for color attributes directly on the border element (fallback)
                        if let Some(border_color) =
                            parse_color(&e.attributes().collect::<Result<Vec<_>, _>>()?)?
                        {
                            color = Some(border_color);
                        }

                        // Parse nested elements (primarily for color)
                        let mut inner_buf = Vec::new();
                        loop {
                            inner_buf.clear();
                            match xml.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref inner_e)) => {
                                    match inner_e.local_name().as_ref() {
                                        b"color" => {
                                            if let Some(border_color) = parse_color(
                                                &inner_e
                                                    .attributes()
                                                    .collect::<Result<Vec<_>, _>>()?,
                                            )? {
                                                color = Some(border_color);
                                            }
                                        }
                                        _ => {}
                                    }
                                }
                                Ok(Event::End(ref inner_e))
                                    if inner_e.local_name().as_ref() == e.local_name().as_ref() =>
                                {
                                    break
                                }
                                Ok(Event::Eof) => return Err(XlsxError::XmlEof("border side")),
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
                                // Check if it's diagonal down or up
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
                content.push_str(&e.unescape()?);
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

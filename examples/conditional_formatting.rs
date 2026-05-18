// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use calamine::{open_workbook, ConditionalFormatRuleType, Reader, Xlsx};

/// Example demonstrating how to read conditional formatting rules from Excel files.
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = format!(
        "{}/tests/conditional_formatting.xlsx",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    let sheet_names = workbook.sheet_names();
    let Some(sheet_name) = sheet_names.first() else {
        println!("No sheets found");
        return Ok(());
    };

    println!("Conditional formatting for sheet: {sheet_name}\n");

    let cfs = workbook.worksheet_conditional_formatting(sheet_name)?;

    for (i, cf) in cfs.iter().enumerate() {
        println!(
            "Block {}: range \"{}\" ({} rule(s))",
            i + 1,
            cf.sqref,
            cf.rules.len()
        );

        for rule in &cf.rules {
            print!("  [priority={}] ", rule.priority);

            match &rule.rule_type {
                ConditionalFormatRuleType::CellIs { operator, formulas } => {
                    println!("CellIs {:?} {:?}", operator, formulas);
                }
                ConditionalFormatRuleType::ColorScale2 {
                    min,
                    max,
                    min_color,
                    max_color,
                } => {
                    println!(
                        "2-Color Scale: {:?}({:?}) {} -> {:?}({:?}) {}",
                        min.value_type, min.value, min_color, max.value_type, max.value, max_color,
                    );
                }
                ConditionalFormatRuleType::ColorScale3 {
                    min,
                    mid,
                    max,
                    min_color,
                    mid_color,
                    max_color,
                } => {
                    println!(
                        "3-Color Scale: {:?}({:?}) {} -> {:?}({:?}) {} -> {:?}({:?}) {}",
                        min.value_type,
                        min.value,
                        min_color,
                        mid.value_type,
                        mid.value,
                        mid_color,
                        max.value_type,
                        max.value,
                        max_color,
                    );
                }
                ConditionalFormatRuleType::DataBar {
                    min,
                    max,
                    fill_color,
                    ..
                } => {
                    println!(
                        "DataBar: {:?}({:?}) -> {:?}({:?}), color={:?}",
                        min.value_type, min.value, max.value_type, max.value, fill_color,
                    );
                }
                ConditionalFormatRuleType::IconSet {
                    icon_type,
                    thresholds,
                    reversed,
                    show_value,
                } => {
                    println!(
                        "IconSet {:?} ({} thresholds, reversed={}, show_value={})",
                        icon_type,
                        thresholds.len(),
                        reversed,
                        show_value,
                    );
                }
                ConditionalFormatRuleType::Top10 {
                    rank,
                    percent,
                    bottom,
                } => {
                    let direction = if *bottom { "Bottom" } else { "Top" };
                    let unit = if *percent { "%" } else { "" };
                    println!("{direction} {rank}{unit}");
                }
                ConditionalFormatRuleType::AboveAverage {
                    above_average,
                    equal_average,
                    std_dev,
                } => {
                    let dir = if *above_average { "Above" } else { "Below" };
                    let eq = if *equal_average { " or equal to" } else { "" };
                    print!("{dir}{eq} average");
                    if *std_dev > 0 {
                        print!(" (std_dev={})", std_dev);
                    }
                    println!();
                }
                ConditionalFormatRuleType::Text { operator, text, .. } => {
                    println!("Text {:?} \"{}\"", operator, text);
                }
                ConditionalFormatRuleType::TimePeriod { period, .. } => {
                    println!("TimePeriod {:?}", period);
                }
                ConditionalFormatRuleType::Expression { formula } => {
                    println!("Expression: {formula}");
                }
                ConditionalFormatRuleType::DuplicateValues => println!("DuplicateValues"),
                ConditionalFormatRuleType::UniqueValues => println!("UniqueValues"),
                ConditionalFormatRuleType::ContainsBlanks { .. } => println!("ContainsBlanks"),
                ConditionalFormatRuleType::NotContainsBlanks { .. } => {
                    println!("NotContainsBlanks")
                }
                ConditionalFormatRuleType::ContainsErrors { .. } => println!("ContainsErrors"),
                ConditionalFormatRuleType::NotContainsErrors { .. } => {
                    println!("NotContainsErrors")
                }
                ConditionalFormatRuleType::Unknown { raw_type } => {
                    println!("Unknown type: {raw_type}");
                }
            }

            if let Some(fmt) = &rule.format {
                if let Some(font) = &fmt.font {
                    if font.is_bold() {
                        print!("    format: bold");
                    }
                    if let Some(color) = font.color {
                        print!("    font_color: {color}");
                    }
                    println!();
                }
                if let Some(fill) = &fmt.fill {
                    if let Some(color) = fill.get_color() {
                        println!("    fill_color: {color}");
                    }
                }
            }
        }
        println!();
    }

    println!("Total: {} conditional formatting blocks", cfs.len());
    Ok(())
}

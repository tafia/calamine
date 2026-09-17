// SPDX-License-Identifier: MIT

use crate::style::{Color, Style};

/// A `<conditionalFormatting>` element from a worksheet, containing one or more rules
/// that apply to a cell range.
#[derive(Debug, Clone, PartialEq)]
pub struct ConditionalFormatting {
    /// The cell range this formatting applies to (e.g. "A1:B10").
    pub sqref: String,
    /// The rules within this conditional formatting block, ordered by priority.
    pub rules: Vec<ConditionalFormatRule>,
}

/// A single `<cfRule>` element within a conditional formatting block.
#[derive(Debug, Clone, PartialEq)]
pub struct ConditionalFormatRule {
    /// Evaluation priority (lower = higher priority).
    pub priority: u32,
    /// If true, stop evaluating lower-priority rules when this rule matches.
    pub stop_if_true: bool,
    /// The differential format to apply when the rule matches (resolved from dxfId).
    pub format: Option<Style>,
    /// The rule type and its parameters.
    pub rule_type: ConditionalFormatRuleType,
}

/// The specific type and parameters of a conditional format rule.
#[derive(Debug, Clone, PartialEq)]
pub enum ConditionalFormatRuleType {
    /// Cell value comparison (cfRule type="cellIs").
    CellIs {
        /// The comparison operator.
        operator: CfOperator,
        /// One or two formula strings (two for Between/NotBetween).
        formulas: Vec<String>,
    },
    /// Formula-based rule (cfRule type="expression").
    Expression {
        /// The formula that determines whether the rule applies.
        formula: String,
    },
    /// 2-color scale (cfRule type="colorScale" with 2 stops).
    ColorScale2 {
        /// Minimum threshold.
        min: CfValueObject,
        /// Maximum threshold.
        max: CfValueObject,
        /// Color for the minimum value.
        min_color: Color,
        /// Color for the maximum value.
        max_color: Color,
    },
    /// 3-color scale (cfRule type="colorScale" with 3 stops).
    ColorScale3 {
        /// Minimum threshold.
        min: CfValueObject,
        /// Midpoint threshold.
        mid: CfValueObject,
        /// Maximum threshold.
        max: CfValueObject,
        /// Color for the minimum value.
        min_color: Color,
        /// Color for the midpoint value.
        mid_color: Color,
        /// Color for the maximum value.
        max_color: Color,
    },
    /// Data bar (cfRule type="dataBar").
    DataBar {
        /// Minimum threshold.
        min: CfValueObject,
        /// Maximum threshold.
        max: CfValueObject,
        /// Bar fill color.
        fill_color: Option<Color>,
        /// Bar border color.
        border_color: Option<Color>,
    },
    /// Icon set (cfRule type="iconSet").
    IconSet {
        /// The icon set style.
        icon_type: IconSetType,
        /// Thresholds for each icon (one fewer than the number of icons).
        thresholds: Vec<CfValueObject>,
        /// If true, the icon order is reversed.
        reversed: bool,
        /// If true, only icons are shown (no cell value).
        show_value: bool,
    },
    /// Top N / Bottom N (cfRule type="top10").
    Top10 {
        /// The rank value (e.g. 10 for "top 10").
        rank: u32,
        /// If true, rank is a percentage.
        percent: bool,
        /// If true, this is a "bottom" rule instead of "top".
        bottom: bool,
    },
    /// Above/below average (cfRule type="aboveAverage").
    AboveAverage {
        /// False means "below average".
        above_average: bool,
        /// If true, includes values equal to the average.
        equal_average: bool,
        /// Standard deviation level (0 = no std dev filter).
        std_dev: u32,
    },
    /// Text-based rules (containsText, notContainsText, beginsWith, endsWith).
    Text {
        /// The text operator.
        operator: CfTextOperator,
        /// The text to match.
        text: String,
        /// The formula used internally by Excel.
        formula: Option<String>,
    },
    /// Time period rule (cfRule type="timePeriod").
    TimePeriod {
        /// The time period.
        period: TimePeriodType,
        /// The formula used internally by Excel.
        formula: Option<String>,
    },
    /// Contains blanks (cfRule type="containsBlanks").
    ContainsBlanks {
        /// The formula used internally by Excel.
        formula: Option<String>,
    },
    /// Does not contain blanks (cfRule type="notContainsBlanks").
    NotContainsBlanks {
        /// The formula used internally by Excel.
        formula: Option<String>,
    },
    /// Contains errors (cfRule type="containsErrors").
    ContainsErrors {
        /// The formula used internally by Excel.
        formula: Option<String>,
    },
    /// Does not contain errors (cfRule type="notContainsErrors").
    NotContainsErrors {
        /// The formula used internally by Excel.
        formula: Option<String>,
    },
    /// Duplicate values (cfRule type="duplicateValues").
    DuplicateValues,
    /// Unique values (cfRule type="uniqueValues").
    UniqueValues,
    /// An unrecognized or unsupported rule type.
    Unknown {
        /// The raw type attribute value.
        raw_type: String,
    },
}

/// Comparison operators for cell value rules.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CfOperator {
    /// Equal to.
    Equal,
    /// Not equal to.
    NotEqual,
    /// Greater than.
    GreaterThan,
    /// Greater than or equal to.
    GreaterThanOrEqual,
    /// Less than.
    LessThan,
    /// Less than or equal to.
    LessThanOrEqual,
    /// Between (inclusive).
    Between,
    /// Not between.
    NotBetween,
}

impl CfOperator {
    pub(crate) fn from_str(s: &str) -> Option<Self> {
        match s {
            "equal" => Some(Self::Equal),
            "notEqual" => Some(Self::NotEqual),
            "greaterThan" => Some(Self::GreaterThan),
            "greaterThanOrEqual" => Some(Self::GreaterThanOrEqual),
            "lessThan" => Some(Self::LessThan),
            "lessThanOrEqual" => Some(Self::LessThanOrEqual),
            "between" => Some(Self::Between),
            "notBetween" => Some(Self::NotBetween),
            _ => None,
        }
    }
}

/// Text comparison operators.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CfTextOperator {
    /// Cell text contains the value.
    Contains,
    /// Cell text does not contain the value.
    NotContains,
    /// Cell text begins with the value.
    BeginsWith,
    /// Cell text ends with the value.
    EndsWith,
}

/// Value type for color scale, data bar, and icon set thresholds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CfValueObjectType {
    /// Minimum value in the range.
    Min,
    /// Maximum value in the range.
    Max,
    /// A literal number.
    Num,
    /// A percentage (0-100).
    Percent,
    /// A percentile.
    Percentile,
    /// A formula.
    Formula,
    /// Automatic (context-dependent).
    Auto,
}

impl CfValueObjectType {
    pub(crate) fn from_str(s: &str) -> Self {
        match s {
            "min" => Self::Min,
            "max" => Self::Max,
            "num" => Self::Num,
            "percent" => Self::Percent,
            "percentile" => Self::Percentile,
            "formula" => Self::Formula,
            "autoMin" | "autoMax" => Self::Auto,
            _ => Self::Auto,
        }
    }
}

/// A conditional format value object (`<cfvo>` element).
#[derive(Debug, Clone, PartialEq)]
pub struct CfValueObject {
    /// The type of the threshold value.
    pub value_type: CfValueObjectType,
    /// The threshold value (may be a number string or formula).
    pub value: Option<String>,
    /// If true (default), the comparison is "greater than or equal to".
    /// If false, the comparison is strictly "greater than".
    pub gte: bool,
}

impl Default for CfValueObject {
    fn default() -> Self {
        Self {
            value_type: CfValueObjectType::Auto,
            value: None,
            gte: true,
        }
    }
}

/// Time period types for date-based conditional formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimePeriodType {
    /// Yesterday.
    Yesterday,
    /// Today.
    Today,
    /// Tomorrow.
    Tomorrow,
    /// Last 7 days.
    Last7Days,
    /// Last week.
    LastWeek,
    /// This week.
    ThisWeek,
    /// Next week.
    NextWeek,
    /// Last month.
    LastMonth,
    /// This month.
    ThisMonth,
    /// Next month.
    NextMonth,
}

impl TimePeriodType {
    pub(crate) fn from_str(s: &str) -> Option<Self> {
        match s {
            "yesterday" => Some(Self::Yesterday),
            "today" => Some(Self::Today),
            "tomorrow" => Some(Self::Tomorrow),
            "last7Days" => Some(Self::Last7Days),
            "lastWeek" => Some(Self::LastWeek),
            "thisWeek" => Some(Self::ThisWeek),
            "nextWeek" => Some(Self::NextWeek),
            "lastMonth" => Some(Self::LastMonth),
            "thisMonth" => Some(Self::ThisMonth),
            "nextMonth" => Some(Self::NextMonth),
            _ => None,
        }
    }
}

/// Icon set types for icon-based conditional formatting.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IconSetType {
    // 3-icon sets
    /// 3 arrows (colored).
    ThreeArrows,
    /// 3 arrows (gray).
    ThreeArrowsGray,
    /// 3 flags.
    ThreeFlags,
    /// 3 traffic lights (no rim).
    ThreeTrafficLights,
    /// 3 traffic lights (with rim).
    ThreeTrafficLightsWithRim,
    /// 3 signs.
    ThreeSigns,
    /// 3 symbols (circled).
    ThreeSymbolsCircled,
    /// 3 symbols (uncircled).
    ThreeSymbols,
    /// 3 stars.
    ThreeStars,
    /// 3 triangles.
    ThreeTriangles,

    // 4-icon sets
    /// 4 arrows (colored).
    FourArrows,
    /// 4 arrows (gray).
    FourArrowsGray,
    /// 4 red-to-black.
    FourRedToBlack,
    /// 4 traffic lights.
    FourTrafficLights,
    /// 4 ratings (histograms).
    FourRatings,

    // 5-icon sets
    /// 5 arrows (colored).
    FiveArrows,
    /// 5 arrows (gray).
    FiveArrowsGray,
    /// 5 ratings (histograms).
    FiveRatings,
    /// 5 quarters.
    FiveQuarters,
    /// 5 boxes.
    FiveBoxes,
}

impl IconSetType {
    pub(crate) fn from_str(s: &str) -> Option<Self> {
        match s {
            "3Arrows" => Some(Self::ThreeArrows),
            "3ArrowsGray" => Some(Self::ThreeArrowsGray),
            "3Flags" => Some(Self::ThreeFlags),
            "3TrafficLights" => Some(Self::ThreeTrafficLights),
            "3TrafficLights2" => Some(Self::ThreeTrafficLightsWithRim),
            "3Signs" => Some(Self::ThreeSigns),
            "3Symbols" => Some(Self::ThreeSymbolsCircled),
            "3Symbols2" => Some(Self::ThreeSymbols),
            "3Stars" => Some(Self::ThreeStars),
            "3Triangles" => Some(Self::ThreeTriangles),
            "4Arrows" => Some(Self::FourArrows),
            "4ArrowsGray" => Some(Self::FourArrowsGray),
            "4RedToBlack" => Some(Self::FourRedToBlack),
            "4TrafficLights" => Some(Self::FourTrafficLights),
            "4Rating" => Some(Self::FourRatings),
            "5Arrows" => Some(Self::FiveArrows),
            "5ArrowsGray" => Some(Self::FiveArrowsGray),
            "5Rating" => Some(Self::FiveRatings),
            "5Quarters" => Some(Self::FiveQuarters),
            "5Boxes" => Some(Self::FiveBoxes),
            _ => None,
        }
    }
}

//! Conditional formatting data structures and parsing

use crate::formats::Color;
use crate::Dimensions;
use std::fmt;

/// Conditional formatting rule type
#[derive(Debug, Clone, PartialEq)]
pub enum ConditionalFormatType {
    /// Cell value comparison
    CellIs {
        /// Comparison operator
        operator: ComparisonOperator,
    },
    /// Expression/formula-based rule
    Expression,
    /// Top/bottom N values or percentiles
    Top10 {
        /// Bottom instead of top
        bottom: bool,
        /// Use percent instead of rank
        percent: bool,
        /// Number of items or percentage
        rank: u32,
    },
    /// Duplicate values
    DuplicateValues,
    /// Unique values
    UniqueValues,
    /// Contains text
    ContainsText {
        /// Text to search for
        text: String,
    },
    /// Contains text (not case sensitive)
    NotContainsText {
        /// Text to search for
        text: String,
    },
    /// Begins with text
    BeginsWith {
        /// Text to match at start
        text: String,
    },
    /// Ends with text
    EndsWith {
        /// Text to match at end
        text: String,
    },
    /// Is blank
    ContainsBlanks,
    /// Is not blank
    NotContainsBlanks,
    /// Contains errors
    ContainsErrors,
    /// Does not contain errors
    NotContainsErrors,
    /// Date occurring
    TimePeriod {
        /// Time period type
        period: TimePeriod,
    },
    /// Above or below average
    AboveAverage {
        /// Below instead of above
        below: bool,
        /// Include equal to average
        equal_average: bool,
        /// Standard deviations
        std_dev: Option<u32>,
    },
    /// Data bar
    DataBar(DataBar),
    /// Color scale
    ColorScale(ColorScale),
    /// Icon set
    IconSet(IconSet),
}

/// Comparison operators for CellIs rules
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComparisonOperator {
    /// Less than
    LessThan,
    /// Less than or equal
    LessThanOrEqual,
    /// Equal
    Equal,
    /// Not equal
    NotEqual,
    /// Greater than or equal
    GreaterThanOrEqual,
    /// Greater than
    GreaterThan,
    /// Between (inclusive)
    Between,
    /// Not between (exclusive)
    NotBetween,
    /// Contains text
    ContainsText,
    /// Does not contain text
    NotContains,
}

/// Time period for date-based rules
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimePeriod {
    /// Today
    Today,
    /// Yesterday
    Yesterday,
    /// Tomorrow
    Tomorrow,
    /// Last 7 days
    Last7Days,
    /// This week
    ThisWeek,
    /// Last week
    LastWeek,
    /// Next week
    NextWeek,
    /// This month
    ThisMonth,
    /// Last month
    LastMonth,
    /// Next month
    NextMonth,
    /// This quarter
    ThisQuarter,
    /// Last quarter
    LastQuarter,
    /// Next quarter
    NextQuarter,
    /// This year
    ThisYear,
    /// Last year
    LastYear,
    /// Next year
    NextYear,
    /// Year to date
    YearToDate,
    /// All dates in January
    AllDatesInJanuary,
    /// All dates in February
    AllDatesInFebruary,
    /// All dates in March
    AllDatesInMarch,
    /// All dates in April
    AllDatesInApril,
    /// All dates in May
    AllDatesInMay,
    /// All dates in June
    AllDatesInJune,
    /// All dates in July
    AllDatesInJuly,
    /// All dates in August
    AllDatesInAugust,
    /// All dates in September
    AllDatesInSeptember,
    /// All dates in October
    AllDatesInOctober,
    /// All dates in November
    AllDatesInNovember,
    /// All dates in December
    AllDatesInDecember,
    /// All dates in Q1
    AllDatesInQ1,
    /// All dates in Q2
    AllDatesInQ2,
    /// All dates in Q3
    AllDatesInQ3,
    /// All dates in Q4
    AllDatesInQ4,
}

/// Conditional format value object (threshold)
#[derive(Debug, Clone, PartialEq)]
pub struct ConditionalFormatValue {
    /// Value type
    pub value_type: CfvoType,
    /// The actual value (if applicable)
    pub value: Option<String>,
    /// Greater than or equal (for percentile)
    pub gte: bool,
}

/// Conditional format value object type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CfvoType {
    /// Minimum value in the range
    Min,
    /// Maximum value in the range  
    Max,
    /// Specific number
    Number,
    /// Percentage
    Percent,
    /// Percentile
    Percentile,
    /// Formula
    Formula,
    /// Automatic minimum
    AutoMin,
    /// Automatic maximum
    AutoMax,
}

/// Data bar configuration
#[derive(Debug, Clone, PartialEq)]
pub struct DataBar {
    /// Minimum threshold
    pub min_cfvo: ConditionalFormatValue,
    /// Maximum threshold
    pub max_cfvo: ConditionalFormatValue,
    /// Bar color
    pub color: Color,
    /// Negative bar color (2010+ extension)
    pub negative_color: Option<Color>,
    /// Show values in cells
    pub show_value: bool,
    /// Minimum bar length (percentage)
    pub min_length: u32,
    /// Maximum bar length (percentage)
    pub max_length: u32,
    /// Bar direction
    pub direction: Option<BarDirection>,
    /// Only show bar (hide value)
    pub bar_only: bool,
    /// Border color
    pub border_color: Option<Color>,
    /// Negative bar border color
    pub negative_border_color: Option<Color>,
    /// Fill type (solid or gradient)
    pub gradient: bool,
    /// Axis position
    pub axis_position: Option<AxisPosition>,
    /// Axis color
    pub axis_color: Option<Color>,
}

/// Bar direction for data bars
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarDirection {
    /// Left to right (default)
    LeftToRight,
    /// Right to left
    RightToLeft,
}

/// Axis position for data bars
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisPosition {
    /// Automatic positioning
    Automatic,
    /// At cell midpoint
    Midpoint,
    /// No axis
    None,
}

/// Color scale configuration
#[derive(Debug, Clone, PartialEq)]
pub struct ColorScale {
    /// Color scale stops (2 or 3)
    pub cfvos: Vec<ConditionalFormatValue>,
    /// Colors corresponding to each stop
    pub colors: Vec<Color>,
}

/// Icon set configuration
#[derive(Debug, Clone, PartialEq)]
pub struct IconSet {
    /// Icon set type
    pub icon_set: IconSetType,
    /// Thresholds (typically 2-4 values)
    pub cfvos: Vec<ConditionalFormatValue>,
    /// Show values in cells
    pub show_value: bool,
    /// Reverse icon order
    pub reverse: bool,
    /// Custom icons (2010+ extension)
    pub custom_icons: Vec<Option<(IconSetType, u32)>>,
    /// Percent values
    pub percent: bool,
}

/// Built-in icon set types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum IconSetType {
    /// 3 arrows (colored)
    Arrows3,
    /// 3 arrows (gray)
    Arrows3Gray,
    /// 4 arrows (colored)
    Arrows4,
    /// 4 arrows (gray)
    Arrows4Gray,
    /// 5 arrows (colored)
    Arrows5,
    /// 5 arrows (gray)
    Arrows5Gray,
    /// 3 flags
    Flags3,
    /// 3 traffic lights (unrimmed)
    TrafficLights3,
    /// 3 traffic lights (rimmed)
    TrafficLights3Rimmed,
    /// 4 traffic lights
    TrafficLights4,
    /// 3 signs
    Signs3,
    /// 3 symbols (circled)
    Symbols3,
    /// 3 symbols (uncircled)
    Symbols3Uncircled,
    /// 4 rating
    Rating4,
    /// 5 rating
    Rating5,
    /// 5 quarters
    Quarters5,
    /// 3 stars
    Stars3,
    /// 3 triangles
    Triangles3,
    /// 5 boxes
    Boxes5,
    /// 3 symbols 2
    Symbols3_2,
    /// 4 red to black
    RedToBlack4,
    /// 4 rating bars
    RatingBars4,
    /// 5 rating bars  
    RatingBars5,
    /// 3 colored arrows
    ColoredArrows3,
    /// 4 colored arrows
    ColoredArrows4,
    /// 5 colored arrows
    ColoredArrows5,
    /// 3 white arrows
    WhiteArrows3,
    /// 4 white arrows
    WhiteArrows4,
    /// 5 white arrows
    WhiteArrows5,
}

/// A single conditional formatting rule
#[derive(Debug, Clone)]
pub struct ConditionalFormatRule {
    /// Rule type and configuration
    pub rule_type: ConditionalFormatType,
    /// Priority (lower number = higher priority)
    pub priority: i32,
    /// Stop if this rule matches
    pub stop_if_true: bool,
    /// Differential format ID (reference to styles.xml)
    pub dxf_id: Option<u32>,
    /// Formula(s) for the rule
    pub formulas: Vec<String>,
    /// Pivot table rule
    pub pivot: bool,
    /// Text for text-based rules
    pub text: Option<String>,
    /// Operator for comparison rules
    pub operator: Option<String>,
    /// Bottom N (as opposed to top N)
    pub bottom: Option<bool>,
    /// Percent flag
    pub percent: Option<bool>,
    /// Rank value
    pub rank: Option<i32>,
    /// Above average flag
    pub above_average: Option<bool>,
    /// Equal average flag
    pub equal_average: Option<bool>,
    /// Standard deviation value
    pub std_dev: Option<i32>,
}

/// Rule scope for conditional formatting
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RuleScope {
    /// Worksheet scope
    Worksheet,
    /// Table scope
    Table,
    /// Pivot table scope
    PivotTable,
    /// Selection scope
    Selection(String),
}

/// Conditional formatting for a range
#[derive(Debug, Clone)]
pub struct ConditionalFormatting {
    /// Cell ranges this formatting applies to (space-separated in XML)
    pub ranges: Vec<Dimensions>,
    /// Rules in priority order
    pub rules: Vec<ConditionalFormatRule>,
    /// Scope of the rule
    pub scope: Option<RuleScope>,
    /// Table name (for table/pivot table scope)
    pub table: Option<String>,
}

/// Differential formatting record
#[derive(Debug, Clone, Default)]
pub struct DifferentialFormat {
    /// Font changes
    pub font: Option<DifferentialFont>,
    /// Fill changes
    pub fill: Option<DifferentialFill>,
    /// Border changes
    pub border: Option<DifferentialBorder>,
    /// Number format
    pub number_format: Option<DifferentialNumberFormat>,
    /// Alignment changes
    pub alignment: Option<DifferentialAlignment>,
    /// Protection changes
    pub protection: Option<DifferentialProtection>,
}

/// Differential font formatting
#[derive(Debug, Clone, Default)]
pub struct DifferentialFont {
    /// Font name
    pub name: Option<String>,
    /// Font size
    pub size: Option<f64>,
    /// Bold
    pub bold: Option<bool>,
    /// Italic
    pub italic: Option<bool>,
    /// Underline
    pub underline: Option<bool>,
    /// Strike through
    pub strike: Option<bool>,
    /// Font color
    pub color: Option<Color>,
    /// Font scheme
    pub scheme: Option<String>,
    /// Font family
    pub family: Option<i32>,
    /// Character set
    pub charset: Option<i32>,
}

/// Differential fill formatting
#[derive(Debug, Clone)]
pub struct DifferentialFill {
    /// Pattern fill
    pub pattern_fill: PatternFill,
}

/// Pattern fill for differential formatting
#[derive(Debug, Clone)]
pub struct PatternFill {
    /// Pattern type
    pub pattern_type: Option<String>,
    /// Foreground color
    pub fg_color: Option<Color>,
    /// Background color
    pub bg_color: Option<Color>,
}

/// Differential border formatting
#[derive(Debug, Clone, Default)]
pub struct DifferentialBorder {
    /// Left border
    pub left: Option<DifferentialBorderSide>,
    /// Right border
    pub right: Option<DifferentialBorderSide>,
    /// Top border
    pub top: Option<DifferentialBorderSide>,
    /// Bottom border
    pub bottom: Option<DifferentialBorderSide>,
    /// Diagonal border
    pub diagonal: Option<DifferentialBorderSide>,
    /// Diagonal up
    pub diagonal_up: Option<bool>,
    /// Diagonal down
    pub diagonal_down: Option<bool>,
    /// Outline
    pub outline: Option<bool>,
}

/// Border side for differential formatting
#[derive(Debug, Clone)]
pub struct DifferentialBorderSide {
    /// Border style
    pub style: Option<String>,
    /// Border color
    pub color: Option<Color>,
}

/// Differential number format
#[derive(Debug, Clone)]
pub struct DifferentialNumberFormat {
    /// Format code
    pub format_code: String,
    /// Format ID
    pub num_fmt_id: Option<u32>,
}

/// Differential alignment
#[derive(Debug, Clone, Default)]
pub struct DifferentialAlignment {
    /// Horizontal alignment
    pub horizontal: Option<String>,
    /// Vertical alignment
    pub vertical: Option<String>,
    /// Text rotation
    pub text_rotation: Option<i32>,
    /// Wrap text
    pub wrap_text: Option<bool>,
    /// Shrink to fit
    pub shrink_to_fit: Option<bool>,
    /// Indent
    pub indent: Option<u32>,
    /// Relative indent
    pub relative_indent: Option<i32>,
    /// Justify last line
    pub justify_last_line: Option<bool>,
    /// Reading order
    pub reading_order: Option<u32>,
}

/// Differential protection
#[derive(Debug, Clone, Default)]
pub struct DifferentialProtection {
    /// Locked
    pub locked: Option<bool>,
    /// Hidden
    pub hidden: Option<bool>,
}

impl fmt::Display for ComparisonOperator {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        use ComparisonOperator::*;
        match self {
            LessThan => write!(f, "lessThan"),
            LessThanOrEqual => write!(f, "lessThanOrEqual"),
            Equal => write!(f, "equal"),
            NotEqual => write!(f, "notEqual"),
            GreaterThanOrEqual => write!(f, "greaterThanOrEqual"),
            GreaterThan => write!(f, "greaterThan"),
            Between => write!(f, "between"),
            NotBetween => write!(f, "notBetween"),
            ContainsText => write!(f, "containsText"),
            NotContains => write!(f, "notContains"),
        }
    }
}

impl fmt::Display for TimePeriod {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        use TimePeriod::*;
        match self {
            Today => write!(f, "today"),
            Yesterday => write!(f, "yesterday"),
            Tomorrow => write!(f, "tomorrow"),
            Last7Days => write!(f, "last7Days"),
            ThisWeek => write!(f, "thisWeek"),
            LastWeek => write!(f, "lastWeek"),
            NextWeek => write!(f, "nextWeek"),
            ThisMonth => write!(f, "thisMonth"),
            LastMonth => write!(f, "lastMonth"),
            NextMonth => write!(f, "nextMonth"),
            ThisQuarter => write!(f, "thisQuarter"),
            LastQuarter => write!(f, "lastQuarter"),
            NextQuarter => write!(f, "nextQuarter"),
            ThisYear => write!(f, "thisYear"),
            LastYear => write!(f, "lastYear"),
            NextYear => write!(f, "nextYear"),
            YearToDate => write!(f, "yearToDate"),
            AllDatesInJanuary => write!(f, "allDatesInPeriodJanuary"),
            AllDatesInFebruary => write!(f, "allDatesInPeriodFebruary"),
            AllDatesInMarch => write!(f, "allDatesInPeriodMarch"),
            AllDatesInApril => write!(f, "allDatesInPeriodApril"),
            AllDatesInMay => write!(f, "allDatesInPeriodMay"),
            AllDatesInJune => write!(f, "allDatesInPeriodJune"),
            AllDatesInJuly => write!(f, "allDatesInPeriodJuly"),
            AllDatesInAugust => write!(f, "allDatesInPeriodAugust"),
            AllDatesInSeptember => write!(f, "allDatesInPeriodSeptember"),
            AllDatesInOctober => write!(f, "allDatesInPeriodOctober"),
            AllDatesInNovember => write!(f, "allDatesInPeriodNovember"),
            AllDatesInDecember => write!(f, "allDatesInPeriodDecember"),
            AllDatesInQ1 => write!(f, "allDatesInPeriodQuarter1"),
            AllDatesInQ2 => write!(f, "allDatesInPeriodQuarter2"),
            AllDatesInQ3 => write!(f, "allDatesInPeriodQuarter3"),
            AllDatesInQ4 => write!(f, "allDatesInPeriodQuarter4"),
        }
    }
}

impl fmt::Display for CfvoType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        use CfvoType::*;
        match self {
            Min => write!(f, "min"),
            Max => write!(f, "max"),
            Number => write!(f, "num"),
            Percent => write!(f, "percent"),
            Percentile => write!(f, "percentile"),
            Formula => write!(f, "formula"),
            AutoMin => write!(f, "autoMin"),
            AutoMax => write!(f, "autoMax"),
        }
    }
}

impl fmt::Display for IconSetType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        use IconSetType::*;
        match self {
            Arrows3 => write!(f, "3Arrows"),
            Arrows3Gray => write!(f, "3ArrowsGray"),
            Arrows4 => write!(f, "4Arrows"),
            Arrows4Gray => write!(f, "4ArrowsGray"),
            Arrows5 => write!(f, "5Arrows"),
            Arrows5Gray => write!(f, "5ArrowsGray"),
            Flags3 => write!(f, "3Flags"),
            TrafficLights3 => write!(f, "3TrafficLights1"),
            TrafficLights3Rimmed => write!(f, "3TrafficLights2"),
            TrafficLights4 => write!(f, "4TrafficLights"),
            Signs3 => write!(f, "3Signs"),
            Symbols3 => write!(f, "3Symbols"),
            Symbols3Uncircled => write!(f, "3Symbols2"),
            Rating4 => write!(f, "4Rating"),
            Rating5 => write!(f, "5Rating"),
            Quarters5 => write!(f, "5Quarters"),
            Stars3 => write!(f, "3Stars"),
            Triangles3 => write!(f, "3Triangles"),
            Boxes5 => write!(f, "5Boxes"),
            Symbols3_2 => write!(f, "3Symbols2"),
            RedToBlack4 => write!(f, "4RedToBlack"),
            RatingBars4 => write!(f, "4RatingBars"),
            RatingBars5 => write!(f, "5RatingBars"),
            ColoredArrows3 => write!(f, "3ColoredArrows"),
            ColoredArrows4 => write!(f, "4ColoredArrows"),
            ColoredArrows5 => write!(f, "5ColoredArrows"),
            WhiteArrows3 => write!(f, "3WhiteArrows"),
            WhiteArrows4 => write!(f, "4WhiteArrows"),
            WhiteArrows5 => write!(f, "5WhiteArrows"),
        }
    }
}

impl fmt::Display for BarDirection {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            BarDirection::LeftToRight => write!(f, "leftToRight"),
            BarDirection::RightToLeft => write!(f, "rightToLeft"),
        }
    }
}

impl fmt::Display for AxisPosition {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            AxisPosition::Automatic => write!(f, "automatic"),
            AxisPosition::Midpoint => write!(f, "midpoint"),
            AxisPosition::None => write!(f, "none"),
        }
    }
}
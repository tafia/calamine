// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Chart types for charts read from XLSX files.
//!
//! Charts are stored in an XLSX package as DrawingML "chartSpace" parts
//! (`xl/charts/chartN.xml`) referenced from a worksheet (or chartsheet)
//! drawing. [`crate::Xlsx::worksheet_charts`] walks those relationships and
//! returns one [`Chart`] per embedded chart, including its plot groups,
//! series (with cached values), axes, title, legend, 3D view settings and
//! shape/line formatting.
//!
//! All classic ECMA-376 chart families are supported (bar, column, line,
//! pie, doughnut, pie-of-pie, area, scatter, radar, stock, bubble and
//! surface, including their stacked and 3D variants), along with data
//! labels (including per-point overrides), trendlines, error bars,
//! up/down bars, drop/high-low/series lines, data tables and the common
//! axis and series options. The Excel 2016+ "chart-ex" families
//! (`xl/charts/chartExN.xml`: funnel, treemap, sunburst, histogram,
//! pareto, box & whisker, waterfall and filled map) are read with their
//! type, literal series data (including hierarchical category levels),
//! layout options ([`ChartExLayout`]), title and legend.
//!
//! Not read: pivot chart sources, manual plot-area layouts and surface
//! band formats.

use crate::datatype::Data;
use crate::style::{Color, Font, RichText};

/// The type of a chart or of one of its plot groups.
///
/// The variants mirror the Excel chart families, including the 3D families.
/// Combo charts contain multiple [`ChartGroup`]s with different types.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[non_exhaustive]
pub enum ChartType {
    /// A 2D area chart.
    Area,
    /// A stacked 2D area chart.
    AreaStacked,
    /// A percent-stacked 2D area chart.
    AreaPercentStacked,
    /// A horizontal bar chart.
    Bar,
    /// A stacked horizontal bar chart.
    BarStacked,
    /// A percent-stacked horizontal bar chart.
    BarPercentStacked,
    /// A vertical column chart.
    Column,
    /// A stacked vertical column chart.
    ColumnStacked,
    /// A percent-stacked vertical column chart.
    ColumnPercentStacked,
    /// A doughnut chart.
    Doughnut,
    /// A line chart.
    Line,
    /// A stacked line chart.
    LineStacked,
    /// A percent-stacked line chart.
    LinePercentStacked,
    /// A pie chart.
    Pie,
    /// A pie-of-pie chart.
    PieOfPie,
    /// A bar-of-pie chart.
    BarOfPie,
    /// A radar chart.
    Radar,
    /// A radar chart with markers.
    RadarWithMarkers,
    /// A filled radar chart.
    RadarFilled,
    /// A scatter chart with markers only.
    Scatter,
    /// A scatter chart with straight connecting lines and no markers.
    ScatterStraight,
    /// A scatter chart with straight connecting lines and markers.
    ScatterStraightWithMarkers,
    /// A scatter chart with smoothed connecting lines and no markers.
    ScatterSmooth,
    /// A scatter chart with smoothed connecting lines and markers.
    ScatterSmoothWithMarkers,
    /// A stock (high-low-close) chart.
    Stock,
    /// A bubble chart.
    Bubble,
    /// A 3D area chart.
    Area3D,
    /// A stacked 3D area chart.
    Area3DStacked,
    /// A percent-stacked 3D area chart.
    Area3DPercentStacked,
    /// A 3D horizontal bar chart.
    Bar3D,
    /// A stacked 3D horizontal bar chart.
    Bar3DStacked,
    /// A percent-stacked 3D horizontal bar chart.
    Bar3DPercentStacked,
    /// A true 3D horizontal bar chart with series along the depth axis
    /// (`c:grouping` `standard`).
    Bar3DStandard,
    /// A 3D vertical column chart.
    Column3D,
    /// A stacked 3D vertical column chart.
    Column3DStacked,
    /// A percent-stacked 3D vertical column chart.
    Column3DPercentStacked,
    /// A true 3D vertical column chart with series along the depth axis
    /// (`c:grouping` `standard`).
    Column3DStandard,
    /// A 3D line chart.
    Line3D,
    /// A 3D pie chart.
    Pie3D,
    /// A 3D surface chart.
    Surface3D,
    /// A wireframe 3D surface chart.
    Surface3DWireframe,
    /// A contour chart (top view of a surface chart).
    Contour,
    /// A wireframe contour chart.
    ContourWireframe,
    /// A funnel chart (Excel 2016+ "chart-ex" chart).
    Funnel,
    /// A treemap chart (Excel 2016+ "chart-ex" chart).
    Treemap,
    /// A sunburst chart (Excel 2016+ "chart-ex" chart).
    Sunburst,
    /// A histogram chart (Excel 2016+ "chart-ex" chart).
    Histogram,
    /// A pareto chart (Excel 2016+ "chart-ex" chart).
    Pareto,
    /// A box & whisker chart (Excel 2016+ "chart-ex" chart).
    BoxWhisker,
    /// A waterfall chart (Excel 2016+ "chart-ex" chart).
    Waterfall,
    /// A filled map chart (Excel 2016+ "chart-ex" chart).
    RegionMap,
    /// An unrecognized chart type.
    #[default]
    Unknown,
}

impl ChartType {
    /// Returns `true` for the 3D chart families.
    pub fn is_3d(self) -> bool {
        matches!(
            self,
            ChartType::Area3D
                | ChartType::Area3DStacked
                | ChartType::Area3DPercentStacked
                | ChartType::Bar3D
                | ChartType::Bar3DStacked
                | ChartType::Bar3DPercentStacked
                | ChartType::Bar3DStandard
                | ChartType::Column3D
                | ChartType::Column3DStacked
                | ChartType::Column3DPercentStacked
                | ChartType::Column3DStandard
                | ChartType::Line3D
                | ChartType::Pie3D
                | ChartType::Surface3D
                | ChartType::Surface3DWireframe
                | ChartType::Contour
                | ChartType::ContourWireframe
        )
    }

    /// Returns `true` for the Excel 2016+ "chart-ex" chart families
    /// (funnel, treemap, sunburst, histogram, pareto, box & whisker,
    /// waterfall and filled map).
    pub fn is_chart_ex(self) -> bool {
        matches!(
            self,
            ChartType::Funnel
                | ChartType::Treemap
                | ChartType::Sunburst
                | ChartType::Histogram
                | ChartType::Pareto
                | ChartType::BoxWhisker
                | ChartType::Waterfall
                | ChartType::RegionMap
        )
    }
}

/// A chart read from an XLSX file.
///
/// Returned by [`crate::Xlsx::worksheet_charts`]. A chart contains one or
/// more plot [`ChartGroup`]s (more than one for combo charts), the axes
/// declared in the plot area, and the chart-level title, legend, 3D view and
/// formatting.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Chart {
    /// The drawing object name, e.g. `Chart 1`.
    pub name: Option<String>,
    /// Where the chart is anchored on the worksheet.
    pub position: Option<ChartPosition>,
    /// The chart title.
    pub title: Option<ChartTitle>,
    /// The chart legend, if shown.
    pub legend: Option<ChartLegend>,
    /// The 3D view settings (`c:view3D`), present for 3D charts.
    pub view_3d: Option<ChartView3d>,
    /// The axes declared in the plot area, in document order.
    pub axes: Vec<ChartAxis>,
    /// The plot groups. Each group has a chart type and its own series.
    pub groups: Vec<ChartGroup>,
    /// The chart style number (`c:style`), 1-48.
    pub style: Option<u32>,
    /// Chart area (chart space) shape formatting.
    pub format: Option<ChartFormat>,
    /// Plot area shape formatting.
    pub plot_area_format: Option<ChartFormat>,
    /// How empty cells are plotted (`c:dispBlanksAs`).
    pub display_blanks_as: Option<ChartDisplayBlanksAs>,
    /// Whether the chart area has rounded corners (`c:roundedCorners`).
    pub rounded_corners: Option<bool>,
    /// The data table shown under the chart (`c:dTable`), if any.
    pub data_table: Option<ChartDataTable>,
    /// Whether the automatic title was deleted (`c:autoTitleDeleted`).
    pub auto_title_deleted: Option<bool>,
    /// Whether only visible cells are plotted (`c:plotVisOnly`).
    pub plot_visible_only: Option<bool>,
    /// Whether data labels over the value axis maximum are shown
    /// (`c:showDLblsOverMax`).
    pub show_data_labels_over_max: Option<bool>,
    /// Whether the chart uses the 1904 date system (`c:date1904`).
    pub date_1904: Option<bool>,
}

impl Chart {
    /// The primary chart type: the type of the first plot group.
    pub fn chart_type(&self) -> ChartType {
        self.groups
            .first()
            .map(|g| g.chart_type)
            .unwrap_or(ChartType::Unknown)
    }

    /// Iterate over all series across all plot groups.
    pub fn series(&self) -> impl Iterator<Item = &ChartSeries> {
        self.groups.iter().flat_map(|g| g.series.iter())
    }

    /// The X (category) axis.
    ///
    /// This is the first category or date axis, or for scatter/bubble charts
    /// (which use two value axes) the first value axis.
    pub fn x_axis(&self) -> Option<&ChartAxis> {
        self.axes
            .iter()
            .find(|a| {
                matches!(
                    a.axis_type,
                    ChartAxisType::Category | ChartAxisType::Date
                )
            })
            .or_else(|| {
                self.axes
                    .iter()
                    .find(|a| a.axis_type == ChartAxisType::Value)
            })
    }

    /// The Y (value) axis.
    ///
    /// This is the first value axis, except for scatter/bubble charts where
    /// it is the second value axis.
    pub fn y_axis(&self) -> Option<&ChartAxis> {
        let mut values = self
            .axes
            .iter()
            .filter(|a| a.axis_type == ChartAxisType::Value);
        let first = values.next();
        let has_cat = self.axes.iter().any(|a| {
            matches!(
                a.axis_type,
                ChartAxisType::Category | ChartAxisType::Date
            )
        });
        if has_cat {
            first
        } else {
            values.next().or(first)
        }
    }

    /// The series (depth) axis of a 3D chart, if present.
    pub fn series_axis(&self) -> Option<&ChartAxis> {
        self.axes
            .iter()
            .find(|a| a.axis_type == ChartAxisType::Series)
    }
}

/// One plot group inside a chart: a chart type plus the series plotted with
/// that type and the group-level layout options.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartGroup {
    /// The chart type of this group.
    pub chart_type: ChartType,
    /// The series in this group, in document order.
    pub series: Vec<ChartSeries>,
    /// Gap between bar/column clusters, as a percentage (`c:gapWidth`).
    pub gap_width: Option<u32>,
    /// Depth gap for 3D bar/column charts, as a percentage (`c:gapDepth`).
    pub gap_depth: Option<u32>,
    /// Overlap between bars/columns, -100 to 100 (`c:overlap`).
    pub overlap: Option<i32>,
    /// Doughnut hole size, as a percentage (`c:holeSize`).
    pub hole_size: Option<u32>,
    /// Rotation of the first pie/doughnut slice, in degrees
    /// (`c:firstSliceAng`).
    pub first_slice_angle: Option<u32>,
    /// Bubble size scale, as a percentage (`c:bubbleScale`).
    pub bubble_scale: Option<u32>,
    /// Whether data points vary in color (`c:varyColors`).
    pub vary_colors: Option<bool>,
    /// The ids of the axes this group plots against (`c:axId`).
    pub axis_ids: Vec<u32>,
    /// Whether markers are shown for line charts (`c:marker`).
    pub show_marker: Option<bool>,
    /// Default data labels for the group (`c:dLbls`).
    pub data_labels: Option<ChartDataLabels>,
    /// The 3D shape of bars/columns in a 3D chart (`c:shape`).
    pub shape: Option<ChartBar3dShape>,
    /// How bubble sizes map to bubble data (`c:sizeRepresents`).
    pub size_represents: Option<ChartSizeRepresents>,
    /// Whether negative-value bubbles are shown (`c:showNegBubbles`).
    pub show_negative_bubbles: Option<bool>,
    /// How the second plot of a pie-of-pie/bar-of-pie chart is split
    /// (`c:splitType`).
    pub split_type: Option<ChartOfPieSplitType>,
    /// The split threshold used with [`split_type`](Self::split_type)
    /// (`c:splitPos`).
    pub split_position: Option<f64>,
    /// The zero-based point indices assigned to the second plot when
    /// [`split_type`](Self::split_type) is
    /// [`ChartOfPieSplitType::Custom`] (`c:custSplit`).
    pub custom_split: Vec<u32>,
    /// The size of the second pie/bar plot, as a percentage
    /// (`c:secondPieSize`).
    pub second_pie_size: Option<u32>,
    /// Drop lines (`c:dropLines`), for line and area charts.
    pub drop_lines: Option<ChartLines>,
    /// High-low lines (`c:hiLowLines`), for line and stock charts.
    pub hi_low_lines: Option<ChartLines>,
    /// Series connector lines (`c:serLines`), for stacked bar and
    /// pie-of-pie charts.
    pub series_lines: Option<ChartLines>,
    /// Up/down bars (`c:upDownBars`), for line and stock charts.
    pub up_down_bars: Option<ChartUpDownBars>,
}

/// A single data series in a chart.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartSeries {
    /// The series index (`c:idx`).
    pub index: Option<u32>,
    /// The plot order of the series (`c:order`).
    pub order: Option<u32>,
    /// The series name (`c:tx`), with its formula and/or cached string.
    pub name: Option<ChartDataSource>,
    /// The category (X) data (`c:cat` or `c:xVal`).
    pub categories: Option<ChartDataSource>,
    /// The value (Y) data (`c:val` or `c:yVal`).
    pub values: Option<ChartDataSource>,
    /// The bubble size data of a bubble chart (`c:bubbleSize`).
    pub bubble_sizes: Option<ChartDataSource>,
    /// The series shape formatting (fill and line).
    pub format: Option<ChartFormat>,
    /// The series marker for line/scatter/radar charts.
    pub marker: Option<ChartMarker>,
    /// Whether the series line is smoothed (`c:smooth`).
    pub smooth: Option<bool>,
    /// Whether negative values invert the fill color
    /// (`c:invertIfNegative`).
    pub invert_if_negative: Option<bool>,
    /// Per-data-point formatting overrides (`c:dPt`).
    pub points: Vec<ChartDataPoint>,
    /// The data labels of the series (`c:dLbls`).
    pub data_labels: Option<ChartDataLabels>,
    /// The trendlines of the series (`c:trendline`).
    pub trendlines: Vec<ChartTrendline>,
    /// The error bars of the series (`c:errBars`), up to one per
    /// direction.
    pub error_bars: Vec<ChartErrorBars>,
    /// Pie/doughnut slice offset from center, as a percentage
    /// (`c:explosion`).
    pub explosion: Option<u32>,
    /// Whether the bubbles of a bubble chart series are drawn in 3D
    /// (`c:bubble3D`).
    pub bubble_3d: Option<bool>,
    /// Chart-ex series layout options (`cx:layoutPr`): binning,
    /// subtotals, statistics, element visibility and parent label
    /// layout. `None` for classic charts.
    pub chart_ex: Option<ChartExLayout>,
}

impl ChartSeries {
    /// The series name as plain text, from the cached value of `c:tx`.
    pub fn name_text(&self) -> Option<&str> {
        let name = self.name.as_ref()?;
        name.values.iter().find_map(|v| match v {
            Data::String(s) => Some(s.as_str()),
            _ => None,
        })
    }
}

/// A data reference used by a chart series: the source formula plus the
/// values cached in the chart part.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartDataSource {
    /// The source range formula, e.g. `Sheet1!$B$2:$B$7` (`c:f`). `None` for
    /// literal (embedded) data.
    pub formula: Option<String>,
    /// The cached data points, indexed by point index. Gaps are
    /// [`Data::Empty`].
    ///
    /// For multi-level (hierarchical) sources this is the innermost
    /// (leaf) level; see [`levels`](Self::levels) for the full
    /// hierarchy.
    pub values: Vec<Data>,
    /// All label levels of a multi-level (hierarchical) category source
    /// (`c:multiLvlStrCache` levels, or the `cx:lvl` levels of a
    /// chart-ex dimension), innermost level first. Empty for
    /// single-level sources.
    pub levels: Vec<Vec<Data>>,
    /// The cached number format code (`c:formatCode`).
    pub number_format: Option<String>,
}

/// A per-data-point override inside a series (`c:dPt`).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartDataPoint {
    /// The zero-based index of the data point this override applies to.
    pub index: u32,
    /// The shape formatting of this data point.
    pub format: Option<ChartFormat>,
}

/// The type of a chart axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartAxisType {
    /// A category axis (`c:catAx`).
    Category,
    /// A value axis (`c:valAx`).
    Value,
    /// A date axis (`c:dateAx`).
    Date,
    /// A series (depth) axis of a 3D chart (`c:serAx`).
    Series,
}

/// The side of the plot area an axis is drawn on (`c:axPos`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartAxisPosition {
    /// Bottom of the plot area.
    Bottom,
    /// Left of the plot area.
    Left,
    /// Right of the plot area.
    Right,
    /// Top of the plot area.
    Top,
}

/// A chart axis.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartAxis {
    /// The axis type.
    pub axis_type: ChartAxisType,
    /// The axis id (`c:axId`), matched by [`ChartGroup::axis_ids`].
    pub id: Option<u32>,
    /// The side of the plot area the axis is drawn on.
    pub position: Option<ChartAxisPosition>,
    /// The axis title.
    pub title: Option<ChartTitle>,
    /// The axis number format code (`c:numFmt`).
    pub number_format: Option<String>,
    /// The minimum axis bound (`c:min`).
    pub min: Option<f64>,
    /// The maximum axis bound (`c:max`).
    pub max: Option<f64>,
    /// The major unit (tick interval) of the axis (`c:majorUnit`).
    pub major_unit: Option<f64>,
    /// The minor unit of the axis (`c:minorUnit`).
    pub minor_unit: Option<f64>,
    /// The logarithmic base, if the axis uses a log scale (`c:logBase`).
    pub log_base: Option<f64>,
    /// Whether the axis direction is reversed (`c:orientation
    /// val="maxMin"`).
    pub reverse: bool,
    /// Whether the axis is hidden (`c:delete val="1"`).
    pub hidden: bool,
    /// Whether major gridlines are shown.
    pub major_gridlines: bool,
    /// Whether minor gridlines are shown.
    pub minor_gridlines: bool,
    /// The major tick mark type (`c:majorTickMark`).
    pub major_tick_mark: Option<ChartTickMark>,
    /// The minor tick mark type (`c:minorTickMark`).
    pub minor_tick_mark: Option<ChartTickMark>,
    /// Where the tick labels are drawn (`c:tickLblPos`).
    pub tick_label_position: Option<ChartTickLabelPosition>,
    /// Where the perpendicular axis crosses this axis (`c:crosses`).
    pub crosses: Option<ChartAxisCrosses>,
    /// The crossing value when [`crosses`](Self::crosses) is
    /// [`ChartAxisCrosses::At`] (`c:crossesAt`).
    pub crosses_at: Option<f64>,
    /// The [`id`](Self::id) of the perpendicular axis this axis crosses
    /// (`c:crossAx`).
    pub crosses_axis_id: Option<u32>,
    /// Whether the value axis crosses between or on category ticks
    /// (`c:crossBetween`).
    pub cross_between: Option<ChartCrossBetween>,
    /// The display units of a value axis (`c:dispUnits`).
    pub display_units: Option<ChartDisplayUnits>,
    /// Whether the display units label is shown (`c:dispUnitsLbl`).
    pub display_units_label: bool,
    /// The interval between tick labels on a category axis
    /// (`c:tickLblSkip`).
    pub tick_label_skip: Option<u32>,
    /// The interval between tick marks on a category axis
    /// (`c:tickMarkSkip`).
    pub tick_mark_skip: Option<u32>,
    /// The label offset of a category axis, as a percentage
    /// (`c:lblOffset`).
    pub label_offset: Option<u32>,
    /// The axis line and area formatting.
    pub format: Option<ChartFormat>,
    /// The font of the axis labels.
    pub font: Option<Font>,
    /// The rotation of the axis label text in degrees
    /// (`c:txPr/a:bodyPr@rot`, converted from 1/60000ths of a degree).
    pub text_rotation: Option<f64>,
}

impl ChartAxis {
    pub(crate) fn new(axis_type: ChartAxisType) -> Self {
        Self {
            axis_type,
            id: None,
            position: None,
            title: None,
            number_format: None,
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            log_base: None,
            reverse: false,
            hidden: false,
            major_gridlines: false,
            minor_gridlines: false,
            major_tick_mark: None,
            minor_tick_mark: None,
            tick_label_position: None,
            crosses: None,
            crosses_at: None,
            cross_between: None,
            display_units: None,
            display_units_label: false,
            tick_label_skip: None,
            tick_mark_skip: None,
            label_offset: None,
            crosses_axis_id: None,
            format: None,
            font: None,
            text_rotation: None,
        }
    }
}

/// A chart or axis title.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartTitle {
    /// The title text with per-run formatting, for rich-text titles.
    pub rich: Option<RichText>,
    /// The source formula, when the title comes from a cell reference.
    pub formula: Option<String>,
    /// The cached title string, when the title comes from a cell reference.
    pub cached: Option<String>,
    /// Whether the title overlays the plot area (`c:overlay`).
    pub overlay: bool,
    /// The default title font.
    pub font: Option<Font>,
    /// The rotation of the title text in degrees
    /// (`a:bodyPr@rot`, converted from 1/60000ths of a degree).
    pub text_rotation: Option<f64>,
}

impl ChartTitle {
    /// The title as plain text, from either the rich text or the cached
    /// string.
    pub fn text(&self) -> Option<String> {
        if let Some(rich) = &self.rich {
            let text = rich.plain_text();
            if !text.is_empty() {
                return Some(text);
            }
        }
        self.cached.clone()
    }
}

/// The position of a chart legend (`c:legendPos`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartLegendPosition {
    /// Right of the plot area (Excel's default).
    #[default]
    Right,
    /// Left of the plot area.
    Left,
    /// Above the plot area.
    Top,
    /// Below the plot area.
    Bottom,
    /// In the top-right corner of the plot area.
    TopRight,
}

/// A chart legend.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartLegend {
    /// Where the legend is positioned.
    pub position: ChartLegendPosition,
    /// Whether the legend overlays the plot area.
    pub overlay: bool,
    /// The legend font.
    pub font: Option<Font>,
}

/// The 3D view settings of a 3D chart (`c:view3D`).
///
/// These correspond to the options in Excel's "3-D Rotation" dialog.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartView3d {
    /// Rotation around the X axis, in degrees (`c:rotX`).
    pub rot_x: Option<i32>,
    /// Rotation around the Y axis, in degrees (`c:rotY`).
    pub rot_y: Option<i32>,
    /// Perspective, in degrees; used when right-angle axes are off
    /// (`c:perspective`, stored as half-degrees in the file and converted).
    pub perspective: Option<u32>,
    /// Depth of the chart as a percentage of its width (`c:depthPercent`).
    pub depth_percent: Option<u32>,
    /// Height of the chart as a percentage of its width (`c:hPercent`).
    pub height_percent: Option<u32>,
    /// Whether the chart axes are drawn at right angles (`c:rAngAx`).
    pub right_angle_axes: Option<bool>,
}

/// Shape formatting of a chart element: area fill and line/border.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartFormat {
    /// The area fill.
    pub fill: Option<ChartFill>,
    /// The line (or border) formatting.
    pub line: Option<ChartLine>,
    /// The linear gradient angle in degrees (`a:gradFill/a:lin@ang`,
    /// converted from 1/60000ths of a degree), when
    /// [`fill`](Self::fill) is [`ChartFill::Gradient`].
    pub gradient_angle: Option<f64>,
}

/// The fill of a chart element.
#[derive(Debug, Clone, PartialEq)]
pub enum ChartFill {
    /// No fill (transparent).
    None,
    /// A solid color fill.
    Solid(Color),
    /// A gradient fill with its color stops.
    Gradient(Vec<ChartGradientStop>),
    /// A pattern fill.
    Pattern {
        /// The pattern preset name, e.g. `pct50` (`a:pattFill@prst`).
        pattern: String,
        /// The foreground color.
        foreground: Option<Color>,
        /// The background color.
        background: Option<Color>,
    },
}

/// One color stop in a gradient fill.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartGradientStop {
    /// The stop position, 0.0 to 100.0 percent.
    pub position: f64,
    /// The stop color.
    pub color: Color,
}

/// The dash type of a chart line (`a:prstDash`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[non_exhaustive]
pub enum ChartLineDashType {
    /// A solid line.
    #[default]
    Solid,
    /// A dotted line.
    Dot,
    /// A dashed line.
    Dash,
    /// A dash-dot line.
    DashDot,
    /// A long dash line.
    LongDash,
    /// A long dash-dot line.
    LongDashDot,
    /// A long dash-dot-dot line.
    LongDashDotDot,
    /// A system dashed line.
    SystemDash,
    /// A system dotted line.
    SystemDot,
    /// A system dash-dot line.
    SystemDashDot,
    /// A system dash-dot-dot line.
    SystemDashDotDot,
}

/// The line (or border) formatting of a chart element.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartLine {
    /// The line color.
    pub color: Option<Color>,
    /// The line width in points.
    pub width: Option<f64>,
    /// The line dash type.
    pub dash_type: Option<ChartLineDashType>,
    /// Whether the line is explicitly hidden (`a:noFill` inside `a:ln`).
    pub hidden: bool,
}

/// The marker symbol of a line/scatter/radar series (`c:symbol`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
#[non_exhaustive]
pub enum ChartMarkerType {
    /// An automatically assigned marker.
    #[default]
    Automatic,
    /// A circle marker.
    Circle,
    /// A dash marker.
    Dash,
    /// A diamond marker.
    Diamond,
    /// A dot marker.
    Dot,
    /// A plus marker.
    Plus,
    /// A square marker.
    Square,
    /// A star marker.
    Star,
    /// A triangle marker.
    Triangle,
    /// An X marker.
    X,
    /// No marker.
    None,
}

/// The marker of a line/scatter/radar series.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartMarker {
    /// The marker symbol.
    pub marker_type: ChartMarkerType,
    /// The marker size in points (2-72).
    pub size: Option<u8>,
    /// The marker fill and outline formatting.
    pub format: Option<ChartFormat>,
}

/// One corner of a chart anchor: a cell plus an EMU offset into it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ChartCellAnchor {
    /// Zero-based column index.
    pub col: u32,
    /// Zero-based row index.
    pub row: u32,
    /// Horizontal offset into the cell, in EMUs (914400 per inch).
    pub col_offset: i64,
    /// Vertical offset into the cell, in EMUs.
    pub row_offset: i64,
}

/// Where a chart is anchored on the worksheet.
///
/// Charts anchored with a `twoCellAnchor` have both [`from`](Self::from) and
/// [`to`](Self::to). Charts anchored with a `oneCellAnchor` have
/// [`from`](Self::from) plus a size, and `absoluteAnchor` charts (typical
/// for chartsheets) only have a size.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ChartPosition {
    /// The top-left anchor cell.
    pub from: Option<ChartCellAnchor>,
    /// The bottom-right anchor cell.
    pub to: Option<ChartCellAnchor>,
    /// The chart width in EMUs, for one-cell and absolute anchors.
    pub width: Option<u64>,
    /// The chart height in EMUs, for one-cell and absolute anchors.
    pub height: Option<u64>,
    /// How the chart moves/resizes with the grid
    /// (`xdr:twoCellAnchor@editAs`). `Some(ChartEditAs::TwoCell)` for
    /// two-cell anchors without an explicit attribute (the spec
    /// default); `None` for one-cell and absolute anchors.
    pub edit_as: Option<ChartEditAs>,
    /// The absolute X position in EMUs (`xdr:absoluteAnchor/xdr:pos@x`).
    pub x: Option<i64>,
    /// The absolute Y position in EMUs (`xdr:absoluteAnchor/xdr:pos@y`).
    pub y: Option<i64>,
}

/// How a two-cell anchored chart behaves when the grid changes
/// (`xdr:twoCellAnchor@editAs`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartEditAs {
    /// The chart keeps its absolute position and size.
    Absolute,
    /// The chart moves with its top-left cell but keeps its size.
    OneCell,
    /// The chart moves and resizes with its anchor cells.
    TwoCell,
}

/// How empty cells are plotted (`c:dispBlanksAs`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartDisplayBlanksAs {
    /// Empty cells leave a gap.
    Gap,
    /// Lines span across empty cells.
    Span,
    /// Empty cells are plotted as zero.
    Zero,
}

/// The data table shown under a chart (`c:dTable`).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartDataTable {
    /// Whether horizontal borders are shown.
    pub show_horizontal_border: bool,
    /// Whether vertical borders are shown.
    pub show_vertical_border: bool,
    /// Whether the table outline is shown.
    pub show_outline: bool,
    /// Whether legend keys are shown next to the series names.
    pub show_legend_keys: bool,
    /// The table font.
    pub font: Option<Font>,
}

/// The position of data labels relative to their data points
/// (`c:dLblPos`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ChartDataLabelPosition {
    /// Centered on the data point.
    Center,
    /// Inside the end of the data point.
    InsideEnd,
    /// Inside the base of the data point.
    InsideBase,
    /// Outside the end of the data point.
    OutsideEnd,
    /// Left of the data point.
    Left,
    /// Right of the data point.
    Right,
    /// Above the data point.
    Above,
    /// Below the data point.
    Below,
    /// Positioned automatically (pie charts).
    BestFit,
}

/// The data labels of a series or plot group (`c:dLbls`).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartDataLabels {
    /// Whether the point values are shown.
    pub show_value: bool,
    /// Whether the category names are shown.
    pub show_category_name: bool,
    /// Whether the series name is shown.
    pub show_series_name: bool,
    /// Whether the legend key is shown next to each label.
    pub show_legend_key: bool,
    /// Whether percentages are shown (pie/doughnut charts).
    pub show_percent: bool,
    /// Whether bubble sizes are shown (bubble charts).
    pub show_bubble_size: bool,
    /// The label position.
    pub position: Option<ChartDataLabelPosition>,
    /// The label number format code (`c:numFmt`).
    pub number_format: Option<String>,
    /// The label font.
    pub font: Option<Font>,
    /// The label area/border formatting.
    pub format: Option<ChartFormat>,
    /// The rotation of the label text in degrees
    /// (`c:txPr/a:bodyPr@rot`, converted from 1/60000ths of a degree).
    pub text_rotation: Option<f64>,
    /// Per-point label overrides (`c:dLbl`).
    pub point_labels: Vec<ChartDataLabel>,
}

/// A per-point data label override (`c:dLbl`).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartDataLabel {
    /// The zero-based index of the data point this label belongs to
    /// (`c:idx`).
    pub index: u32,
    /// Whether the label is deleted (hidden) for this point
    /// (`c:delete`).
    pub delete: bool,
    /// Custom label text (`c:tx/c:rich`), flattened to plain text.
    pub text: Option<String>,
    /// The label position, when overridden for this point.
    pub position: Option<ChartDataLabelPosition>,
    /// The label number format code, when overridden (`c:numFmt`).
    pub number_format: Option<String>,
    /// The label font, when overridden.
    pub font: Option<Font>,
    /// The label area/border formatting, when overridden.
    pub format: Option<ChartFormat>,
}

/// The type of a series trendline (`c:trendlineType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartTrendlineType {
    /// An exponential trendline.
    Exponential,
    /// A linear trendline.
    #[default]
    Linear,
    /// A logarithmic trendline.
    Logarithmic,
    /// A moving-average trendline.
    MovingAverage,
    /// A polynomial trendline.
    Polynomial,
    /// A power trendline.
    Power,
}

/// A trendline attached to a chart series (`c:trendline`).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartTrendline {
    /// The trendline type.
    pub trendline_type: ChartTrendlineType,
    /// The custom trendline name (`c:name`).
    pub name: Option<String>,
    /// The polynomial order, 2-6 (`c:order`).
    pub order: Option<u32>,
    /// The moving-average period (`c:period`).
    pub period: Option<u32>,
    /// Forecast periods forward (`c:forward`).
    pub forward: Option<f64>,
    /// Forecast periods backward (`c:backward`).
    pub backward: Option<f64>,
    /// The forced Y-axis intercept (`c:intercept`).
    pub intercept: Option<f64>,
    /// Whether the trendline equation is displayed (`c:dispEq`).
    pub display_equation: bool,
    /// Whether the R-squared value is displayed (`c:dispRSqr`).
    pub display_r_squared: bool,
    /// The trendline formatting.
    pub format: Option<ChartFormat>,
}

/// The direction of series error bars (`c:errDir`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartErrorBarsDirection {
    /// Horizontal (X) error bars.
    X,
    /// Vertical (Y) error bars.
    Y,
}

/// Which directions the error bars extend in (`c:errBarType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartErrorBarsType {
    /// Both plus and minus.
    #[default]
    Both,
    /// Minus only.
    Minus,
    /// Plus only.
    Plus,
}

/// How the error amount is determined (`c:errValType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartErrorBarsValueType {
    /// Custom plus/minus data ranges.
    Custom,
    /// A fixed value.
    #[default]
    FixedValue,
    /// A percentage of each value.
    Percentage,
    /// A number of standard deviations.
    StandardDeviation,
    /// The standard error.
    StandardError,
}

/// Error bars attached to a chart series (`c:errBars`).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartErrorBars {
    /// The direction of the error bars.
    pub direction: Option<ChartErrorBarsDirection>,
    /// Which directions the bars extend in.
    pub error_type: ChartErrorBarsType,
    /// How the error amount is determined.
    pub value_type: ChartErrorBarsValueType,
    /// The fixed/percentage/standard-deviation amount (`c:val`).
    pub value: Option<f64>,
    /// Whether the bars are drawn without end caps (`c:noEndCap`).
    pub no_end_cap: bool,
    /// Custom plus values (`c:plus`).
    pub plus_values: Option<ChartDataSource>,
    /// Custom minus values (`c:minus`).
    pub minus_values: Option<ChartDataSource>,
    /// The error bar formatting.
    pub format: Option<ChartFormat>,
}

/// The tick mark type of an axis (`c:majorTickMark` /
/// `c:minorTickMark`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartTickMark {
    /// No tick marks.
    None,
    /// Tick marks inside the plot area.
    Inside,
    /// Tick marks outside the plot area.
    Outside,
    /// Tick marks crossing the axis.
    Cross,
}

/// Where axis tick labels are drawn (`c:tickLblPos`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartTickLabelPosition {
    /// Next to the axis (the default).
    NextTo,
    /// At the high end of the perpendicular axis.
    High,
    /// At the low end of the perpendicular axis.
    Low,
    /// Not shown.
    None,
}

/// Where the perpendicular axis crosses an axis (`c:crosses`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartAxisCrosses {
    /// At zero (or the automatic position).
    AutoZero,
    /// At the minimum value.
    Min,
    /// At the maximum value.
    Max,
    /// At the value given by [`ChartAxis::crosses_at`].
    At,
}

/// Whether a value axis crosses between or on category ticks
/// (`c:crossBetween`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartCrossBetween {
    /// The axis crosses between categories.
    Between,
    /// The axis crosses at the category midpoints.
    MidCat,
}

/// The display units of a value axis (`c:builtInUnit` /
/// `c:custUnit`).
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ChartDisplayUnits {
    /// Hundreds.
    Hundreds,
    /// Thousands.
    Thousands,
    /// Tens of thousands.
    TenThousands,
    /// Hundreds of thousands.
    HundredThousands,
    /// Millions.
    Millions,
    /// Tens of millions.
    TenMillions,
    /// Hundreds of millions.
    HundredMillions,
    /// Billions.
    Billions,
    /// Trillions.
    Trillions,
    /// A custom unit divisor.
    Custom(f64),
}

/// The 3D shape of the bars/columns of a 3D bar or column chart
/// (`c:shape`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartBar3dShape {
    /// A box (the default).
    #[default]
    Box,
    /// A cone tapering to a point.
    Cone,
    /// A cone truncated at the maximum value.
    ConeToMax,
    /// A cylinder.
    Cylinder,
    /// A pyramid tapering to a point.
    Pyramid,
    /// A pyramid truncated at the maximum value.
    PyramidToMax,
}

/// How the bubble size data maps to the drawn bubbles
/// (`c:sizeRepresents`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ChartSizeRepresents {
    /// Bubble size data sets the bubble area (the default).
    #[default]
    Area,
    /// Bubble size data sets the bubble width (diameter).
    Width,
}

/// How the second plot of a pie-of-pie or bar-of-pie chart is populated
/// (`c:splitType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartOfPieSplitType {
    /// Split determined automatically.
    Auto,
    /// Split by custom point assignment.
    Custom,
    /// Split by percentage threshold.
    Percent,
    /// Split by position (the last N points).
    Position,
    /// Split by value threshold.
    Value,
}

/// Auxiliary lines of a plot group: drop lines, high-low lines or series
/// lines.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartLines {
    /// The line formatting, if specified.
    pub format: Option<ChartFormat>,
}

/// The up/down bars of a line or stock chart (`c:upDownBars`).
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartUpDownBars {
    /// The gap between the bars, as a percentage (`c:gapWidth`).
    pub gap_width: Option<u32>,
    /// The formatting of the up bars.
    pub up_format: Option<ChartFormat>,
    /// The formatting of the down bars.
    pub down_format: Option<ChartFormat>,
}

/// The layout options of a chart-ex series (`cx:layoutPr`).
///
/// Which fields are populated depends on the chart type: binning for
/// histogram/pareto, subtotals and connector lines for waterfall,
/// statistics and mean/outlier visibility for box & whisker, and parent
/// label layout for treemap.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartExLayout {
    /// How parent labels are laid out in a treemap
    /// (`cx:parentLabelLayout`).
    pub parent_label_layout: Option<ChartExParentLabelLayout>,
    /// The histogram bin size (`cx:binning/cx:binSize`).
    pub bin_size: Option<f64>,
    /// The histogram bin count (`cx:binning/cx:binCount`).
    pub bin_count: Option<u32>,
    /// The histogram overflow bin threshold (`cx:binning@overflow`).
    /// `None` when automatic.
    pub overflow: Option<f64>,
    /// The histogram underflow bin threshold (`cx:binning@underflow`).
    /// `None` when automatic.
    pub underflow: Option<f64>,
    /// The box & whisker quartile method
    /// (`cx:statistics@quartileMethod`).
    pub quartile_method: Option<ChartExQuartileMethod>,
    /// Whether the box & whisker mean marker is shown
    /// (`cx:visibility@meanMarker`).
    pub mean_marker: Option<bool>,
    /// Whether the box & whisker mean line is shown
    /// (`cx:visibility@meanLine`).
    pub mean_line: Option<bool>,
    /// Whether box & whisker outlier points are shown
    /// (`cx:visibility@outliers`).
    pub outliers: Option<bool>,
    /// Whether box & whisker non-outlier points are shown
    /// (`cx:visibility@nonoutliers`).
    pub non_outliers: Option<bool>,
    /// Whether waterfall connector lines are shown
    /// (`cx:visibility@connectorLines`).
    pub connector_lines: Option<bool>,
    /// The zero-based point indices treated as waterfall subtotals
    /// (`cx:subtotals`).
    pub subtotals: Vec<u32>,
}

/// How parent labels are laid out in a treemap chart
/// (`cx:parentLabelLayout@val`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ChartExParentLabelLayout {
    /// Parent labels are not shown.
    None,
    /// Parent labels are shown as banners above their group.
    Banner,
    /// Parent labels overlap their group.
    Overlapping,
}

/// The quartile calculation method of a box & whisker chart
/// (`cx:statistics@quartileMethod`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartExQuartileMethod {
    /// Inclusive median quartile calculation.
    Inclusive,
    /// Exclusive median quartile calculation.
    Exclusive,
}

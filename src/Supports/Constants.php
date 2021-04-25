<?php

namespace Lysice\XlsWriter\Supports;

class Constants {
    // -----------------------excel constants grid line-------------------------------//
    const GRIDLINES_HIDE_ALL    = 0; // hide and print gridLine
    const GRIDLINES_SHOW_SCREEN = 1; // show gridLine
    const GRIDLINES_SHOW_PRINT  = 2; // show print gridLine
    const GRIDLINES_SHOW_ALL    = 3; // show print and screen gridLine

    // -----------------------chart constants start-------------------------------//
    // area chart
    const CHART_AREA = 1;
    const CHART_AREA_STACKED = 2;
    const CHART_AREA_PERCENT = 3;
    // bar chart
    const CHART_BAR = 4;
    const CHART_BAR_STACKED = 5;
    const CHART_BAR_STACKED_PERCENT = 6;
    // column chart
    const CHART_COLUMN = 7;
    const CHART_COLUMN_STACKED = 8;
    const CHART_COLUMN_STACKED_PERCENT = 9;

    const CHART_DOUGHNUT = 10;
    const CHART_LINE = 11;
    const CHART_PIE = 12;
    const CHART_SCATTER = 13;
    const CHART_SCATTER_STRAIGHT = 14;
    const CHART_SCATTER_STRAIGHT_WITH_MARKERS = 15;
    const CHART_SCATTER_SMOOTH = 16;
    const CHART_SCATTER_SMOOTH_WITH_MARKERS = 17;
    const CHART_RADAR = 18;
    const CHART_RADAR_WITH_MARKERS = 19;
    const CHART_RADAR_FILLED = 20;
    // -----------------------chart constants start-------------------------------//


    // -----------------------format constants start-------------------------------//
    // border type
    const BORDER_THIN = 1;
    const BORDER_MEDIUM = 2;
    const BORDER_DASHED = 3;
    const BORDER_DOTTED = 4;
    const BORDER_THICK = 5;
    const BORDER_DOUBLE = 6;
    const BORDER_HAIR= 7;
    const BORDER_MEDIUM_DASHED = 8;
    const BORDER_DASH_DOT = 9;
    const BORDER_MEDIUM_DASH_DOT = 10;
    const BORDER_DASH_DOT_DOT = 11;
    const BORDER_MEDIUM_DASH_DOT_DOT = 12;
    const BORDER_SLANT_DASH_DOT = 13;

    // align type
    const ALIGN_LEFT = 1;
    const ALIGN_CENTER = 2;
    const ALIGN_RIGHT = 3;
    const ALIGN_FILL = 4;
    const ALIGN_JUSTIFY = 5;
    const ALIGN_CENTER_ACROSS = 6;
    const ALIGN_DISTRIBUTED = 7;
    const ALIGN_VERTICAL_TOP = 8;
    const ALIGN_VERTICAL_BOTTOM = 9;
    const ALIGN_VERTICAL_CENTER = 10;
    const ALIGN_VERTICAL_JUSTIFY = 11;
    const ALIGN_VERTICAL_DISTRIBUTED = 12;

    // underline type
    const UNDERLINE_SINGLE = 1;
    const UNDERLINE_DOUBLE = 2;
    const UNDERLINE_SINGLE_ACCOUNTING = 3;
    const UNDERLINE_DOUBLE_ACCOUNTING = 4;

    // colors
    const COLOR_BLACK = \Vtiful\Kernel\Format::COLOR_BLACK;
    const COLOR_BROWN = \Vtiful\Kernel\Format::COLOR_BROWN;
    const COLOR_BLUE = \Vtiful\Kernel\Format::COLOR_BLUE;
    const COLOR_CYAN = \Vtiful\Kernel\Format::COLOR_CYAN;
    const COLOR_GRAY = \Vtiful\Kernel\Format::COLOR_GRAY;
    const COLOR_LIME = \Vtiful\Kernel\Format::COLOR_LIME;
    const COLOR_GREEN = \Vtiful\Kernel\Format::COLOR_GREEN;
    const COLOR_YELLOW = \Vtiful\Kernel\Format::COLOR_YELLOW;
    const COLOR_WHITE = \Vtiful\Kernel\Format::COLOR_WHITE;
    const COLOR_SILVER = \Vtiful\Kernel\Format::COLOR_SILVER;
    const COLOR_RED = \Vtiful\Kernel\Format::COLOR_RED;
    const COLOR_PURPLE = \Vtiful\Kernel\Format::COLOR_PURPLE;
    const COLOR_PINK = \Vtiful\Kernel\Format::COLOR_PINK;
    const COLOR_ORANGE = \Vtiful\Kernel\Format::COLOR_ORANGE;
    const COLOR_NAVY = \Vtiful\Kernel\Format::COLOR_NAVY;
    const COLOR_MAGENTA = \Vtiful\Kernel\Format::COLOR_MAGENTA;

    // background patterns
    const PATTERN_NONE = \Vtiful\Kernel\Format::PATTERN_NONE;
    const PATTERN_SOLID = \Vtiful\Kernel\Format::PATTERN_SOLID;
    const PATTERN_MEDIUM_GRAY = \Vtiful\Kernel\Format::PATTERN_MEDIUM_GRAY;
    const PATTERN_DARK_GRAY = \Vtiful\Kernel\Format::PATTERN_DARK_GRAY;
    const PATTERN_LIGHT_GRAY = \Vtiful\Kernel\Format::PATTERN_LIGHT_GRAY;
    const PATTERN_DARK_HORIZONTAL = \Vtiful\Kernel\Format::PATTERN_DARK_HORIZONTAL;
    const PATTERN_DARK_VERTICAL = \Vtiful\Kernel\Format::PATTERN_DARK_VERTICAL;
    const PATTERN_DARK_DOWN = \Vtiful\Kernel\Format::PATTERN_DARK_DOWN;
    const PATTERN_DARK_UP = \Vtiful\Kernel\Format::PATTERN_DARK_UP;
    const PATTERN_DARK_GRID = \Vtiful\Kernel\Format::PATTERN_DARK_GRID;
    const PATTERN_DARK_TRELLIS = \Vtiful\Kernel\Format::PATTERN_DARK_TRELLIS;
    const PATTERN_LIGHT_HORIZONTAL = \Vtiful\Kernel\Format::PATTERN_LIGHT_HORIZONTAL;
    const PATTERN_LIGHT_VERTICAL = \Vtiful\Kernel\Format::PATTERN_LIGHT_VERTICAL;
    const PATTERN_LIGHT_DOWN = \Vtiful\Kernel\Format::PATTERN_LIGHT_DOWN;
    const PATTERN_LIGHT_UP = \Vtiful\Kernel\Format::PATTERN_LIGHT_UP;
    const PATTERN_LIGHT_GRID = \Vtiful\Kernel\Format::PATTERN_LIGHT_GRID;
    const PATTERN_LIGHT_TRELLIS = \Vtiful\Kernel\Format::PATTERN_LIGHT_TRELLIS;
    const PATTERN_GRAY_125 = \Vtiful\Kernel\Format::PATTERN_GRAY_125;
    const PATTERN_GRAY_0625 = \Vtiful\Kernel\Format::PATTERN_GRAY_0625;
    // -----------------------format const  end-------------------------------//
}

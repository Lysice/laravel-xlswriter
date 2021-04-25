<?php

namespace Lysice\XlsWriter\Supports\Format;

use Lysice\XlsWriter\Exceptions\FormatParamErrorException;
use Lysice\XlsWriter\Interfaces\FormatInterface;
use Lysice\XlsWriter\Supports\Constants;

/**
 * Class Format
 * @package Lysice\XlsWriter\Supports
 */
class Format implements FormatInterface {
    /**
     * @var \Vtiful\Kernel\Format
     */
    protected $format;

    /**
     * format border types
     */
    const TYPE_BORDER = [
        Constants::BORDER_THIN,
        Constants::BORDER_MEDIUM,
        Constants::BORDER_DASHED,
        Constants::BORDER_DOTTED,
        Constants::BORDER_THICK,
        Constants::BORDER_DOUBLE,
        Constants::BORDER_HAIR,
        Constants::BORDER_MEDIUM_DASHED,
        Constants::BORDER_DASH_DOT,
        Constants::BORDER_MEDIUM_DASH_DOT,
        Constants::BORDER_DASH_DOT_DOT,
        Constants::BORDER_MEDIUM_DASH_DOT_DOT,
        Constants::BORDER_SLANT_DASH_DOT
    ];

    /**
     * format align types
     */
    const TYPE_ALIGN = [
        Constants::ALIGN_LEFT,
        Constants::ALIGN_CENTER,
        Constants::ALIGN_RIGHT,
        Constants::ALIGN_FILL,
        Constants::ALIGN_JUSTIFY,
        Constants::ALIGN_CENTER_ACROSS,
        Constants::ALIGN_DISTRIBUTED,
        Constants::ALIGN_VERTICAL_TOP,
        Constants::ALIGN_VERTICAL_BOTTOM,
        Constants::ALIGN_VERTICAL_CENTER,
        Constants::ALIGN_VERTICAL_JUSTIFY,
        Constants::ALIGN_VERTICAL_DISTRIBUTED
    ];

    /**
     * underline style
     */
    const STYLE_UNDERLINE = [
        Constants::UNDERLINE_SINGLE,
        Constants::UNDERLINE_DOUBLE,
        Constants::UNDERLINE_SINGLE_ACCOUNTING,
        Constants::UNDERLINE_DOUBLE_ACCOUNTING
    ];

    /**
     * format color
     */
    const STYLE_COLORS = [
        Constants::COLOR_BLACK,
        Constants::COLOR_BROWN,
        Constants::COLOR_BLUE,
        Constants::COLOR_CYAN,
        Constants::COLOR_GRAY,
        Constants::COLOR_LIME,
        Constants::COLOR_GREEN,
        Constants::COLOR_YELLOW,
        Constants::COLOR_WHITE,
        Constants::COLOR_SILVER,
        Constants::COLOR_RED,
        Constants::COLOR_PURPLE,
        Constants::COLOR_PINK,
        Constants::COLOR_ORANGE,
        Constants::COLOR_NAVY,
        Constants::COLOR_MAGENTA
    ];

    /**
     * number format
     */
    const STYLE_NUMBER = [
        "0.000",
        "#,##0",
        "#,##0.00",
        "0.00"
    ];

    /**
     * get format instance
     * @param $fileHandle
     * @return static
     * @throws FormatParamErrorException
     */
    public static function createFormat($fileHandle)
    {
        $instance = new static();
        if (!is_resource($fileHandle)) {
            throw new FormatParamErrorException('fileHandle is not a resource');
        }
        $instance->format = new \Vtiful\Kernel\Format($fileHandle);
        return $instance;
    }

    /**
     * bold style
     */
    public function bold()
    {
        $this->format->bold();
        return $this;
    }

    /**
     * italic
     */
    public function italic()
    {
        $this->format->italic();
        return $this;
    }

    /**
     * @param $border int
     * @return $this
     * @throws FormatParamErrorException
     */
    public function border(int $border)
    {
        if (empty($border) or !in_array($border, static::TYPE_BORDER)) {
            throw new FormatParamErrorException("Border type doesn't support");
        }

        $this->format->border($border);
        return $this;
    }

    /**
     * align
     * @param mixed ...$align
     * @return $this
     * @throws FormatParamErrorException
     */
    public function align(...$align)
    {
        if (is_array($align)) {
            foreach ($align as $item) {
                if (empty($item) or !in_array($item, static::TYPE_ALIGN)) {
                    throw new FormatParamErrorException("Align type ".$item."doesn't support");
                }
            }
        } else {
            if (empty($align) or !in_array($align, static::TYPE_ALIGN)) {
                throw new FormatParamErrorException("Align type doesn't support");
            }
        }

        $this->format->align($align);
        return $this;
    }

    /**
     * @param string $fontName
     * @return $this
     */
    public function font($fontName = '')
    {
        $this->format->font($fontName);
        return $this;
    }

    /**
     * @param $fontColor
     * @return $this
     * @throws FormatParamErrorException
     */
    public function fontColor($fontColor)
    {
        if (!in_array($fontColor, static::STYLE_COLORS) && !ctype_xdigit($fontColor)) {
            throw new FormatParamErrorException('fontcolor string is not a valid color');
        }
        $this->format->fontColor($fontColor);
        return $this;
    }

    /**
     * background
     * @param $background
     * @param int $pattern
     * @return $this
     */
    public function background($background, $pattern = 0)
    {
        $this->format->background($background, $pattern);
        return $this;
    }

    /**
     * fontSize
     * @param $size
     * @return $this
     */
    public function fontSize($size)
    {
        $this->format->fontSize($size);
        return $this;
    }

    /**
     * @param $format
     * @return $this
     * @throws FormatParamErrorException
     */
    public function number($format)
    {
        if (!in_array($format, static::STYLE_NUMBER)) {
            throw new FormatParamErrorException('number format does not support!');
        }
        $this->format->number($format);
        return $this;
    }

    /**
     * @param $style
     * @return $this
     * @throws FormatParamErrorException
     */
    public function underline($style)
    {
        if (empty($style) or !in_array($style, static::STYLE_UNDERLINE)) {
            throw new FormatParamErrorException("Underline style ' . $style . ' doesn't support");
        }

        $this->format->underline($style);
        return $this;
    }

    /**
     * wrap
     * @return $this
     */
    public function wrap()
    {
        $this->format->wrap();
        return $this;
    }

    /**
     * strikeout
     * @return $this
     */
    public function strikeout()
    {
        $this->format->strikeout();
        return $this;
    }

    /**
     * get format resource
     * @return resource
     */
    public function toResource()
    {
        return $this->format->toResource();
    }
}

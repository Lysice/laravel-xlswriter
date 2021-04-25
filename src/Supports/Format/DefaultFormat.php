<?php

namespace Lysice\XlsWriter\Supports\Format;

use Lysice\XlsWriter\Interfaces\FormatInterface;
use function foo\func;

/**
 * Class DefaultFormat
 * @package Lysice\XlsWriter\Supports\Format
 */
class DefaultFormat extends Format implements FormatInterface
{
    /**
     * the Format function callback array
     * @var array
     */
    protected $styles = [];

    /**
     * @return static
     */
    public static function create()
    {
        return new static();
    }

    /**
     * @param $styleName
     * @param $callBack
     */
    public function addStyle($styleName, $callBack)
    {
        $this->styles[$styleName] = $callBack;
    }

    /**
     * @return $this|Format
     */
    public function bold()
    {
        $this->addStyle('bold', function ($instance) {
            return $instance->bold();
        });
        return $this;
    }

    /**
     * @return $this|Format
     */
    public function italic()
    {
        $this->addStyle('italic', function($instance) {
            return $instance->italic();
        });
        return $this;
    }

    /**
     * @param int $border
     * @return $this|FormatInterface|Format
     * @throws \Lysice\XlsWriter\Exceptions\FormatParamErrorException
     */
    public function border(int $border)
    {
        $this->addStyle('border', function ($instance) use ($border) {
            return $instance->border($border);
        });
        return $this;
    }

    /**
     * @param mixed ...$align
     * @return $this|Format
     * @throws \Lysice\XlsWriter\Exceptions\FormatParamErrorException
     */
    public function align(...$align)
    {
        $this->addStyle('align', function ($instance) use ($align) {
            return $instance->align($align);
        });
        return $this;
    }

    /**
     * @param string $fontName
     * @return $this|FormatInterface|Format
     */
    public function font($fontName = '')
    {
        $this->addStyle('font', function ($instance) use ($fontName) {
            return $instance->font($fontName);
        });
        return $this;
    }

    /**
     * @param $fontColor
     * @return $this|FormatInterface|Format
     * @throws \Lysice\XlsWriter\Exceptions\FormatParamErrorException
     */
    public function fontColor($fontColor)
    {
        $this->addStyle('fontColor', function ($instance) use ($fontColor) {
            return $instance->fontColor($fontColor);
        });
        return $this;
    }

    /**
     * @param $background
     * @param int $pattern
     * @return $this|Format
     */
    public function background($background, $pattern = 0)
    {
        $this->addStyle('background', function ($instance) use ($background, $pattern){
            return $instance->background($background, $pattern);
        });
        return $this;
    }

    /**
     * @param $size
     * @return $this|Format
     */
    public function fontSize($size)
    {
        $this->addStyle('fontSize', function ($instance) use ($size) {
            return $instance->fontSize($size);
        });
        return $this;
    }

    /**
     * @param $format
     * @return $this|FormatInterface|Format
     */
    public function number($format)
    {
        $this->addStyle('number', function ($instance) use ($format) {
            return $instance->number($format);
        });
        return $this;
    }

    /**
     * @param $style
     * @return $this|FormatInterface|Format
     * @throws \Lysice\XlsWriter\Exceptions\FormatParamErrorException
     */
    public function underline($style)
    {
        $this->addStyle('underline', function ($instance) use ($style) {
            return $instance->underline($style);
        });
        return $this;
    }

    /**
     * @return $this|Format
     */
    public function wrap()
    {
        $this->addStyle('wrap', function ($instance) {
            return $instance->wrap();
        });
        return $this;
    }

    public function strikeout()
    {
        $this->addStyle('strikeout', function ($instance) {
            return $instance->strikeout();
        });
        return $this;
    }

    /**
     * @param $fileHandle
     * @return resource
     * @throws \Lysice\XlsWriter\Exceptions\FormatParamErrorException
     */
    public function getFormatResource($fileHandle)
    {
        $format = Format::createFormat($fileHandle);
        foreach ($this->styles as $key => $callback) {
            $callback($format);
        }
        return $format->toResource();
    }
}

<?php

namespace Lysice\XlsWriter\Interfaces;

interface FormatInterface
{

    /**
     * bold style
     */
    public function bold();

    /**
     * italic
     */
    public function italic();

    /**
     * @param $border int
     * @return $this
     * @throws FormatParamErrorException
     */
    public function border(int $border);

    public function align(...$align);

    /**
     * @param string $fontName
     * @return $this
     */
    public function font($fontName = '');

    /**
     * @param $fontColor
     * @return $this
     * @throws FormatParamErrorException
     */
    public function fontColor($fontColor);

    public function background($background, $pattern);

    public function fontSize($size);

    /**
     * @param $format
     * @return $this
     * @throws FormatParamErrorException
     */
    public function number($format);

    /**
     * @param $style
     * @return $this
     * @throws FormatParamErrorException
     */
    public function underline($style);

    public function wrap();

    public function strikeout();
}

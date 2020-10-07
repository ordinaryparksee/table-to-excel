<?php

namespace Vaxy\TableToExcel;

use ArrayAccess;

class CssParser implements ArrayAccess {

    public static function parse($source) : self
    {
        $style = [];
        if ($source) {
            foreach (explode(';', $source) as $fragment) {
                if ($fragment) {
                    list($key, $value) = explode(':', $fragment, 2);
                    if ($key && $value) {
                        $style[strtolower(trim($key))] = strtolower(trim($value));
                    }
                }
            }
        }
        return new self($style);
    }

    private $attributes = [];

    public function __construct(array $attributes)
    {
        $this->attributes = $attributes;
    }

    public function has($key) : bool
    {
        return array_key_exists($key, $this->attributes);
    }

    public function offsetExists($offset) : bool
    {
        return $this->has($offset);
    }

    public function offsetGet($offset)
    {
        return $this->attributes[$offset];
    }

    public function offsetSet($offset , $value)
    {
        $this->attributes[$offset] = $value;
    }

    public function offsetUnset($offset)
    {
        unset($this->attributes[$offset]);
    }

}

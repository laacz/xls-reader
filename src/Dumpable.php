<?php
namespace laacz\XLSParser;


class Dumpable {
    private $repr_these = [];
    private $slots = [];

    public function dump($header = Null, $footer = Null, $indent = 0)
    {
        $dump = [];
        if (property_exists($this, 'slots')) {
            $alist = [];
            foreach ($this->slots as $attr) {
                $alist[] = $attr;
            }
        } else {
            $alist = array_keys(get_object_vars($this));
        }
        sort($alist);

        $pad = str_repeat(" ", $indent);

        if ($header) $dump[] = $header;
        foreach ($alist as $attr) {
            if ($attr !== 'book' && is_object($this->$attr) && property_exists($this->$attr, 'dump')) {
                $this->$attr->dump(sprintf("%s%s (%s object): ", $pad, $attr, get_class($this->$attr)), null, $indent + 4);
            } elseif (!in_array($attr, $this->repr_these) && is_array($this->$attr)) {
                $dump[] = sprintf("%s%s: array, len = %d", $pad, $attr, count($this->$attr));
            } else {
                $dump[] = sprintf("%s%s: %s", $pad, $attr, is_object($this->$attr) ? get_class($this->$attr) : $this->$attr);
            }
        }
        if ($footer) $dump[] = $footer;
        return join("\n", $dump);
    }
}
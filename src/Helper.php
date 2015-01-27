<?php

namespace laacz\XLSParser;

define('BIG_ENDIAN', (unpack('n', "\x01\x00") == unpack('S', "\x01\x00")));

class Helper
{

    static $log_enabled = false;
    static $is_debug = false;

    /**
     * @param $tgt_obj
     * @param $src
     * @param $manifest
     */
    static public function upkbits(&$tgt_obj, $src, $manifest)
    {
        foreach ($manifest as $row) {
            list($n, $mask, $attr) = $row;
            $tgt_obj->$attr = ($src & $mask) >> $n;
        }
    }

    /**
     * @param $tgt_obj
     * @param $src
     * @param $manifest
     */
    static public function upkbitsL(&$tgt_obj, $src, $manifest)
    {
        foreach ($manifest as $row) {
            list($n, $mask, $attr) = $row;
            $tgt_obj->$attr = (int)(($src & $mask) >> $n);
        }
    }

    /**
     * Finds constant name by value using prefix, if provided.
     *
     * @param $val
     * @param string $prefix
     * @return bool|int|string
     */
    static public function explain_const($val, $prefix = '')
    {
        static $consts = [];
        if (!$consts) {
            foreach (get_defined_constants() as $name => $value) {
                if (!is_int($value)) continue;
                $consts[$value][] = $name;
            }
        }
        if (isset($consts[$val])) {
            foreach ($consts[$val] as $name) {
                if (substr($name, 0, strlen($prefix)) == $prefix) {
                    return $name;
                }
            }
        }
        return false;

    }

    static public function unpack_string($data, $pos, $encoding, $lenlen = 1)
    {
        # <BH
        $fmt = 'Cv';
        list($nchars) = array_values(unpack($fmt{$lenlen-1}, substr($data, $pos, $lenlen)));
        $pos += $lenlen;
//        die($encoding);
        return Helper::convert_encoding(substr($data, $pos, $nchars), 'utf-8', $encoding);
    }

    static public function unpack_string_update_pos($data, $pos, $encoding, $lenlen = 1, $known_len = Null)
    {
        if (!$known_len !== Null) {
            $nchars = $known_len;
        } else {
            # <BH
            $fmt = 'Cv';
            list($nchars) = array_values(unpack($fmt{$lenlen-1}, substr($data, $pos, $lenlen)));
            $pos += $lenlen;
        }
        $newpos = $pos + $nchars;
        return [mb_convert_encoding(substr($data, $pos, $nchars), 'utf-8', $encoding), $newpos];
    }

    static public function unpack_unicode_update_pos($data, $pos, $lenlen = 2, $known_len = Null)
    {
        if ($known_len != Null) {
            $nchars = $known_len;
        } else {
            # <BH
            $fmt = 'Cv';
            list($nchars) = array_values(unpack($fmt{$lenlen-1}, substr($data, $pos, $lenlen)));
            $pos+=$lenlen;
        }
        if (!$nchars && !substr($data, $pos)) {
            return ['', $pos];
        }
        $options = ord($data{$pos});
        $pos ++;
        $phonetic = $options & 0x04;
        $richtext = $options & 0x08;

        $rt = $sz = 0;

        if ($richtext) {
            # <H
            list($rt) = array_values(unpack('v', substr($data, $pos, $pos+2)));
            $pos += 2;
        }
        if ($phonetic) {
            # <i
            list($sz) = array_values(unpack('i', substr($data, $pos, $pos+4)));
            $pos += 4;
        }
        if ($options & 0x01) {
            # Uncompressed UTF-16-LE
            $strg = mb_convert_encoding(substr($data, $pos, 2*$nchars), 'utf-8', 'utf-16le');
            $pos += 2 * $nchars;
        } else {
            # Note: this is COMPRESSED (not ASCII!) encoding!!!
            $strg = mb_convert_encoding(substr($data, $pos, $nchars), 'utf-8', 'latin1');
            $pos += $nchars;
        }
        if ($richtext) {
            $pos += 4 * $rt;
        }
        if ($phonetic) {
            $pos += $sz;
        }
        return [$strg, $pos];

    }

    static public function unpack_unicode($data, $pos, $lenlen = 2)
    {
        # <BH
        $fmt = 'Cv';
        list($nchars) = array_values(unpack($fmt{$lenlen-1}, substr($data, $pos, $lenlen)));
        $pos += $lenlen;
        $options = ord($data[$pos]);
        $pos ++;
        # phonetic = options & 0x04
        # richtext = options & 0x08
        if ($options & 0x08) {
            # <H
            # rt = unpack('<v', data[pos:pos+2])[0] # unused
            $pos += 2;
        }
        if ($options & 0x04) {
            # <i
            # sz = unpack('<i', data[pos:pos+4])[0] # unused
            $pos += 4;
        }
        if ($options & 0x01) {
            # Uncompressed UTF-16-LE
            $rawstrg = substr($data, $pos, 2*$nchars);
            # if DEBUG: print "nchars=%d pos=%d rawstrg=%r" % (nchars, pos, rawstrg)
            $strg = mb_convert_encoding($rawstrg, 'utf-8', 'utf-16le');
            # pos += 2*nchars
        } else {
            # Note: this is COMPRESSED (not ASCII!) encoding!!!
            # Merely returning the raw bytes would work OK 99.99% of the time
            # if the local codepage was cp1252 -- however this would rapidly go pear-shaped
            # for other codepages so we grit our Anglocentric teeth and return Unicode :-)
            $strg = mb_convert_encoding(substr($data, $pos, $nchars), 'utf-8', 'latin1');
            # pos += nchars
        }
        # if richtext:
        #     pos += 4 * rt
        # if phonetic:
        #     pos += sz
        # return (strg, pos)
        return $strg;
    }

    static public function unpack_cell_range_address_list_update_pos(&$output_list, $data, $pos, $bv, $addr_size = 6)
    {
        # output_list is updated in situ
        assert($addr_size == 6 || $addr_size == 8);
        # Used to assert size == 6 if not BIFF8, but pyWLWriter writes
        # BIFF8-only MERGEDCELLS records in a BIFF5 file!
        # <H
        list($n) = array_values(unpack('v', substr($data, $pos, 2)));
        $pos += 2;
        if ($n) {
            # <HHBB | <HHHH
            $fmt = $addr_size == 6 ? 'v2a/C2b' : 'v4a';
            for ($i = 0; $i < $n; $i++) {
                list($ra, $rb, $ca, $cb) = array_values(unpack($fmt, substr($data, $pos, $addr_size)));
                $output_list[] = [$ra, $rb + 1, $ca, $cb + 1];
                $pos += $addr_size;
            }
        }
        return $pos;
    }

    /**
     * @param $rk_str
     * @return float
     */
    static public function unpack_RK($rk_str)
    {
        $flags = ord($rk_str[0]);
        if ($flags & 2) {
            # There's a SIGNED 30-bit integer in there!
            # <i
            list($i) = array_values(unpack('i', $rk_str));
            $i >>= 2; # div by 4 to drop the 2 flag bits
            if ($flags & 1) {
                return $i / 100.0;
            }
            return (float)$i;
        } else {
            # It's the most significant 30 bits of an IEEE 754 64-bit FP number
            # <d
            list($d) = array_values(unpack('d', "\x00\x00\x00\x00" . chr($flags & 252) . substr($rk_str, 1, 3)));
            if ($flags & 1) {
                return $d / 100.0;
            }
            return $d;
        }
    }

    static public function log()
    {
        if (self::$log_enabled) {
            echo call_user_func_array('sprintf', func_get_args()) . "\n";
        }
    }

    static public function debug()
    {
        if (self::$is_debug) {
            $args = func_get_args();
            if (count($args) == 1) {
                echo func_get_arg(0) . "\n";
            } else if (count($args) > 1) {
                echo call_user_func_array('sprintf', func_get_args()) . "\n";
            }
        }
    }

    static public function as_hex($data)
    {
        return join(' ', array_map(function($el){return str_pad(dechex(ord($el)), 2, '0', STR_PAD_LEFT);}, str_split($data)));
    }

    static public function get_nul_terminated_unicode($buf, $ofs)
    {
        $ofs_in = $ofs;
        # <L
        list($nb) = array_values(unpack('V', substr($buf, $ofs, 4)));
        $nb *= 2;
        $ofs += 4;
        $uc = mb_convert_encoding(substr($buf, $ofs, $nb - 1), 'utf-8', 'utf-16le');

        $ofs += $nb;
        return [$uc, $ofs];
    }

    static public function nearest_colour_index($colour_map, $rgb)
    {
        # General purpose function. Uses Euclidean distance.
        # So far used only for pre-BIFF8 WINDOW2 record.
        # Doesn't have to be fast.
        # Doesn't have to be fancy.

        $best_metric = 3 * 256 * 256;
        $best_colourx = 0;
        foreach ($colour_map as $colourx=>$cand_rgb) {
            if ($cand_rgb === Null) {
                continue;
            }
            $metric = 0;
            foreach (array_map(null, $rgb, $cand_rgb) as $v) {
                list($v1, $v2) = $v;
                $metric += ($v1 - $v2) * ($v1 - $v2);
            }
            if ($metric < $best_metric) {
                $best_metric = $metric;
                $best_colourx = $colourx;
                if ($metric == 0) {}
                break;
            }
        }
        return $best_colourx;

    }

    static public function convert_encoding($str, $to, $from)
    {
        static $use_iconv = Null;
        if (!$use_iconv === Null) {
            if (function_exists('iconv')) {
                $use_iconv = true;
            }
        }
        return $use_iconv ? iconv($from, $to, $str) : mb_convert_encoding($str, $to, $from);
    }

}

<?php

namespace laacz\XLSParser;

class CompDocException extends \Exception {};

class CompDoc
{
    public $mem = null;
    public $SAT = [];
    public $seen = [];
    public $sec_size = 0;
    public $dir_first_sec_sid = 0;
    public $min_size_std_stream;

    function __construct($data)
    {
        if (substr($data, 0, 8) !== "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1") {
            throw new CompDocException('Not an OLE2 compound document (' . Helper::as_hex(substr($data, 0, 8)) . ')!');
        }

        if (substr($data, 28, 2) !== "\xFE\xFF") {
            throw new CompDocException('Expected "little-endian" marker, found 0x' . dechex(ord($data[28])) . dechex(ord($data[29])));
        }

        $this->mem = $data;

        # <HH
        # Unused local vars
        # list($revision, $version) = array_values(unpack('v2', substr($this->mem, 24, 4)));
        # <HH
        list($ssz, $sssz) = array_values(unpack('v2', substr($this->mem, 30, 4)));

        if ($ssz > 20) $ssz = 9;
        if ($sssz > $ssz) $sssz = 6;

        $this->sec_size = $sec_size = 1 << $ssz;
        $this->short_sec_size = 1 << $sssz;

        # <iiiiiiii
        list (/*$SAT_tot_secs*/, $this->dir_first_sec_sid, , $this->min_size_std_stream,
            $SSAT_first_sec_sid, $SSAT_tot_secs,
            $MSATX_first_sec_sid, $MSATX_tot_secs
            ) = array_values(unpack('l8', substr($this->mem, 44, 32)));

        $MSATX_first_sec_sid = ($MSATX_first_sec_sid << 1) >> 1;

        $mem_data_len = strlen($this->mem) - 512;
        $mem_data_secs = (int)($mem_data_len / $sec_size);
        $left_over = $mem_data_len % $sec_size;
        if ($left_over) {
            $mem_data_secs++;
        }
        $this->mem_data_secs = $mem_data_secs;
        $this->mem_data_len = $mem_data_len;
        $this->seen = array_fill(0, $mem_data_secs, 0);

        $nent = (int)($sec_size / 4); # number of SID entries in a sector
        # <"<%di" % nent
        $fmt = "l" . $nent;
        # $trunc_warned = 0;

        # Build the MSAT
        # <109i
        $MSAT = array_values(unpack('l109', substr($this->mem, 76, 436)));

        # Both are unused in code.
        # $SAT_sectors_reqd = (int)(($mem_data_secs + $nent - 1) / $nent);
        # $expected_MSATX_sectors = max(0, (int)(($SAT_sectors_reqd - 109 + $nent - 2) / ($nent - 1)));

        $actual_MSATX_sectors = 0;
        if ($MSATX_tot_secs == 0 && in_array($MSATX_first_sec_sid, [EOCSID, FREESID, 0])) {
            # Strictly, if there is no MSAT extension, then MSATX_first_sec_sid
            # should be set to EOCSID ... FREESID and 0 have been met in the wild.
            ;# Presuming no extension
        } else {
            $sid = $MSATX_first_sec_sid;
            while (!in_array($sid, [EOCSID, FREESID])) {
                if ($sid >= $mem_data_secs) {
                    throw new CompDocException(sprintf("MSAT extension: accessing sector %d but only %d in file", $sid, $mem_data_secs));
                } elseif ($sid < 0) {
                    throw new CompDocException(sprintf("MSAT extension: invalid sector id: %d", $sid));
                }
                if ($this->seen[$sid]) {
                    throw new CompDocException(sprintf("MSAT corruption: seen[%d] == %d", $sid, $this->seen[$sid]));
                }
                $this->seen[$sid] = 1;
                $actual_MSATX_sectors ++;
                $offset = 512 + $sec_size * $sid;
                foreach (array_values(unpack($fmt, substr($this->mem, $offset, $sec_size))) as $v) {
                    $MSAT[] = $v;
                }

                $sid = array_pop($MSAT);
            }
        }

        # Build the SAT

        $this->SAT = [];
        $actual_SAT_sectors = 0;
        for ($msidx = 0; $msidx < count($MSAT); $msidx++) {
            $msid = $MSAT[$msidx];
            if (in_array($msid, [FREESID, EOCSID])) {
                continue;
            }
            if ($msid > $mem_data_secs) {
                $MSAT[$msid] = EVILSID;
                continue;
            } elseif ($msid < -2) {
                throw new CompDocException("MSAT: invalid sector id: $msid");
            }
            if ($this->seen[$msid]) {
                throw new CompDocException("MSAT extension corruption: seen[{$msid}] == {$this->seen[$msid]}");
            }
            $this->seen[$msid] = 2;
            $actual_SAT_sectors++;
            $offset = 512 + $sec_size * $msid;
            $SAT2 = array_values(unpack($fmt, substr($this->mem, $offset, $sec_size)));
            foreach ($SAT2 as $ch) $this->SAT[] = $ch;
        }
        # Build the directory

        $dbytes = $this->get_stream($this->mem, 512, $this->SAT, $this->sec_size, $this->dir_first_sec_sid, null, 'directory', 3);
        $dirlist = [];
        $did = -1;
        for ($pos = 0; $pos < strlen($dbytes); $pos += 128) {
            $did++;
            $dirlist[] = new DirNode($did, substr($dbytes, $pos, 128), 0);
        }
        $this->dirlist = $dirlist;
        $this->build_family_tree($dirlist, 0, $dirlist[0]->root_DID); # and stand well back ...

        # Get the SSCS
        $sscs_dir = $this->dirlist[0];
        assert($sscs_dir->etype == 5); # root entry

        if ($sscs_dir->first_SID < 0 || $sscs_dir->tot_size == 0) {
            $this->SSCS = '';
        } else {
            $this->SSCS = $this->get_stream($this->mem, 512, $this->SAT, $sec_size, $sscs_dir->first_SID, $sscs_dir->tot_size, NULL, "SSCS", 4);
        }

        # Build the SSAT
        $this->SSAT = [];
        if ($sscs_dir->tot_size > 0) {
            $sid = $SSAT_first_sec_sid;
            $nsecs = $SSAT_tot_secs;
            while ($sid >= 0 && $nsecs > 0) {
                if ($this->seen[$sid]) {
                    throw new CompDocException("SSAT corruption: seen[{$sid}] == {$this->seen[$sid]}");
                }
                $this->seen[$sid] = 5;
                $nsecs--;
                $start_pos = 512 + $sid * $sec_size;
                $news = unpack($fmt, substr($data, $start_pos, $sec_size));
                foreach ($news as $v) $this->SSAT[] = $v;
                $sid = $this->SAT[$sid];
            }
        }
    }

    public function get_stream($mem, $base, $sat, $sec_size, $start_sid, $size = null, $name = '', $seen_id = null)
    {
        $sectors = '';
        $s = $start_sid;
        if ($size === null) {
            while ($s >= 0) {
                if ($seen_id !== null) {
                    if ($this->seen[$s]) {
                        throw new CompDocException("%name corruption: seen[{$s}] == {%this->seen[$s]}");
                    }
                    $this->seen[$s] = $seen_id;
                }
                $start_pos = $base + $s * $sec_size;
                $sectors .= substr($mem, $start_pos, $sec_size);
                if (!isset($sat[$s])) {
                    throw new CompDocException("OLE2 stream: $name: sector allocation table invalid entry ($s)");
                }
                $s = $sat[$s];
            }
            assert($s == EOCSID);
        } else {
            $todo = $size;
            while ($s >= 0) {
                if ($seen_id !== null) {
                    if ($this->seen[$s]) {
                        throw new CompDocException("%name corruption: seen[{$s}] == {%this->seen[$s]}");
                    }
                    $this->seen[$s] = $seen_id;
                }
                $start_pos = $base + $s * $sec_size;
                $grab = $sec_size;
                if ($grab > $todo) {
                    $grab = $todo;
                }
                $todo -= $grab;
                $sectors .= substr($mem, $start_pos, $grab);
                if (!isset($sat[$s])) {
                    throw new CompDocException("OLE2 stream: $name: sector allocation table invalid entry ($s)");
                }
                $s = $sat[$s];
            }
            assert($s == EOCSID);
        }
        return $sectors;
    }

    public function build_family_tree(&$dirlist, $parent_DID, $child_DID)
    {
        if ($child_DID < 0) return;
        $this->build_family_tree($dirlist, $parent_DID, $dirlist[$child_DID]->left_DID);
        $dirlist[$parent_DID]->children[] = $child_DID;
        $dirlist[$child_DID]->parent = $parent_DID;
        $this->build_family_tree($dirlist, $parent_DID, $dirlist[$child_DID]->right_DID);
        if ($dirlist[$child_DID]->etype == 1) {
            $this->build_family_tree($dirlist, $child_DID, $dirlist[$child_DID]->root_DID);
        }
    }

    /**
     * @param $path
     * @param int $storage_DID
     * @return DirNode|null
     * @throws CompDocException
     */
    public function dir_search($path, $storage_DID = 0)
    {
        # Return matching DirNode instance or None
        $head = $path[0];
        $tail = array_slice($path, 1);
        $dl = $this->dirlist;
        $ret = null;
        foreach ($dl[$storage_DID]->children as $child) {
            if (strtolower($dl[$child]->name) == strtolower($head)) {
                $et = $dl[$child]->etype;
                if ($et == 2) {
                    $ret = $dl[$child];
                    break;
                }
                if ($et == 1) {
                    if (!$tail) {
                        throw new CompDocException("Requested component is a 'storage'");
                    }
                    $ret = $this->dir_search($tail, $child);
                    break;
                }
                throw new CompDocException("Requested stream is not a 'user stream'");
            }
        }
        return $ret;
    }

    public function get_named_stream($qname)
    {
        $d = $this->dir_search(explode('/', $qname));
        if ($d === null) {
            return null;
        }
        if ($d->tot_size >= $this->min_size_std_stream) {
            return $this->get_stream($this->mem, 512, $this->SAT, $this->sec_size, $d->first_SID, $d->tot_size, $qname, $d->DID + 6);
        } else {
            return $this->get_stream($this->SSCS, 0, $this->SSAT, $this->short_sec_size, $d->first_SID, $d->tot_size, "$qname *from SSCS)", null);
        }
    }

    public function locate_named_stream($qname)
    {
        $d = $this->dir_search(explode('/', $qname));
        if ($d == null) {
            return [null, 0, 0];
        }

        if ($d->tot_size > $this->mem_data_len) {
            throw new CompDocException("$qname stream length ({$d->tot_size}) > file data size ({$this->mem_data_len})");
        }

        if ($d->tot_size >= $this->min_size_std_stream) {
            $result = $this->locate_stream($this->mem, 512, $this->SAT, $this->sec_size, $d->first_SID, $d->tot_size, $qname, $d->DID + 6);
        } else {
            $result = [$this->get_stream($this->SSCS, 0, $this->SSAT, $this->short_sec_size, $d->first_SID, $d->tot_size, "$qname (from SSCS)", null),
            0,
            $d->tot_size];
//            print_r($result);
        }

        return $result;

    }

    function locate_stream($data, $base, $sat, $sec_size, $start_sid, $expected_stream_size, $qname, $seen_id)
    {
        $s = $start_sid;
        if ($s < 0) {
            throw new CompDocException("locate_Stream: start_sid ($s) is -ve");
        }

        $p = -99; # dummy previous SID
        $start_pos = -9999;
        $end_pos = -8888;
        $slices = [];
        $tot_found = 0;
        $found_limit = (int)(($expected_stream_size + $sec_size - 1) / $sec_size);

        while ($s >= 0) {
            if ($this->seen[$s]) {
                throw new CompDocException("$qname corruption: seen[{$s}] == {$this->seen[$s]}");
            }
            $this->seen[$s] = $seen_id;
            $tot_found++;
            if ($tot_found > $found_limit) {
                throw new CompDocException("%qname: size exceeds expected " . ($found_limit * $sec_size) . " bytes; corrupt?");
            }
            if ($s == $p + 1) {
                $end_pos += $sec_size;
            } else {
                if ($p > 0) {
                    $slices[] = [$start_pos, $end_pos];
                }
                $start_pos = $base + $s * $sec_size;
                $end_pos = $start_pos + $sec_size;
            }
            $p = $s;
            $s = $sat[$s];
        }

        assert($s == EOCSID);
        assert($tot_found == $found_limit);
        if (!$slices) {
            return [$data, $start_pos, $expected_stream_size];
        }
        $slices[] = [$start_pos, $end_pos];
        $ret = '';
        foreach ($slices as $slice) {
            $ret .= substr($data, $slice[0], $slice[1] - $slice[0]);
        }
        return [$ret, 0, $expected_stream_size];

    }


}

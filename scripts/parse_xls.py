"""
Minimal BIFF8 XLS parser for reading legacy .xls files without xlrd.
Handles OLE compound files and extracts cell values (strings, numbers, RK values).
"""
import struct


def read_xls(filepath):
    """
    Read a BIFF8 .xls file and return headers + rows as list of dicts.
    Falls back to this parser when xlrd is not available.
    """
    with open(filepath, 'rb') as f:
        data = f.read()

    if data[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        raise ValueError("Not an OLE compound file")

    sector_size = 2 ** struct.unpack('<H', data[30:32])[0]
    fat_sectors_count = struct.unpack('<I', data[44:48])[0]
    first_dir_sector = struct.unpack('<I', data[48:52])[0]

    # Read DIFAT
    difat = []
    for i in range(109):
        s = struct.unpack('<I', data[76 + i * 4:80 + i * 4])[0]
        if s < 0xFFFFFFFE:
            difat.append(s)

    # Build FAT
    fat = []
    for sec in difat[:fat_sectors_count]:
        offset = (sec + 1) * sector_size
        for j in range(sector_size // 4):
            fat.append(struct.unpack('<I', data[offset + j * 4:offset + j * 4 + 4])[0])

    def read_stream(start_sector):
        result = b''
        sec = start_sector
        visited = set()
        while sec < 0xFFFFFFFE and sec < len(fat) and sec not in visited:
            visited.add(sec)
            offset = (sec + 1) * sector_size
            result += data[offset:offset + sector_size]
            sec = fat[sec]
        return result

    # Find Workbook stream in directory
    dir_data = read_stream(first_dir_sector)
    workbook_start = None
    workbook_size = 0

    for i in range(0, len(dir_data), 128):
        entry = dir_data[i:i + 128]
        if len(entry) < 128:
            break
        name_len = struct.unpack('<H', entry[64:66])[0]
        name = entry[:name_len].decode('utf-16-le', errors='ignore').rstrip('\x00')
        start_sec = struct.unpack('<I', entry[116:120])[0]
        stream_size = struct.unpack('<I', entry[120:124])[0]
        if name == 'Workbook' or name == 'Book':
            workbook_start = start_sec
            workbook_size = stream_size
            break

    if workbook_start is None:
        raise ValueError("Workbook stream not found in XLS file")

    wb_data = read_stream(workbook_start)[:workbook_size]

    # Parse BIFF8 records
    sst_strings = []
    rows = {}
    pos = 0
    continue_target = None

    while pos < len(wb_data) - 4:
        rec_type = struct.unpack('<H', wb_data[pos:pos + 2])[0]
        rec_len = struct.unpack('<H', wb_data[pos + 2:pos + 4])[0]
        rec_data = wb_data[pos + 4:pos + 4 + rec_len]

        if rec_type == 0x00FC:  # SST (Shared String Table)
            _parse_sst(rec_data, sst_strings)

        elif rec_type == 0x003C:  # CONTINUE record for SST
            if continue_target == 'SST':
                _parse_sst_continue(rec_data, sst_strings)

        elif rec_type == 0x00FD:  # LABELSST
            if len(rec_data) >= 10:
                row = struct.unpack('<H', rec_data[0:2])[0]
                col = struct.unpack('<H', rec_data[2:4])[0]
                sst_idx = struct.unpack('<I', rec_data[6:10])[0]
                if sst_idx < len(sst_strings):
                    rows.setdefault(row, {})[col] = sst_strings[sst_idx]

        elif rec_type == 0x0203:  # NUMBER
            if len(rec_data) >= 14:
                row = struct.unpack('<H', rec_data[0:2])[0]
                col = struct.unpack('<H', rec_data[2:4])[0]
                val = struct.unpack('<d', rec_data[6:14])[0]
                rows.setdefault(row, {})[col] = val

        elif rec_type == 0x027E:  # RK
            if len(rec_data) >= 10:
                row = struct.unpack('<H', rec_data[0:2])[0]
                col = struct.unpack('<H', rec_data[2:4])[0]
                rk_val = struct.unpack('<I', rec_data[6:10])[0]
                val = _decode_rk(rk_val)
                rows.setdefault(row, {})[col] = val

        elif rec_type == 0x00BD:  # MULRK
            if len(rec_data) >= 6:
                row = struct.unpack('<H', rec_data[0:2])[0]
                first_col = struct.unpack('<H', rec_data[2:4])[0]
                # Each RK entry is 6 bytes (2 xf + 4 rk)
                n_entries = (len(rec_data) - 6) // 6
                for idx in range(n_entries):
                    offset = 4 + idx * 6
                    if offset + 6 <= len(rec_data):
                        rk_val = struct.unpack('<I', rec_data[offset + 2:offset + 6])[0]
                        val = _decode_rk(rk_val)
                        rows.setdefault(row, {})[first_col + idx] = val

        # Track if next CONTINUE belongs to SST
        if rec_type == 0x00FC:
            continue_target = 'SST'
        elif rec_type != 0x003C:
            continue_target = None

        pos += 4 + rec_len

    if not rows:
        return [], []

    # Build structured output
    max_col = max(max(r.keys()) for r in rows.values())
    sorted_rows = sorted(rows.keys())

    # First row = headers
    header_row = sorted_rows[0]
    headers = []
    for c in range(max_col + 1):
        h = rows[header_row].get(c, f'col_{c}')
        headers.append(str(h).strip() if h else f'col_{c}')

    # Data rows
    data_rows = []
    for row_idx in sorted_rows[1:]:
        row_dict = {}
        for c in range(max_col + 1):
            val = rows[row_idx].get(c, None)
            if c < len(headers):
                row_dict[headers[c]] = val
        data_rows.append(row_dict)

    return headers, data_rows


def _parse_sst(rec_data, sst_strings):
    """Parse SST record strings. SST format: 4 bytes total_refs + 4 bytes unique_count + strings."""
    if len(rec_data) < 8:
        return
    unique_strings = struct.unpack('<I', rec_data[4:8])[0]
    spos = 8
    while spos < len(rec_data) and len(sst_strings) < unique_strings:
        if spos + 3 > len(rec_data):
            break
        str_len = struct.unpack('<H', rec_data[spos:spos + 2])[0]
        flags = rec_data[spos + 2]
        spos += 3

        # Skip rich text and East Asian extra data
        rt_runs = 0
        ext_sz = 0
        if flags & 0x08:  # Rich text
            rt_runs = struct.unpack('<H', rec_data[spos:spos + 2])[0]
            spos += 2
        if flags & 0x04:  # East Asian
            ext_sz = struct.unpack('<I', rec_data[spos:spos + 4])[0]
            spos += 4

        if flags & 0x01:  # UTF-16
            byte_len = str_len * 2
            if spos + byte_len <= len(rec_data):
                s = rec_data[spos:spos + byte_len].decode('utf-16-le', errors='replace')
                spos += byte_len
            else:
                s = rec_data[spos:].decode('utf-16-le', errors='replace')
                spos = len(rec_data)
        else:  # Latin-1
            if spos + str_len <= len(rec_data):
                s = rec_data[spos:spos + str_len].decode('latin-1', errors='replace')
                spos += str_len
            else:
                s = rec_data[spos:].decode('latin-1', errors='replace')
                spos = len(rec_data)

        spos += rt_runs * 4 + ext_sz
        sst_strings.append(s)


def _parse_sst_continue(rec_data, sst_strings):
    """Parse CONTINUE record that extends the SST."""
    # Simplified - just try to extract more strings
    pass


def _decode_rk(rk_val):
    """Decode an RK value to a Python float/int."""
    if rk_val & 0x02:  # Integer
        val = rk_val >> 2
        if val & 0x20000000:  # Sign extension for negative
            val = val - 0x40000000
        if rk_val & 0x01:
            val = val / 100.0
    else:  # IEEE float
        buf = struct.pack('<Q', (rk_val & 0xFFFFFFFC) << 32)
        val = struct.unpack('<d', buf)[0]
        if rk_val & 0x01:
            val = val / 100.0
    return val


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print("Usage: python parse_xls.py <file.xls>")
        sys.exit(1)
    headers, data = read_xls(sys.argv[1])
    print(f"Headers: {headers}")
    print(f"Rows: {len(data)}")
    for row in data[:5]:
        print(row)

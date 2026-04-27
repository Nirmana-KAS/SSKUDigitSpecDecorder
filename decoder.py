def decode_codes(codes, stone_lookup, polish_lookup, shape_lookup, origin_lookup):
    results = []
    skipped = []

    for raw_code in codes:
        code = str(raw_code).strip()
        if not code:
            continue

        # Check if code ends with ROUGH
        if code.upper().endswith('ROUGH'):
            stone_code = code[:4].upper()
            if stone_code not in stone_lookup:
                skipped.append({
                    'code': code,
                    'reason': 'Stone code not found in library',
                })
                continue
            results.append({
                'code': code,
                'stone_name': stone_lookup[stone_code],
                'polishing': '',
                'shape': 'ROUGH',
                'colour': '',
                'origin': '',
            })
            continue

        # Standard 11-digit code
        if len(code) < 11:
            skipped.append({
                'code': code,
                'reason': 'Invalid code length (must be 11 digits or end with ROUGH)',
            })
            continue

        stone_code = code[0:4].upper()
        polish_code = code[4:5].upper()
        shape_code = code[5:7].upper()
        colour_code = code[7:9]
        origin_code = code[9:11].upper()

        # Skip if stone not in library
        if stone_code not in stone_lookup:
            skipped.append({
                'code': code,
                'reason': 'Stone code not found in library',
            })
            continue

        stone_name = stone_lookup.get(stone_code, '')
        polishing = polish_lookup.get(polish_code, polish_code)
        shape = shape_lookup.get(shape_code, shape_code)
        colour = colour_code
        origin = origin_lookup.get(origin_code, origin_code)

        results.append({
            'code': code,
            'stone_name': stone_name,
            'polishing': polishing,
            'shape': shape,
            'colour': colour,
            'origin': origin,
        })

    return results, skipped

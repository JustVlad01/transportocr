#!/usr/bin/env python3
"""
Test script to verify Cork and Northern Ireland route variants are working correctly
"""

import re

def generate_route_variants(route_label):
    """Return normalized variants for matching, with aliases and zeros handled.
    Examples:
      - 'Dublin 001' -> dublin 001/dublin001/dublin 1/dublin1
      - 'Northern Ireland 1' -> northern ireland 1/ni 1/northern ireland one/ni one and packed/space variants
      - 'Cork 1' -> cork 1/cork1/cork one/corkone
    """
    try:
        label = (route_label or "").strip()
        lower_label = label.lower()

        # Define canonical bases and their aliases
        base_aliases = {
            "dublin": ["dublin"],
            "northern ireland": ["northern ireland", "ni"],
            "cork": ["cork"],
        }

        # Define number to word mappings for Cork routes
        cork_number_words = {
            1: "one", 2: "two", 3: "three", 4: "four", 5: "five",
            6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten",
            11: "eleven", 12: "twelve", 13: "thirteen", 14: "fourteen",
            15: "fifteen", 16: "sixteen"
        }

        # Define number to word mappings for Northern Ireland routes
        ni_number_words = {
            1: "one", 2: "two", 3: "three", 4: "four", 5: "five",
            6: "six", 7: "seven", 8: "eight", 9: "nine", 10: "ten",
            11: "eleven", 12: "twelve", 13: "thirteen", 14: "fourteen",
            15: "fifteen", 16: "sixteen", 17: "seventeen", 18: "eighteen",
            19: "nineteen", 20: "twenty", 21: "twenty-one", 22: "twenty-two"
        }

        detected_num = None
        matched_aliases = None
        is_cork_route = False
        is_ni_route = False

        for canonical, aliases in base_aliases.items():
            for alias in aliases:
                pattern = rf"(?i)\b{re.escape(alias)}\b\s*0*(\d{{1,3}})\b"
                m = re.search(pattern, lower_label, flags=re.IGNORECASE)
                if m:
                    detected_num = int(m.group(1))
                    matched_aliases = aliases
                    is_cork_route = (canonical == "cork")
                    is_ni_route = (canonical == "northern ireland")
                    break
                
                # If this is a Cork route, also try to match word patterns (e.g., "Cork One")
                if canonical == "cork" and detected_num is None:
                    for num, word in cork_number_words.items():
                        word_pattern = rf"(?i)\b{re.escape(alias)}\b\s+{re.escape(word)}\b"
                        if re.search(word_pattern, lower_label, flags=re.IGNORECASE):
                            detected_num = num
                            matched_aliases = aliases
                            is_cork_route = True
                            break
                
                # If this is a Northern Ireland route, also try to match word patterns (e.g., "Northern Ireland One")
                if canonical == "northern ireland" and detected_num is None:
                    for num, word in ni_number_words.items():
                        word_pattern = rf"(?i)\b{re.escape(alias)}\b\s+{re.escape(word)}\b"
                        if re.search(word_pattern, lower_label, flags=re.IGNORECASE):
                            detected_num = num
                            matched_aliases = aliases
                            is_ni_route = True
                            break
                
                if detected_num is not None:
                    break

        if detected_num is None:
            base = lower_label
            return [base, base.replace(" ", "")]

        num_padded = f"{detected_num:03d}"
        num_unpadded = str(detected_num)

        variants = []
        for alias in matched_aliases:
            alias_l = alias.lower()
            variants.extend([
                f"{alias_l} {num_padded}",
                f"{alias_l}{num_padded}",
                f"{alias_l} {num_unpadded}",
                f"{alias_l}{num_unpadded}",
            ])

        # Add Cork word variants if this is a Cork route
        if is_cork_route and detected_num in cork_number_words:
            word_variant = cork_number_words[detected_num]
            variants.extend([
                f"cork {word_variant}",
                f"cork{word_variant}",
            ])

        # Add Northern Ireland word variants if this is an NI route
        if is_ni_route and detected_num in ni_number_words:
            word_variant = ni_number_words[detected_num]
            variants.extend([
                f"northern ireland {word_variant}",
                f"northern ireland{word_variant}",
                f"ni {word_variant}",
                f"ni{word_variant}",
            ])

        # unique, lowered
        lowered = []
        for v in variants:
            v = v.lower()
            if v not in lowered:
                lowered.append(v)
        return lowered
    except Exception:
        base = (route_label or "").lower()
        return [base, base.replace(" ", "")]

def test_cork_routes():
    """Test Cork route variants"""
    print("Testing Cork route variants...")
    print("=" * 50)
    
    # Test Cork numbered routes
    test_routes = [
        "Cork 1", "Cork 2", "Cork 3", "Cork 4", "Cork 5",
        "Cork 6", "Cork 7", "Cork 8", "Cork 9", "Cork 10",
        "Cork 11", "Cork 12", "Cork 13", "Cork 14", "Cork 15", "Cork 16"
    ]
    
    for route in test_routes:
        variants = generate_route_variants(route)
        print(f"\n{route}:")
        for variant in variants:
            print(f"  - {variant}")

def test_ni_routes():
    """Test Northern Ireland route variants"""
    print("\n" + "=" * 50)
    print("Testing Northern Ireland route variants...")
    
    # Test NI numbered routes
    test_routes = [
        "NI 1", "NI 2", "NI 3", "NI 4", "NI 5",
        "NI 6", "NI 7", "NI 8", "NI 9", "NI 10",
        "NI 11", "NI 12", "NI 13", "NI 14", "NI 15", "NI 16",
        "NI 17", "NI 18", "NI 19", "NI 20", "NI 21", "NI 22"
    ]
    
    for route in test_routes:
        variants = generate_route_variants(route)
        print(f"\n{route}:")
        for variant in variants:
            print(f"  - {variant}")

def test_ni_word_routes():
    """Test Northern Ireland word-based routes"""
    print("\n" + "=" * 50)
    print("Testing Northern Ireland word-based routes...")
    
    # Test NI word routes
    word_routes = [
        "Northern Ireland One", "Northern Ireland Two", "Northern Ireland Three",
        "Northern Ireland Four", "Northern Ireland Five", "Northern Ireland Six",
        "Northern Ireland Seven", "Northern Ireland Eight", "Northern Ireland Nine",
        "Northern Ireland Ten", "Northern Ireland Eleven", "Northern Ireland Twelve"
    ]
    
    for route in word_routes:
        variants = generate_route_variants(route)
        print(f"\n{route}:")
        for variant in variants:
            print(f"  - {variant}")

def test_matching():
    """Test that routes with different formats generate matching variants"""
    print("\n" + "=" * 50)
    print("Testing route format matching...")
    
    # Test Cork matching
    cork1_variants = generate_route_variants("Cork 1")
    corkone_variants = generate_route_variants("Cork One")
    
    print(f"\nCork 1 variants: {cork1_variants}")
    print(f"Cork One variants: {corkone_variants}")
    
    cork_overlap = set(cork1_variants) & set(corkone_variants)
    if cork_overlap:
        print(f"✓ Cork routes share variants: {cork_overlap}")
    else:
        print("✗ Cork routes share no variants")
    
    # Test NI matching
    ni1_variants = generate_route_variants("NI 1")
    ni001_variants = generate_route_variants("NI 001")
    ni_one_variants = generate_route_variants("NI One")
    
    print(f"\nNI 1 variants: {ni1_variants}")
    print(f"NI 001 variants: {ni001_variants}")
    print(f"NI One variants: {ni_one_variants}")
    
    ni_overlap = set(ni1_variants) & set(ni001_variants) & set(ni_one_variants)
    if ni_overlap:
        print(f"✓ NI routes share variants: {ni_overlap}")
    else:
        print("✗ NI routes share no variants")

if __name__ == "__main__":
    test_cork_routes()
    test_ni_routes()
    test_ni_word_routes()
    test_matching()

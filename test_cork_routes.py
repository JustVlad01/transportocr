#!/usr/bin/env python3
"""
Test script to verify Cork route variants are working correctly
"""

import re

def generate_route_variants(route_label):
    """Return normalized variants for matching, with aliases and zeros handled.
    Examples:
      - 'Dublin 001' -> dublin 001/dublin001/dublin 1/dublin1
      - 'Northern Ireland 1' -> northern ireland 1/ni 1 and packed/space variants
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

        detected_num = None
        matched_aliases = None
        is_cork_route = False

        for canonical, aliases in base_aliases.items():
            for alias in aliases:
                # First try to match numeric patterns (e.g., "Cork 1", "Cork 001")
                pattern = rf"(?i)\b{re.escape(alias)}\b\s*0*(\d{{1,3}})\b"
                m = re.search(pattern, lower_label, flags=re.IGNORECASE)
                if m:
                    detected_num = int(m.group(1))
                    matched_aliases = aliases
                    is_cork_route = (canonical == "cork")
                    break
                
                # If this is a Cork route, also try to match word patterns (e.g., "Cork One")
                if canonical == "cork":
                    for num, word in cork_number_words.items():
                        word_pattern = rf"(?i)\b{re.escape(alias)}\b\s+{re.escape(word)}\b"
                        if re.search(word_pattern, lower_label, flags=re.IGNORECASE):
                            detected_num = num
                            matched_aliases = aliases
                            is_cork_route = True
                            break
                    if detected_num is not None:
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
    
    print("\n" + "=" * 50)
    print("Testing word-based Cork routes...")
    
    # Test Cork word routes
    word_routes = [
        "Cork One", "Cork Two", "Cork Three", "Cork Four", "Cork Five",
        "Cork Six", "Cork Seven", "Cork Eight", "Cork Nine", "Cork Ten",
        "Cork Eleven", "Cork Twelve", "Cork Thirteen", "Cork Fourteen",
        "Cork Fifteen", "Cork Sixteen"
    ]
    
    for route in word_routes:
        variants = generate_route_variants(route)
        print(f"\n{route}:")
        for variant in variants:
            print(f"  - {variant}")

def test_matching():
    """Test that Cork 1 and Cork One generate the same variants"""
    print("\n" + "=" * 50)
    print("Testing that Cork 1 and Cork One generate matching variants...")
    
    cork1_variants = generate_route_variants("Cork 1")
    corkone_variants = generate_route_variants("Cork One")
    
    print(f"\nCork 1 variants: {cork1_variants}")
    print(f"Cork One variants: {corkone_variants}")
    
    # Check for overlap
    overlap = set(cork1_variants) & set(corkone_variants)
    if overlap:
        print(f"\n✓ Found matching variants: {overlap}")
    else:
        print("\n✗ No matching variants found")

if __name__ == "__main__":
    test_cork_routes()
    test_matching()

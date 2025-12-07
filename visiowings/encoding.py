"""Encoding detection and handling for VBA files.

This module provides functions to detect the appropriate encoding for VBA files
based on Visio's language settings or system defaults.

VBA files exported by Office applications use the system's ANSI codepage, which
varies by locale:
- Western Europe/US: cp1252
- Russia/Cyrillic: cp1251
- Central Europe: cp1250
- etc.
"""

# Default fallback codepage (Western European).
# This codepage is used as a fallback when document language detection fails
# or when the LCID (Locale ID) is not present in the mapping table below.
# cp1252 is the Windows ANSI codepage for Western European languages,
# and is the default for most English and Western European locales.
DEFAULT_CODEPAGE = 'cp1252'

# Based on Windows NLS (National Language Support) defaults
""" 
LCID_TO_CODEPAGE maps Windows Locale IDs (LCIDs) to their default ANSI codepages.

Structure:
    dict[int, str]: LCID (int) → codepage name (str, e.g., 'cp1252')

Purpose:
    Used to determine the correct encoding for VBA files exported by Office applications,
    which use the system's ANSI codepage based on locale.

Data source:
    Based on Windows National Language Support (NLS) defaults.
    See: https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
"""
# Mapping of Windows Locale IDs (LCID) to ANSI codepages
# Based on Windows NLS (National Language Support) defaults
LCID_TO_CODEPAGE = {
    # Western European (cp1252)
    1033: 'cp1252',  # en-US
    2057: 'cp1252',  # en-GB
    3081: 'cp1252',  # en-AU
    4105: 'cp1252',  # en-CA
    1031: 'cp1252',  # de-DE
    2055: 'cp1252',  # de-CH
    3079: 'cp1252',  # de-AT
    1036: 'cp1252',  # fr-FR
    2060: 'cp1252',  # fr-BE
    3084: 'cp1252',  # fr-CA
    4108: 'cp1252',  # fr-CH
    1040: 'cp1252',  # it-IT
    1034: 'cp1252',  # es-ES
    2058: 'cp1252',  # es-MX
    1043: 'cp1252',  # nl-NL
    2067: 'cp1252',  # nl-BE
    1046: 'cp1252',  # pt-BR
    2070: 'cp1252',  # pt-PT
    1053: 'cp1252',  # sv-SE
    1030: 'cp1252',  # da-DK
    1044: 'cp1252',  # nb-NO
    2068: 'cp1252',  # nn-NO
    1035: 'cp1252',  # fi-FI
    1039: 'cp1252',  # is-IS
    1027: 'cp1252',  # ca-ES
    1069: 'cp1252',  # eu-ES (Basque)
    1110: 'cp1252',  # gl-ES (Galician)
    
    # Central European (cp1250)
    1045: 'cp1250',  # pl-PL
    1029: 'cp1250',  # cs-CZ
    1038: 'cp1250',  # hu-HU
    1051: 'cp1250',  # sk-SK
    1060: 'cp1250',  # sl-SI
    1050: 'cp1250',  # hr-HR
    2074: 'cp1250',  # sr-Latn-CS
    1048: 'cp1250',  # ro-RO
    1052: 'cp1250',  # sq-AL (Albanian)
    
    # Cyrillic (cp1251)
    1049: 'cp1251',  # ru-RU
    1058: 'cp1251',  # uk-UA
    1059: 'cp1251',  # be-BY
    1026: 'cp1251',  # bg-BG
    3098: 'cp1251',  # sr-Cyrl-CS
    1071: 'cp1251',  # mk-MK
    1087: 'cp1251',  # kk-KZ
    2092: 'cp1251',  # az-Cyrl-AZ
    2115: 'cp1251',  # uz-Cyrl-UZ
    1064: 'cp1251',  # tg-Cyrl-TJ
    1088: 'cp1251',  # ky-KG
    1092: 'cp1251',  # tt-RU
    1104: 'cp1251',  # mn-MN
    
    # Greek (cp1253)
    1032: 'cp1253',  # el-GR
    
    # Turkish (cp1254)
    1055: 'cp1254',  # tr-TR
    1068: 'cp1254',  # az-Latn-AZ
    
    # Hebrew (cp1255)
    1037: 'cp1255',  # he-IL
    
    # Arabic (cp1256)
    1025: 'cp1256',  # ar-SA
    5121: 'cp1256',  # ar-DZ
    15361: 'cp1256', # ar-BH
    3073: 'cp1256',  # ar-EG
    2049: 'cp1256',  # ar-IQ
    11265: 'cp1256', # ar-JO
    13313: 'cp1256', # ar-KW
    12289: 'cp1256', # ar-LB
    4097: 'cp1256',  # ar-LY
    6145: 'cp1256',  # ar-MA
    8193: 'cp1256',  # ar-OM
    16385: 'cp1256', # ar-QA
    10241: 'cp1256', # ar-SY
    7169: 'cp1256',  # ar-TN
    14337: 'cp1256', # ar-AE
    9217: 'cp1256',  # ar-YE
    1065: 'cp1256',  # fa-IR (Persian/Farsi)
    1056: 'cp1256',  # ur-PK (Urdu)
    
    # Baltic (cp1257)
    1063: 'cp1257',  # lt-LT
    1062: 'cp1257',  # lv-LV
    1061: 'cp1257',  # et-EE
    
    # Vietnamese (cp1258)
    1066: 'cp1258',  # vi-VN
    
    # Thai (cp874)
    1054: 'cp874',   # th-TH
    
    # Japanese (cp932 / Shift-JIS)
    1041: 'cp932',   # ja-JP
    
    # Simplified Chinese (cp936 / GBK)
    2052: 'cp936',   # zh-CN
    4100: 'cp936',   # zh-SG
    
    # Korean (cp949)
    1042: 'cp949',   # ko-KR
    
    # Traditional Chinese (cp950 / Big5)
    1028: 'cp950',   # zh-TW
    3076: 'cp950',   # zh-HK
    5124: 'cp950',   # zh-MO
}

def get_encoding_from_document(document, debug=False):
    """Detect the appropriate encoding from Visio document's language property.
    
    Uses Document.Language which returns the language ID recorded in the 
    document's VERSIONINFO resource - i.e., the language the document was
    created with.
    
    Args:
        document: Visio Document COM object
        debug: If True, print debug information
        
    Returns:
        Encoding string (e.g., 'cp1251', 'cp1252') or None if detection fails.
    """
    try:
        lcid = document.Language
        if debug:
            print(f"[DEBUG] Document Language LCID: {lcid}")
        
        encoding = LCID_TO_CODEPAGE.get(lcid)
        if encoding and debug:
            print(f"[DEBUG] Detected encoding from document language: {encoding}")
        return encoding
    except Exception as e:
        if debug:
            print(f"[DEBUG] Could not get Document.Language: {e}")
    
    return None



def resolve_encoding(document, user_codepage=None, debug=False):
    """Resolve the encoding to use for VBA files.
    
    Priority order:
    1. User-specified encoding (if provided)
    2. Auto-detected from Document's language property
    3. System default encoding
    
    Args:
        document: Visio Document COM object
        user_codepage: User-specified encoding string (optional)
        debug: If True, print debug information
        
    Returns:
        Encoding string (e.g., 'cp1251', 'cp1252').
    """
    # If user specified an encoding, use it
    if user_codepage:
        if debug:
            print(f"[DEBUG] Using user-specified encoding: {user_codepage}")
        return user_codepage
    
    # Auto-detect from document language
    if document:
        detected = get_encoding_from_document(document, debug)
        if detected:
            return detected
    
    # Fall back to system default
    return DEFAULT_CODEPAGE

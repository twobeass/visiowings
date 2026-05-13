# Codepages

Visio stores VBA module text in the system's ANSI codepage. The codepage
in use depends on the locale the document was created in. visiowings
detects the right codepage automatically from `Document.Language`, but
you can override it with `--codepage`.

## Locale matrix

| Codepage | Region / scripts                                  | Sample LCIDs                          |
| -------- | ------------------------------------------------- | ------------------------------------- |
| `cp1252` | Western European, English, German, French, …     | 1033 (en-US), 1031 (de-DE), 1036 (fr-FR) |
| `cp1250` | Central European: Polish, Czech, Hungarian, …    | 1045, 1029, 1038                      |
| `cp1251` | Cyrillic: Russian, Ukrainian, Bulgarian, …       | 1049, 1058, 1026                      |
| `cp1253` | Greek                                             | 1032                                  |
| `cp1254` | Turkish                                           | 1055                                  |
| `cp1255` | Hebrew                                            | 1037                                  |
| `cp1256` | Arabic, Persian, Urdu                             | 1025, 1065, 1056                      |
| `cp1257` | Baltic: Lithuanian, Latvian, Estonian             | 1063, 1062, 1061                      |
| `cp1258` | Vietnamese                                        | 1066                                  |
| `cp874`  | Thai                                              | 1054                                  |
| `cp932`  | Japanese (Shift-JIS)                              | 1041                                  |
| `cp936`  | Simplified Chinese (GBK)                          | 2052                                  |
| `cp949`  | Korean                                            | 1042                                  |
| `cp950`  | Traditional Chinese (Big5)                        | 1028                                  |

The full mapping lives in `visiowings/encoding.py`. Submit a PR if your
locale is missing — the lookup table is the only thing that needs
updating.

## Auto-detection

```python
# visiowings.encoding.resolve_encoding pseudocode
encoding = (
    user_codepage
    or LCID_TO_CODEPAGE.get(document.Language)
    or "cp1252"  # safe default
)
```

## Manual override

```bash
visiowings edit --file Привет.vsdm --codepage cp1251
```

## BOMs

If you save a `.bas` file with a UTF-8 / UTF-16 BOM (some editors do
this on Windows), visiowings detects and strips it before re-importing
into Visio. The BOM bytes do not pollute your VBA module name.

## Round-trip guarantees

The test suite parametrises every supported codepage and verifies that
encode → bytes → decode produces identical text for representative
samples. See `tests/test_encoding_roundtrip.py`.

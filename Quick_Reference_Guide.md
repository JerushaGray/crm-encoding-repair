# Quick Reference: Most Common Encoding Issues

## Top 20 Patterns by Language

### German 🇩🇪
| Corrupted | Fixed | Example Names |
|-----------|-------|---------------|
| Ã¼ | ü | Müller, Günther |
| Ã¶ | ö | Schröder, Böhm |
| Ã¤ | ä | Bäcker, Schäfer |
| ÃŸ | ß | Straße, Groß |
| Ã„ | Ä | Ärzte, Äpfel |
| Ã– | Ö | Österreich |
| Ãœ | Ü | Über, Tür |

### French 🇫🇷
| Corrupted | Fixed | Example Names |
|-----------|-------|---------------|
| Ã© | é | René, Café |
| Ã¨ | è | Père, Système |
| Ãª | ê | Tête, Forêt |
| Ã§ | ç | François, Garçon |
| Ã  | à | À, Voilà |
| Ã´ | ô | Côte, Hôtel |
| Ã« | ë | Noël, Citroën |

### Spanish 🇪🇸
| Corrupted | Fixed | Example Names |
|-----------|-------|---------------|
| Ã± | ñ | Señor, España |
| Ã¡ | á | García, Martínez |
| Ã© | é | José, Pérez |
| Ã­ | í | María, Díaz |
| Ã³ | ó | López, Gómez |
| Ãº | ú | Raúl, Perú |
| Â¿ | ¿ | ¿Cómo estás? |
| Â¡ | ¡ | ¡Hola! |

### Polish 🇵🇱
| Corrupted | Fixed | Example Names |
|-----------|-------|---------------|
| Å‚ | ł | Kowalski, Wałęsa |
| Ä… | ą | Dąbrowski |
| Ä™ | ę | Będziński |
| Ã³ | ó | Wróbel, Kraków |
| Ä‡ | ć | Jaśko |
| Å„ | ń | Gdańsk |
| Å› | ś | Śląsk |
| Åº | ź | Źrebak |
| Å¼ | ż | Żabka |

### Swedish/Norwegian 🇸🇪 🇳🇴
| Corrupted | Fixed | Example Names |
|-----------|-------|---------------|
| Ã¥ | å | Håkan, Malmö |
| Ã¤ | ä | Mäkinen, Täby |
| Ã¶ | ö | Lindström, Örebro |
| Ã¸ | ø | København, Strømstad |
| Ã˜ | Ø | Østergaard |
| Ã… | Å | Åsa, Ångström |
| Ã† | Æ | Ærø |

### Czech/Slovak 🇨🇿 🇸🇰
| Corrupted | Fixed | Example Names |
|-----------|-------|---------------|
| Ä | č | Dvořák, Čech |
| Å¡ | š | Jakubíšek |
| Å™ | ř | Jiří, Příbram |
| Å¾ | ž | Žižkov |
| Ã½ | ý | Nový |
| Ã¡ | á | Bratislava |
| Ã© | é | René |

## Common Punctuation Issues

### Quotes
| Corrupted | Fixed | Usage |
|-----------|-------|-------|
| â€œ | " | Left double quote |
| â€ | " | Right double quote |
| â€˜ | ' | Left single quote |
| â€™ | ' | Right single quote/apostrophe |

### Dashes
| Corrupted | Fixed | Usage |
|-----------|-------|-------|
| â€" | – | En dash (ranges) |
| â€" | — | Em dash (breaks) |
| â€• | ― | Horizontal bar |

### Other Symbols
| Corrupted | Fixed | Usage |
|-----------|-------|-------|
| â€¦ | … | Ellipsis |
| â€¢ | • | Bullet point |
| â‚¬ | € | Euro symbol |
| Â© | © | Copyright |
| Â® | ® | Registered |
| â„¢ | ™ | Trademark |
| Â° | ° | Degree |

## HTML Entities
| Corrupted | Fixed | Usage |
|-----------|-------|-------|
| &#39; | ' | Apostrophe |
| &quot; | " | Quote |
| &amp; | & | Ampersand |
| &nbsp; | (space) | Non-breaking space |

## Recognition Patterns

### How to Spot Encoding Issues:
1. **Ã followed by special characters** → Usually accented letters
2. **â€ followed by anything** → Usually punctuation or quotes
3. **Å or Ä followed by special chars** → Usually Eastern European
4. **Multiple special chars where one should be** → Encoding problem

### Examples:
- `Ã©` = 2 characters that should be 1 → `é`
- `â€™` = 3 characters that should be 1 → `'`
- `ÃƒÂ©` = 4 characters that should be 1 → `é` (double-encoded)

## Quick Test

### Is This an Encoding Issue?
✅ YES if you see:
- Multiple weird characters where an accent should be
- � (replacement character)
- Patterns like Ã+special char
- â€+anything
- Names that look "garbled" but you can guess what they should be

❌ NO if you see:
- Random unrelated characters
- Numbers in place of letters
- Complete gibberish with no pattern

## Language Coverage

The script handles these language families:

**Western European:**
- German, French, Spanish, Portuguese, Italian, Dutch

**Nordic:**
- Swedish, Norwegian, Danish, Icelandic, Finnish

**Eastern European:**
- Polish, Czech, Slovak, Romanian, Hungarian, Croatian

**Baltic:**
- Latvian, Lithuanian, Estonian

**Other:**
- Turkish, Albanian

## When to Run

Run the script after:
- ✅ Importing CSV files
- ✅ Exporting from Salesforce/HubSpot/Dynamics
- ✅ Receiving data from international offices
- ✅ Migrating between systems
- ✅ Copy/paste from emails or web

## Pro Tips

1. **Always keep the log file** - it's your audit trail
2. **Run on a copy first** if you're nervous
3. **Check the summary** - if it says 0 fixes, your data was already clean
4. **Look for patterns** - if many names from one country are broken, they'll all fix the same way
5. **Share the script** - your international colleagues will love you

---

**Need help?** Check the full documentation or send me new patterns you encounter!

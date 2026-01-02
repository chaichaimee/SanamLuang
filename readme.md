# SanamLuang

<p align="center">
  <img src="https://www.nvaccess.org/files/nvda/documentation/userGuide/images/nvda.ico" alt="NVDA Logo" width="120">
  <br><br>
  <strong>SanamLuang</strong><br>
  <em>An NVDA add-on to correct common Thai text reading errors</em>
</p>

<p align="center">
  <strong>Author:</strong> Chai Chaimee<br>
  <strong>Repository:</strong> <a href="https://github.com/chaichaimee/SanamLuang">https://github.com/chaichaimee/SanamLuang</a>
</p>

---

## Description

**SanamLuang** is an NVDA add-on designed to fix frequent Thai text reading errors caused by:

- Incorrect spelling commonly used by younger generations
- Voice input (speech-to-text) mistakes
- Auto-correction / word suggestion errors
- Wrong positioning of vowel signs, tone marks, and other diacritics from OCR scanning

These issues often cause speech synthesizers to pronounce Thai text incorrectly.

## Hotkeys

- **NVDA+Shift+F4** → Open SanamLuang settings window  
- **NVDA+Shift+F4** (with text selected or entire document highlighted) → Automatically correct all errors in the selected text  
- **Double-press NVDA+Shift+F4 quickly** → Toggle error reading mode on/off

## How It Works

After installation, SanamLuang works immediately by detecting and reading misplaced/incorrect vowel signs and tone marks based on its built-in initial dictionary.

You can add your own custom misspelled or incorrectly positioned words via the settings window (opened with **NVDA+Shift+F4**), similar to NVDA's built-in speech dictionary.

To disable SanamLuang's error reading, double-press **NVDA+Shift+F4** quickly. You will hear “SanamLuang off” and the reading will revert to standard NVDA behavior.

### For OCR-processed documents

When working with scanned/OCR documents (which often contain misplaced vowels/tone marks such as ํา ํ่า ํ้า เเ etc.), simply:

1. Select all text (Ctrl+A)
2. Press **NVDA+Shift+F4**

SanamLuang will automatically scan and correct all defined vowel/tone mark positioning errors throughout the document.  
Only entries present in SanamLuang's dictionary will be corrected — you can expand the dictionary yourself.

> **Note:** This saves you from manually fixing each occurrence one by one.

## Additional Benefits

Besides regular document correction, SanamLuang also helps normalize reading of intentionally altered Thai words commonly found on social media (to evade content filtering), such as:

- เงิu → เงิน  
- เvมร / เขมs → เขมร  
- and many other evasion-style spellings

This greatly improves speech synthesizer accuracy when dealing with increasingly distorted and creative forms of written Thai across online platforms.

---

## Correction Dictionary Format (JSONL)

SanamLuang uses a simple **JSONL** file format for its correction dictionary:

```json
{"word": [["corrected form"]]}

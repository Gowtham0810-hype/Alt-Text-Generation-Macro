# 📝 Alt Text Generation Macro

## 📌 Project Overview

**Alt Text Generation Macro** is a VBA-based Word macro designed to automatically generate descriptive alternative (alt) text for all images embedded in a Microsoft Word document. It leverages the **Groq Vision API** to analyze and describe images in natural language, making documents more accessible and SEO-friendly.

---

## 🚀 Features

- ✅ Scans all images (`InlineShapes` and `Shapes`) in the Word document.
- ✅ Extracts image content as base64.
- ✅ Sends image data to Groq Vision API.
- ✅ Receives and inserts descriptive alt text for each image.
- ✅ Enhances accessibility and compliance with web standards (e.g., WCAG).

---

## 🧰 Technologies Used

- **VBA (Visual Basic for Applications)**
- **Groq Vision API**
- **Microsoft Word (Office VBA)**

---

## 📦 Prerequisites

- Microsoft Word (with macro support)
- Groq API Key with access to Vision model
- Internet access for API calls
- Developer access to the Word VBA Editor

---

## 🔧 Setup Instructions

### 1. Get a Groq API Key
Sign up and obtain your API key from [https://groq.com](https://groq.com).

### 2. Open Word VBA Editor
- Open Microsoft Word.
- Press `ALT + F11` to open the VBA editor.

### 3. Add the Macro Code
In the VBA editor:
- Insert a new module (`Insert > Module`)
- Paste the macro code from the file `AltTextMacro.bas` (see below for snippet)

### 4. Set Your API Key
Replace the placeholder `YOUR_GROQ_API_KEY` in the code with your actual Groq API key.

---

## 📄 Example Use

1. Open your Word document.
2. Run the macro `GenerateAltTextForImagesWithGroq`.
3. Watch as alt text is automatically applied to each image.
4. Save the document with enhanced accessibility.

---

## 🛡️ Disclaimer

- Image data is sent to a third-party service (Groq API); do not use with confidential documents.
- This macro is a proof-of-concept and not production-hardened.

---

## 📬 Contact

For questions, contributions, or suggestions, feel free to open an issue or contact the maintainer.

